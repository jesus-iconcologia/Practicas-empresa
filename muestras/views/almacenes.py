import base64
import json
import io
import os
import zipfile

import openpyxl
import pandas as pd
from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required, permission_required
from django.db import connection, transaction
from django.db.models import Count, Prefetch, Q
from django.http import FileResponse, HttpResponse, StreamingHttpResponse
from django.shortcuts import redirect, render
from django.template import loader
from django.utils.text import slugify
from openpyxl.styles import Font, PatternFill

from ..forms import (
    CongeladorForm,
    CrearRecursoCongeladorForm,
    ReparacionRecursoForm,
    UploadExcel,
)
from ..models import Bandeja, Caja, Congelador, Estante, Rack, ReparacionRecurso, Subposicion
from ..parameters_config import get_excel_colors, get_upload_messages
from .common import _json_progress, _should_update


ESTRUCTURA_LABELS = dict(Congelador.ESTRUCTURA_CHOICES)
BOX_LAYOUTS = {
    '5x5': {'label': '5x5', 'rows': list('ABCDE'), 'columns': [str(i) for i in range(1, 6)]},
    '9x9': {'label': '9x9', 'rows': list('ABCDEFGHI'), 'columns': [str(i) for i in range(1, 10)]},
    '10x10': {'label': '10x10', 'rows': list('ABCDEFGHIJ'), 'columns': [str(i) for i in range(1, 11)]},
    '8x12': {'label': '8x12', 'rows': list('ABCDEFGH'), 'columns': [str(i) for i in range(1, 13)]},
}


def _get_congelador_importacion(request):
    """Recupera el recurso objetivo de una importación si llega por GET o POST."""
    congelador_id = request.GET.get('congelador_id') or request.POST.get('congelador_id')
    if not congelador_id:
        return None
    try:
        return Congelador.objects.get(pk=congelador_id)
    except (Congelador.DoesNotExist, ValueError, TypeError):
        return None


def _usa_estructura_simple(congelador):
    """Indica si el recurso usa una jerarquía reducida sin bandejas ni subposiciones."""
    return bool(congelador and congelador.tipo_estructura in {
        Congelador.ESTRUCTURA_CONGELADOR_20,
        Congelador.ESTRUCTURA_NEVERA,
    })


def _get_recurso_por_nombre(nombre_congelador):
    return Congelador.objects.filter(congelador=nombre_congelador).first()


def _layout_capacity(layout_key):
    layout = BOX_LAYOUTS.get(layout_key)
    if not layout:
        return 0
    return len(layout['rows']) * len(layout['columns'])


def _iter_layout_positions(layout_key):
    layout = BOX_LAYOUTS.get(layout_key)
    if not layout:
        raise ValueError('Formato de caja no soportado.')
    for row in layout['rows']:
        for column in layout['columns']:
            yield row, column


def _describe_subposicion_conflicto(subposicion):
    muestra_label = subposicion.muestra.nom_lab if subposicion.muestra_id else 'una muestra'
    caja = subposicion.caja
    rack = caja.rack
    estante = rack.estante
    recurso = estante.congelador
    rack_label = 'Cajón' if _usa_estructura_simple(recurso) else 'Rack'
    parts = [
        f'Estante {estante.numero}',
        f'{rack_label} {rack.numero}',
    ]
    if caja.bandeja_id:
        parts.append(f'Bandeja {caja.bandeja.numero}')
    parts.extend([
        f'Caja {caja.numero}',
        f'Subposición {subposicion.numero}',
    ])
    return muestra_label, ' > '.join(parts)


def _build_delete_conflict_message(label, nombre, subposicion, *, article='el'):
    muestra_label, ruta = _describe_subposicion_conflicto(subposicion)
    return (
        f'No se puede eliminar {article} {label} "{nombre}" porque contiene muestras almacenadas. '
        f'Por ejemplo, la muestra "{muestra_label}" está en {ruta}.'
    )


def _get_or_create_rack(estante, nombre, posicion):
    rack, created = Rack.objects.get_or_create(
        estante=estante,
        numero=nombre,
        defaults={'posicion_rack_estante': str(posicion)}
    )
    if not created and str(rack.posicion_rack_estante) != str(posicion):
        return None, f'El rack/cajón "{nombre}" ya existe en otra posición del estante.'
    if Rack.objects.filter(estante=estante, posicion_rack_estante=str(posicion)).exclude(pk=rack.pk).exists():
        return None, f'La posición {posicion} del estante ya está ocupada por otro rack/cajón.'
    return rack, None


def _get_or_create_bandeja(rack, nombre, posicion):
    bandeja, created = Bandeja.objects.get_or_create(
        rack=rack,
        numero=nombre,
        defaults={'posicion_bandeja_rack': str(posicion)}
    )
    if not created and str(bandeja.posicion_bandeja_rack) != str(posicion):
        return None, f'La bandeja "{nombre}" ya existe en otra posición del rack.'
    if Bandeja.objects.filter(rack=rack, posicion_bandeja_rack=str(posicion)).exclude(pk=bandeja.pk).exists():
        return None, f'La posición {posicion} de bandeja ya está ocupada.'
    return bandeja, None


def _get_or_create_caja(rack, bandeja, nombre, posicion):
    defaults = {'posicion_caja_rack': str(posicion) if bandeja else ''}
    caja, created = Caja.objects.get_or_create(
        rack=rack,
        bandeja=bandeja,
        numero=nombre,
        defaults=defaults
    )
    if bandeja:
        if not created and str(caja.posicion_caja_rack) != str(posicion):
            return None, f'La caja "{nombre}" ya existe en otra posición de la bandeja.'
        if Caja.objects.filter(bandeja=bandeja, posicion_caja_rack=str(posicion)).exclude(pk=caja.pk).exists():
            return None, f'La posición {posicion} de caja ya está ocupada.'
    return caja, None


def _build_structure_capacity(estantes):
    total = 0
    for estante in estantes:
        for rack in estante.get('racks', []):
            if 'bandejas' in rack:
                for bandeja in rack.get('bandejas', []):
                    for caja in bandeja.get('cajas', []):
                        total += _layout_capacity(caja.get('layout'))
            else:
                for caja in rack.get('cajas', []):
                    total += _layout_capacity(caja.get('layout'))
    return total


def _build_import_preview(recurso, filas_validas, simple_structure):
    preview = {
        'congelador': recurso.congelador if recurso else (filas_validas[0]['congelador'] if filas_validas else ''),
        'tipo_estructura': recurso.tipo_estructura if recurso else None,
        'simple_structure': simple_structure,
        'estantes': [],
        'summary_total_posiciones': len(filas_validas),
        'summary_ocupadas': 0,
        'summary_libres': len(filas_validas),
        'summary_cajas': 0,
    }

    estantes_map = {}
    cajas_keys = set()

    for fila in filas_validas:
        estante_key = str(fila['estante'])
        estante_data = estantes_map.get(estante_key)
        if not estante_data:
            estante_data = {
                'numero': fila['estante'],
                'racks': [],
                '_racks_map': {},
            }
            estantes_map[estante_key] = estante_data
            preview['estantes'].append(estante_data)

        rack_key = (str(fila['rack']), str(fila['posicion_rack_estante']))
        rack_data = estante_data['_racks_map'].get(rack_key)
        if not rack_data:
            rack_data = {
                'numero': fila['rack'],
                'posicion_rack_estante': fila['posicion_rack_estante'],
                'numero_posiciones': 0,
                'numero_cajas': 0,
                '_cajas_keys': set(),
            }
            if simple_structure:
                rack_data['cajas'] = []
                rack_data['_cajas_map'] = {}
            else:
                rack_data['bandejas'] = []
                rack_data['_bandejas_map'] = {}
            estante_data['_racks_map'][rack_key] = rack_data
            estante_data['racks'].append(rack_data)

        rack_data['numero_posiciones'] += 1

        if simple_structure:
            caja_key = str(fila['caja'])
            caja_data = rack_data['_cajas_map'].get(caja_key)
            if not caja_data:
                caja_data = {
                    'numero': fila['caja'],
                    'numero_posiciones': 0,
                }
                rack_data['_cajas_map'][caja_key] = caja_data
                rack_data['cajas'].append(caja_data)
                rack_data['numero_cajas'] += 1
            caja_data['numero_posiciones'] += 1
            cajas_keys.add((estante_key, rack_key, caja_key))
            continue

        bandeja_key = (str(fila['bandeja']), str(fila['posicion_bandeja_rack']))
        bandeja_data = rack_data['_bandejas_map'].get(bandeja_key)
        if not bandeja_data:
            bandeja_data = {
                'numero': fila['bandeja'],
                'posicion_bandeja_rack': fila['posicion_bandeja_rack'],
                'cajas': [],
                'numero_posiciones': 0,
                'numero_cajas': 0,
                '_cajas_map': {},
            }
            rack_data['_bandejas_map'][bandeja_key] = bandeja_data
            rack_data['bandejas'].append(bandeja_data)

        bandeja_data['numero_posiciones'] += 1
        caja_key = (str(fila['caja']), str(fila['posicion_caja_rack']))
        caja_data = bandeja_data['_cajas_map'].get(caja_key)
        if not caja_data:
            caja_data = {
                'numero': fila['caja'],
                'posicion_caja_rack': fila['posicion_caja_rack'],
                'numero_posiciones': 0,
            }
            bandeja_data['_cajas_map'][caja_key] = caja_data
            bandeja_data['cajas'].append(caja_data)
            bandeja_data['numero_cajas'] += 1
            rack_data['numero_cajas'] += 1
        caja_data['numero_posiciones'] += 1
        cajas_keys.add((estante_key, rack_key, bandeja_key, caja_key))

    preview['summary_cajas'] = len(cajas_keys)

    for estante_data in preview['estantes']:
        estante_data.pop('_racks_map', None)
        for rack_data in estante_data['racks']:
            if simple_structure:
                rack_data.pop('_cajas_map', None)
            else:
                for bandeja_data in rack_data['bandejas']:
                    bandeja_data.pop('_cajas_map', None)
                rack_data.pop('_bandejas_map', None)
            rack_data.pop('_cajas_keys', None)

    return preview


def _infer_box_layout(caja):
    rows = sorted({sub.fila for sub in caja.subposiciones.all() if sub.fila})
    columns = sorted({sub.columna for sub in caja.subposiciones.all() if sub.columna}, key=lambda value: (len(str(value)), str(value)))
    total = caja.subposiciones.count()

    for key, layout in BOX_LAYOUTS.items():
        if len(rows) == len(layout['rows']) and len(columns) == len(layout['columns']) and total == _layout_capacity(key):
            return key
    return '9x9'


def _serialize_existing_structure(recurso):
    estantes = (
        Estante.objects
        .filter(congelador=recurso)
        .prefetch_related(
            Prefetch(
                'racks',
                queryset=Rack.objects.prefetch_related(
                    Prefetch(
                        'bandejas',
                        queryset=Bandeja.objects.prefetch_related(
                            Prefetch('cajas', queryset=Caja.objects.prefetch_related('subposiciones').order_by('posicion_caja_rack', 'numero'))
                        ).order_by('posicion_bandeja_rack', 'numero')
                    ),
                    Prefetch('cajas', queryset=Caja.objects.filter(bandeja__isnull=True).prefetch_related('subposiciones').order_by('numero')),
                ).order_by('posicion_rack_estante', 'numero')
            )
        )
        .order_by('numero')
    )

    data = {'estantes': []}
    for estante in estantes:
        estante_data = {
            'id': f'estante-{estante.id}',
            'name': estante.numero,
            'existing': True,
            'racks': [],
        }
        for rack in estante.racks.all():
            rack_data = {
                'id': f'rack-{rack.id}',
                'name': rack.numero,
                'existing': True,
            }
            if _usa_estructura_simple(recurso):
                rack_data['cajas'] = [
                    {
                        'id': f'caja-{caja.id}',
                        'name': caja.numero,
                        'layout': _infer_box_layout(caja),
                        'existing': True,
                    }
                    for caja in rack.cajas.all()
                ]
            else:
                rack_data['bandejas'] = []
                for bandeja in rack.bandejas.all():
                    bandeja_data = {
                        'id': f'bandeja-{bandeja.id}',
                        'name': bandeja.numero,
                        'existing': True,
                        'cajas': [
                            {
                                'id': f'caja-{caja.id}',
                                'name': caja.numero,
                                'layout': _infer_box_layout(caja),
                                'existing': True,
                            }
                            for caja in bandeja.cajas.all()
                        ],
                    }
                    rack_data['bandejas'].append(bandeja_data)
            estante_data['racks'].append(rack_data)
        data['estantes'].append(estante_data)
    return data


def _validate_estructura_manual(recurso, structure_data):
    estantes = structure_data.get('estantes', [])
    if not estantes:
        return 'Debe añadir al menos un estante.'

    nombres_estantes = set()
    simple_structure = _usa_estructura_simple(recurso)

    for estante_data in estantes:
        nombre_estante = (estante_data.get('name') or '').strip()
        if not nombre_estante:
            return 'Todos los estantes deben tener nombre.'

        nombre_estante_key = nombre_estante.casefold()
        if nombre_estante_key in nombres_estantes:
            return 'No puede haber dos estantes con el mismo nombre.'
        nombres_estantes.add(nombre_estante_key)

        nombres_racks = set()
        for rack_data in estante_data.get('racks', []):
            nombre_rack = (rack_data.get('name') or '').strip()
            if not nombre_rack:
                return f'Cada rack/cajón del estante "{nombre_estante}" debe tener nombre.'

            nombre_rack_key = nombre_rack.casefold()
            if nombre_rack_key in nombres_racks:
                return f'No puede haber dos {"cajones" if simple_structure else "racks"} con el mismo nombre en el estante "{nombre_estante}".'
            nombres_racks.add(nombre_rack_key)

            if simple_structure:
                nombres_cajas = set()
                for caja_data in rack_data.get('cajas', []):
                    nombre_caja = (caja_data.get('name') or '').strip()
                    if not nombre_caja:
                        return f'Cada caja del cajón "{nombre_rack}" debe tener nombre.'

                    nombre_caja_key = nombre_caja.casefold()
                    if nombre_caja_key in nombres_cajas:
                        return f'No puede haber dos cajas con el mismo nombre en el cajón "{nombre_rack}".'
                    nombres_cajas.add(nombre_caja_key)

                    if caja_data.get('layout') not in BOX_LAYOUTS:
                        return f'La caja "{nombre_caja}" tiene un formato no soportado.'
                continue

            nombres_bandejas = set()
            for bandeja_data in rack_data.get('bandejas', []):
                nombre_bandeja = (bandeja_data.get('name') or '').strip()
                if not nombre_bandeja:
                    return f'Cada bandeja del rack "{nombre_rack}" debe tener nombre.'

                nombre_bandeja_key = nombre_bandeja.casefold()
                if nombre_bandeja_key in nombres_bandejas:
                    return f'No puede haber dos bandejas con el mismo nombre en el rack "{nombre_rack}".'
                nombres_bandejas.add(nombre_bandeja_key)

                nombres_cajas = set()
                for caja_data in bandeja_data.get('cajas', []):
                    nombre_caja = (caja_data.get('name') or '').strip()
                    if not nombre_caja:
                        return f'Cada caja de la bandeja "{nombre_bandeja}" debe tener nombre.'

                    nombre_caja_key = nombre_caja.casefold()
                    if nombre_caja_key in nombres_cajas:
                        return f'No puede haber dos cajas con el mismo nombre en la bandeja "{nombre_bandeja}".'
                    nombres_cajas.add(nombre_caja_key)

                    if caja_data.get('layout') not in BOX_LAYOUTS:
                        return f'La caja "{nombre_caja}" tiene un formato no soportado.'

    return None


def _crear_estructura_manual(recurso, structure_data):
    estantes = structure_data.get('estantes', [])
    validation_error = _validate_estructura_manual(recurso, structure_data)
    if validation_error:
        return validation_error

    try:
        with transaction.atomic():
            for estante_data in estantes:
                nombre_estante = (estante_data.get('name') or '').strip()
                if not nombre_estante:
                    raise ValueError('Todos los estantes deben tener nombre.')
                estante, _ = Estante.objects.get_or_create(congelador=recurso, numero=nombre_estante)

                for rack_index, rack_data in enumerate(estante_data.get('racks', []), start=1):
                    nombre_rack = (rack_data.get('name') or '').strip()
                    if not nombre_rack:
                        raise ValueError(f'Cada rack/cajón del estante "{nombre_estante}" debe tener nombre.')
                    rack, error = _get_or_create_rack(estante, nombre_rack, rack_index)
                    if error:
                        raise ValueError(error)

                    if _usa_estructura_simple(recurso):
                        cajas = rack_data.get('cajas', [])
                        for caja_index, caja_data in enumerate(cajas, start=1):
                            nombre_caja = (caja_data.get('name') or '').strip()
                            if not nombre_caja:
                                raise ValueError(f'Cada caja del cajón "{nombre_rack}" debe tener nombre.')
                            layout_key = caja_data.get('layout')
                            caja, error = _get_or_create_caja(rack, None, nombre_caja, caja_index)
                            if error:
                                raise ValueError(error)
                            if not (caja_data.get('existing') and caja.subposiciones.exists()):
                                for fila, columna in _iter_layout_positions(layout_key):
                                    Subposicion.objects.get_or_create(
                                        caja=caja,
                                        fila=fila,
                                        columna=columna,
                                        defaults={'numero': f'{fila}-{columna}'}
                                    )
                        continue

                    for bandeja_index, bandeja_data in enumerate(rack_data.get('bandejas', []), start=1):
                        nombre_bandeja = (bandeja_data.get('name') or '').strip()
                        if not nombre_bandeja:
                            raise ValueError(f'Cada bandeja del rack "{nombre_rack}" debe tener nombre.')
                        bandeja, error = _get_or_create_bandeja(rack, nombre_bandeja, bandeja_index)
                        if error:
                            raise ValueError(error)

                        for caja_index, caja_data in enumerate(bandeja_data.get('cajas', []), start=1):
                            nombre_caja = (caja_data.get('name') or '').strip()
                            if not nombre_caja:
                                raise ValueError(f'Cada caja de la bandeja "{nombre_bandeja}" debe tener nombre.')
                            layout_key = caja_data.get('layout')
                            caja, error = _get_or_create_caja(rack, bandeja, nombre_caja, caja_index)
                            if error:
                                raise ValueError(error)
                            if not (caja_data.get('existing') and caja.subposiciones.exists()):
                                for fila, columna in _iter_layout_positions(layout_key):
                                    Subposicion.objects.get_or_create(
                                        caja=caja,
                                        fila=fila,
                                        columna=columna,
                                        defaults={'numero': f'{fila}-{columna}'}
                                    )
    except ValueError as exc:
        return str(exc)
    return None


def _set_foreign_key_checks(enabled: bool):
    """
    Activa o desactiva temporalmente las comprobaciones de FK según el motor.

    La edición del nombre del recurso requiere sincronizar campos que usan
    `to_field='congelador'`, y no todos los motores soportan la misma sintaxis.
    """
    with connection.cursor() as cursor:
        if connection.vendor == 'mysql':
            cursor.execute(f"SET FOREIGN_KEY_CHECKS={1 if enabled else 0}")
        elif connection.vendor == 'sqlite':
            cursor.execute(f"PRAGMA foreign_keys = {'ON' if enabled else 'OFF'}")


def _attach_resource_summaries(congeladores):
    for congelador in congeladores:
        total_posiciones = 0
        ocupadas = 0
        cajas = 0

        for estante in congelador.estantes.all():
            for rack in estante.racks.all():
                if _usa_estructura_simple(congelador):
                    rack_cajas = list(rack.cajas.all())
                    cajas += len(rack_cajas)
                    for caja in rack_cajas:
                        total_posiciones += getattr(caja, 'numero_posiciones', 0)
                        ocupadas += getattr(caja, 'numero_muestras', 0)
                    continue

                for bandeja in rack.bandejas.all():
                    bandeja_cajas = list(bandeja.cajas.all())
                    cajas += len(bandeja_cajas)
                    for caja in bandeja_cajas:
                        total_posiciones += getattr(caja, 'numero_posiciones', 0)
                        ocupadas += getattr(caja, 'numero_muestras', 0)

        congelador.summary_total_posiciones = total_posiciones
        congelador.summary_ocupadas = ocupadas
        congelador.summary_libres = max(total_posiciones - ocupadas, 0)
        congelador.summary_cajas = cajas


def _get_upload_config(congelador):
    """Devuelve el mapeo de columnas esperado según la estructura del recurso."""
    if _usa_estructura_simple(congelador):
        return {
            'rename_columns': {
                'Recurso': 'congelador',
                'Estante': 'estante',
                'Posición del cajón en el estante': 'posicion_rack_estante',
                'Cajón': 'rack',
                'Caja': 'caja',
                'Fila': 'fila',
                'Columna': 'columna',
            },
            'simple_structure': True,
        }
    recurso_header = 'Recurso' if congelador else 'Congelador'
    return {
        'rename_columns': {
            recurso_header: 'congelador',
            'Estante': 'estante',
            'Posición del rack en el estante': 'posicion_rack_estante',
            'Rack': 'rack',
            'Posición de la bandeja en el rack': 'posicion_bandeja_rack',
            'Bandeja': 'bandeja',
            'Posición de la caja en la bandeja': 'posicion_caja_rack',
            'Caja': 'caja',
            'Fila': 'fila',
            'Columna': 'columna',
        },
        'simple_structure': False,
    }

@login_required
@permission_required('muestras.can_view_localizaciones_web')
def localizaciones(request):
    """Muestra el árbol de recursos y posiciones con filtros por nivel."""

    filtro_congelador = request.GET.get('filtro_congelador', '').strip()
    filtro_tipo = request.GET.get('filtro_tipo', '').strip()
    filtro_estante = request.GET.get('filtro_estante', '').strip()
    filtro_rack = request.GET.get('filtro_rack', '').strip()
    filtro_cajon = request.GET.get('filtro_cajon', '').strip()
    filtro_bandeja = request.GET.get('filtro_bandeja', '').strip()
    filtro_caja = request.GET.get('filtro_caja', '').strip()

    recurso_filtrado = None
    if filtro_congelador:
        recurso_filtrado = (
            Congelador.objects
            .filter(congelador=filtro_congelador)
            .only('tipo_estructura')
            .first()
        )
    recurso_simple = _usa_estructura_simple(recurso_filtrado)

    if recurso_simple:
        filtro_rack = ''
        filtro_bandeja = ''
        filtro_caja = ''
    else:
        filtro_cajon = ''

    # Construir querysets filtrados para cada nivel.
    cajas_qs = Caja.objects.annotate(
        numero_posiciones=Count('subposiciones'),
        numero_muestras=Count(
            'subposiciones',
            filter=Q(subposiciones__vacia=False)
        ),
    )
    if filtro_caja:
        cajas_qs = cajas_qs.filter(id=filtro_caja)

    bandejas_qs = Bandeja.objects.all()
    if filtro_bandeja:
        bandejas_qs = bandejas_qs.filter(id=filtro_bandeja)
    bandejas_qs = bandejas_qs.prefetch_related(
        Prefetch('cajas', queryset=cajas_qs)
    )

    racks_qs = Rack.objects.all()
    if filtro_rack:
        racks_qs = racks_qs.filter(id=filtro_rack)
    if filtro_cajon:
        racks_qs = racks_qs.filter(id=filtro_cajon)
    racks_qs = racks_qs.prefetch_related(
        Prefetch('bandejas', queryset=bandejas_qs),
        Prefetch('cajas', queryset=cajas_qs.filter(bandeja__isnull=True)),
    )

    estantes_qs = Estante.objects.all()
    if filtro_estante:
        estantes_qs = estantes_qs.filter(id=filtro_estante)
    estantes_qs = estantes_qs.prefetch_related(
        Prefetch('racks', queryset=racks_qs)
    )

    # Construir queryset de recursos con un único prefetch anidado.
    congeladores = Congelador.objects.all().prefetch_related(
        Prefetch('estantes', queryset=estantes_qs)
    )

    # Filtrar recursos asegurando que contienen los niveles hijos seleccionados.
    if filtro_congelador:
        congeladores = congeladores.filter(congelador=filtro_congelador)
    if filtro_tipo:
        congeladores = congeladores.filter(tipo_estructura=filtro_tipo)
    if filtro_estante:
        congeladores = congeladores.filter(estantes__id=filtro_estante).distinct()
    if filtro_rack:
        congeladores = congeladores.filter(estantes__racks__id=filtro_rack).distinct()
    if filtro_cajon:
        congeladores = congeladores.filter(estantes__racks__id=filtro_cajon).distinct()
    if filtro_bandeja:
        congeladores = congeladores.filter(estantes__racks__bandejas__id=filtro_bandeja).distinct()
    if filtro_caja:
        congeladores = congeladores.filter(
            Q(estantes__racks__bandejas__cajas__id=filtro_caja) |
            Q(estantes__racks__cajas__id=filtro_caja)
        ).distinct()

    template = loader.get_template('localizaciones_todas.html')

    # Opciones para filtros del frontend.
    recursos = list(
        Congelador.objects
        .all()
        .values('congelador', 'tipo_estructura')
        .distinct()
        .order_by('congelador')
    )
    resource_type_map = {
        r['congelador']: (r['tipo_estructura'] or '')
        for r in recursos
    }

    congeladores = list(congeladores)
    _attach_resource_summaries(congeladores)

    context = {
        'congeladores': congeladores,
        'opciones_congeladores': [r['congelador'] for r in recursos],
        'opciones_tipos': Congelador.ESTRUCTURA_CHOICES,
        'resource_type_map_json': json.dumps(resource_type_map),
        'recurso_simple_filtrado': recurso_simple,
        'filtro_congelador': filtro_congelador,
        'filtro_tipo': filtro_tipo,
        'filtro_estante': filtro_estante,
        'filtro_rack': filtro_rack,
        'filtro_cajon': filtro_cajon,
        'filtro_bandeja': filtro_bandeja,
        'filtro_caja': filtro_caja,
    }
    return HttpResponse(template.render(context, request))


@login_required
@permission_required('muestras.can_add_localizaciones_web')
def opciones_creacion_localizaciones(request, nombre_congelador):
    recurso = _get_recurso_por_nombre(nombre_congelador)
    if not recurso:
        messages.error(request, 'No se ha encontrado el recurso solicitado.')
        return redirect('localizaciones_todas')

    estructura_label = ESTRUCTURA_LABELS.get(
        recurso.tipo_estructura,
        recurso.get_tipo_estructura_display() if recurso.tipo_estructura else 'No definida',
    )
    return render(
        request,
        'opciones_creacion_localizaciones.html',
        {
            'congelador': recurso,
            'estructura_label': estructura_label,
            'simple_structure': _usa_estructura_simple(recurso),
        },
    )


@login_required
@permission_required('muestras.can_add_localizaciones_web')
def crear_localizacion_manual(request, nombre_congelador):
    recurso = _get_recurso_por_nombre(nombre_congelador)
    if not recurso:
        messages.error(request, 'No se ha encontrado el recurso solicitado.')
        return redirect('localizaciones_todas')

    layout_options = [
        {'value': key, 'label': f"{config['label']} ({_layout_capacity(key)} posiciones)"}
        for key, config in BOX_LAYOUTS.items()
    ]

    if request.method == 'POST':
        structure_raw = request.POST.get('structure_json', '')
        try:
            structure_data = json.loads(structure_raw) if structure_raw else {}
        except json.JSONDecodeError:
            structure_data = {}
            messages.error(request, 'La estructura enviada no es válida.')
        else:
            error = _crear_estructura_manual(recurso, structure_data)
            if error:
                messages.error(request, error)
            else:
                return redirect('detalles_congelador', nombre_congelador=recurso.congelador)
    else:
        structure_data = _serialize_existing_structure(recurso)

    return render(
        request,
        'crear_localizacion_manual.html',
        {
            'congelador': recurso,
            'simple_structure': _usa_estructura_simple(recurso),
            'layout_options': layout_options,
            'initial_structure_json': structure_data,
        },
    )


@login_required
@permission_required('muestras.can_add_localizaciones_web')
def upload_excel_localizaciones(request):
    """Valida e importa la estructura de posiciones de un recurso desde Excel."""
    congelador_context = _get_congelador_importacion(request)
    upload_config = _get_upload_config(congelador_context)
    rename_columns = upload_config['rename_columns']
    simple_structure = upload_config['simple_structure']
    estructura_label = None
    if congelador_context and congelador_context.tipo_estructura:
        estructura_label = ESTRUCTURA_LABELS.get(
            congelador_context.tipo_estructura,
            congelador_context.get_tipo_estructura_display(),
        )

    def build_upload_context(extra=None):
        context = {
            'form': UploadExcel(),
            'congelador': congelador_context,
            'estructura_label': estructura_label,
            'simple_structure': simple_structure,
        }
        if extra:
            context.update(extra)
        return context
    # Mostrar confirmación después de la validación con progreso.
    if request.GET.get('mostrar_confirmacion') and 'confirmacion_pendiente' in request.session:
        datos_conf = request.session.pop('confirmacion_pendiente')
        for msg in datos_conf.get('mensajes', []):
            getattr(messages, msg['level'])(request, msg['text'])
        context = datos_conf.get('context', {})
        if congelador_context:
            context.update({
                'congelador': congelador_context,
                'estructura_label': estructura_label,
            })
        return render(request, datos_conf['template'], context)
    if request.method=="POST":
        form = UploadExcel(request.POST, request.FILES)
        rename_columns = {
                    'Congelador': 'congelador', 
                    'Estante': 'estante',
                    'Posición del rack en el estante': 'posicion_rack_estante',
                    'Rack': 'rack',
                    'Posición de la bandeja en el rack': 'posicion_bandeja_rack',
                    'Bandeja': 'bandeja',
                    'Posición de la caja en la bandeja': 'posicion_caja_rack',
                    'Caja': 'caja',
                    'Fila': 'fila',
                    'Columna': 'columna'
        }
        rename_columns = upload_config['rename_columns']

        def limpiar_numero(valor):
            if pd.isna(valor):
                return None

            # Normalizar floats tipo 1.0 -> 1.
            if isinstance(valor, float) and valor.is_integer():
                valor = int(valor)

            valor = str(valor).strip()

            if valor == "":
                return None

            return valor

        # Normalizar campos de texto para tratar mayúsculas/minúsculas igual.
        def normalizar_texto(valor):
            v = limpiar_numero(valor)
            if v is None:
                return None
            try:
                return str(v).upper()
            except Exception:
                return str(v)

        # Si el usuario confirma, se guardan las localizaciones en base de datos.
        if 'confirmar' in request.POST:
            filas = request.session.pop('filas_validas', [])
            total = len(filas)
            redirect_destino = '/archivo/'
            if congelador_context:
                redirect_destino = f"/archivo/detalles_congelador/{congelador_context.congelador}"

            def gen_confirmar_localizaciones():
                yield _json_progress(0, total, 'start')
                try:
                    with transaction.atomic():
                        for i, fila in enumerate(filas):
                            congelador, _ = Congelador.objects.get_or_create(
                                congelador=fila['congelador']
                            )
                            estante, _ = Estante.objects.get_or_create(
                                congelador=congelador,
                                numero=fila['estante']
                            )
                            rack, _ = Rack.objects.get_or_create(
                                estante=estante,
                                numero=fila['rack'],
                                defaults={'posicion_rack_estante': fila['posicion_rack_estante']}
                            )
                            if simple_structure:
                                caja, _ = Caja.objects.get_or_create(
                                    rack=rack,
                                    bandeja=None,
                                    numero=fila['caja'],
                                    defaults={'posicion_caja_rack': ''}
                                )
                                Subposicion.objects.get_or_create(
                                    caja=caja,
                                    fila=fila['fila'],
                                    columna=fila['columna'],
                                    defaults={'numero': f"{fila['fila']}-{fila['columna']}"}
                                )
                                if _should_update(i, total):
                                    yield _json_progress(i + 1, total, 'processing')
                                continue
                            bandeja, _ = Bandeja.objects.get_or_create(
                                rack=rack,
                                numero=fila['bandeja'],
                                defaults={'posicion_bandeja_rack': fila['posicion_bandeja_rack']}
                            )
                            caja, _ = Caja.objects.get_or_create(
                                rack=rack,
                                bandeja=bandeja,
                                numero=fila['caja'],
                                defaults={'posicion_caja_rack': fila['posicion_caja_rack']}
                            )
                            Subposicion.objects.get_or_create(
                                caja=caja,
                                fila=fila['fila'],
                                columna=fila['columna'],
                                defaults={'numero': f"{fila['fila']}-{fila['columna']}"}
                            )
                            if _should_update(i, total):
                                yield _json_progress(i + 1, total, 'processing')
                    yield _json_progress(total, total, 'done', 'Localizaciones importadas correctamente', redirect_destino)
                except Exception as e:
                    yield _json_progress(0, total, 'error', str(e))

            return StreamingHttpResponse(gen_confirmar_localizaciones(), content_type='application/x-ndjson')
        # Si el usuario cancela, no se guarda ningún cambio.
        elif 'cancelar' in request.POST:
            messages.error(request, 'Las localizaciones del excel no se han añadido.')
            if congelador_context:
                return redirect('detalles_congelador', nombre_congelador=congelador_context.congelador)
            return redirect('localizaciones_todas')
        # Si se sube un Excel, se procesa y valida.
        elif 'excel_file' in request.FILES:
            # Limpiar sesión residual de cargas anteriores.
            if 'columnas_adicionales' in request.session:
                del request.session['columnas_adicionales']
            if form.is_valid():
                # Leer excel
                excel_file = request.FILES['excel_file']
                excel_bytes = excel_file.read()
                request.session['excel_file_base64']= base64.b64encode(excel_bytes).decode()
                excel_stream = io.BytesIO(excel_bytes)
                
                # Validar que sea un archivo Excel válido.
                try:
                    df = pd.read_excel(excel_stream)
                except Exception as e:
                    return render(request, 'localizacion_nueva.html', build_upload_context({'form': form, 'error': f'El archivo no es un Excel válido (.xlsx). Error: {str(e)}'}))
                
                # Validar que tenga al menos una fila de datos
                if len(df) == 0:
                    return render(request, 'localizacion_nueva.html', build_upload_context({'form': form, 'error': 'Error de formato: El archivo no presenta registros'}))
                
                # Validar que tenga todas las columnas esperadas
                columnas_esperadas = set(rename_columns.keys())
                columnas_existentes = set(df.columns)
                columnas_faltantes = columnas_esperadas - columnas_existentes
                
                if columnas_faltantes:
                    columnas_str = ', '.join(sorted(columnas_faltantes))
                    return render(request, 'localizacion_nueva.html', build_upload_context({'form': form, 'error': f'Error de formato: El archivo Excel no contiene las siguientes columnas esperadas: {columnas_str}'}))
                
                # Validar que no haya columnas adicionales
                columnas_adicionales = columnas_existentes - columnas_esperadas
                extra_columns = False
                columnas_adicionales_str = ''
                if columnas_adicionales:
                    columnas_adicionales_str = ', '.join(sorted(columnas_adicionales))
                    # Registrar en sesión el Excel original para descarga y avisos.
                    request.session['excel_file_base64'] = base64.b64encode(excel_bytes).decode()
                    request.session['excel_file_name'] = excel_file.name
                    # Guardar columnas adicionales para incluirlas en el Excel de errores.
                    request.session['columnas_adicionales'] = columnas_adicionales_str
                    extra_columns = True

                # renombrar columnas
                df.rename(columns=rename_columns, inplace=True)
                
                # Procesar y validar filas
                total_filas = len(df)

                def gen_validar_localizaciones():
                    errores = {}
                    filas_validas = []
                    numero_registros = len(df)
                    # Mapas para comprobar consistencias:
                    # - que una misma posición de caja no tenga cajas diferentes
                    # - que una misma posición de rack no tenga racks diferentes
                    pos_to_caja = {}
                    pos_rack_to_rack = {}
                    pos_bandeja_to_bandeja = {}
                    # Detectar si la misma subposición se usa más de una vez dentro del Excel.
                    subposiciones_usadas = set()

                    yield _json_progress(0, total_filas, 'start')

                    for idx, row in df.iterrows():
                        fila_numero = idx + 2
                        errores[fila_numero] = {"bloqueantes": []}
                        
                        # Limpiar y normalizar los valores
                        congelador = normalizar_texto(row['congelador'])
                        estante = limpiar_numero(row['estante'])
                        posicion_rack_estante = normalizar_texto(row['posicion_rack_estante'])
                        rack = normalizar_texto(row['rack'])
                        posicion_bandeja_rack = normalizar_texto(row.get('posicion_bandeja_rack')) if 'posicion_bandeja_rack' in row else None
                        bandeja = normalizar_texto(row.get('bandeja')) if 'bandeja' in row else None
                        posicion_caja_rack = normalizar_texto(row.get('posicion_caja_rack')) if 'posicion_caja_rack' in row else None
                        caja = normalizar_texto(row.get('caja')) if 'caja' in row else None
                        fila = normalizar_texto(row.get('fila')) if 'fila' in row else None
                        columna = normalizar_texto(row.get('columna')) if 'columna' in row else None
                        
                        if congelador_context and congelador != congelador_context.congelador.upper():
                            errores[fila_numero]["bloqueantes"].append("congelador_distinto_recurso")
                            if _should_update(idx, total_filas):
                                yield _json_progress(idx + 1, total_filas, 'processing')
                            continue
                        
                        if simple_structure:
                            campos = {
                                'congelador': congelador,
                                'estante': estante,
                                'posicion_rack_estante': posicion_rack_estante,
                                'rack': rack,
                                'caja': caja,
                                'fila': fila,
                                'columna': columna,
                            }

                            for nombre_campo, valor in campos.items():
                                if valor is None:
                                    errores[fila_numero]["bloqueantes"].append(f"campo_obligatorio_vacio:{nombre_campo}")
                                for nombre_campo_semicolon, valor_semicolon in campos.items():
                                    if valor_semicolon and isinstance(valor_semicolon, str) and ';' in valor_semicolon:
                                        errores[fila_numero]["bloqueantes"].append(f"caracter_invalido_semicolon:{nombre_campo_semicolon}")

                            if not errores[fila_numero]["bloqueantes"]:
                                try:
                                    if estante is not None and int(estante) <= 0:
                                        raise ValueError()
                                except Exception:
                                    errores[fila_numero]["bloqueantes"].append("formato_incorrecto:estante")

                                try:
                                    if posicion_rack_estante is not None and int(posicion_rack_estante) <= 0:
                                        raise ValueError()
                                except Exception:
                                    errores[fila_numero]["bloqueantes"].append("formato_incorrecto:posicion_rack_estante")

                            if errores[fila_numero]["bloqueantes"]:
                                if _should_update(idx, total_filas):
                                    yield _json_progress(idx + 1, total_filas, 'processing')
                                continue

                            pos_rack_key = (congelador, estante, posicion_rack_estante)
                            if pos_rack_key in pos_rack_to_rack and pos_rack_to_rack[pos_rack_key] != rack:
                                errores[fila_numero]["bloqueantes"].append("rack_inconsistente")
                                if _should_update(idx, total_filas):
                                    yield _json_progress(idx + 1, total_filas, 'processing')
                                continue
                            pos_rack_to_rack[pos_rack_key] = rack

                            simple_key = (congelador, estante, posicion_rack_estante, rack, caja, fila, columna)
                            if simple_key in subposiciones_usadas:
                                errores[fila_numero]["bloqueantes"].append("subposicion_duplicada_excel")
                                if _should_update(idx, total_filas):
                                    yield _json_progress(idx + 1, total_filas, 'processing')
                                continue
                            subposiciones_usadas.add(simple_key)

                            if Rack.objects.filter(
                                estante__congelador__congelador__iexact=congelador,
                                estante__numero__iexact=estante,
                                posicion_rack_estante__iexact=posicion_rack_estante
                            ).exclude(numero__iexact=rack).exists():
                                errores[fila_numero]["bloqueantes"].append("posicion_rack_ocupada")
                                if _should_update(idx, total_filas):
                                    yield _json_progress(idx + 1, total_filas, 'processing')
                                continue

                            if Subposicion.objects.filter(
                                fila__iexact=fila,
                                columna__iexact=columna,
                                caja__numero__iexact=caja,
                                caja__bandeja__isnull=True,
                                caja__rack__numero__iexact=rack,
                                caja__rack__estante__numero=estante,
                                caja__rack__estante__congelador__congelador__iexact=congelador,
                            ).exists():
                                errores[fila_numero]["bloqueantes"].append("localizacion_duplicada")
                            else:
                                filas_validas.append({
                                    'congelador': congelador,
                                    'estante': estante,
                                    'posicion_rack_estante': posicion_rack_estante,
                                    'rack': rack,
                                    'caja': caja,
                                    'fila': fila,
                                    'columna': columna,
                                })

                            if _should_update(idx, total_filas):
                                yield _json_progress(idx + 1, total_filas, 'processing')
                            continue

                        # Comprobar si hay campos vacíos.
                        campos = {
                            'congelador': congelador,
                            'estante': estante,
                            'posicion_rack_estante': posicion_rack_estante,
                            'rack': rack,
                            'posicion_bandeja_rack': posicion_bandeja_rack,
                            'bandeja': bandeja,
                            'posicion_caja_rack': posicion_caja_rack,
                            'caja': caja,
                            'fila': fila,
                            'columna': columna
                        }
                        
                        for nombre_campo, valor in campos.items():
                            if valor is None:
                                errores[fila_numero]["bloqueantes"].append(f"campo_obligatorio_vacio:{nombre_campo}")
                            # Validar que no haya punto y coma en ningún campo.
                            for nombre_campo, valor in campos.items():
                                if valor and isinstance(valor, str) and ';' in valor:
                                    errores[fila_numero]["bloqueantes"].append(f"caracter_invalido_semicolon:{nombre_campo}")
                        
                        # Validar que ciertos campos numéricos sean enteros positivos (>0).
                        if not errores[fila_numero]["bloqueantes"]:
                            try:
                                if estante is not None:
                                    if int(estante) <= 0:
                                        raise ValueError()
                            except Exception:
                                errores[fila_numero]["bloqueantes"].append("formato_incorrecto:estante")

                            try:
                                if posicion_rack_estante is not None:
                                    if int(posicion_rack_estante) <= 0:
                                        raise ValueError()
                            except Exception:
                                errores[fila_numero]["bloqueantes"].append("formato_incorrecto:posicion_rack_estante")

                            try:
                                if posicion_bandeja_rack is not None:
                                    if int(posicion_bandeja_rack) <= 0:
                                        raise ValueError()
                            except Exception:
                                errores[fila_numero]["bloqueantes"].append("formato_incorrecto:posicion_bandeja_rack")

                            try:
                                if posicion_caja_rack is not None:
                                    if int(posicion_caja_rack) <= 0:
                                        raise ValueError()
                            except Exception:
                                errores[fila_numero]["bloqueantes"].append("formato_incorrecto:posicion_caja_rack")

                        # Si hay errores bloqueantes hasta ahora, saltar validaciones posteriores
                        if errores[fila_numero]["bloqueantes"]:
                            if _should_update(idx, total_filas):
                                yield _json_progress(idx + 1, total_filas, 'processing')
                            continue

                        # Comprobar consistencia de rack: misma (congelador, estante, posicion_rack_estante)
                        # no puede mapear a racks distintos
                        pos_rack_key = (congelador, estante, posicion_rack_estante)
                        if pos_rack_key in pos_rack_to_rack:
                            if pos_rack_to_rack[pos_rack_key] != rack:
                                errores[fila_numero]["bloqueantes"].append("rack_inconsistente")
                                if _should_update(idx, total_filas):
                                    yield _json_progress(idx + 1, total_filas, 'processing')
                                continue
                        else:
                            pos_rack_to_rack[pos_rack_key] = rack

                        pos_bandeja_key = (congelador, estante, posicion_rack_estante, rack, posicion_bandeja_rack)
                        if pos_bandeja_key in pos_bandeja_to_bandeja:
                            if pos_bandeja_to_bandeja[pos_bandeja_key] != bandeja:
                                errores[fila_numero]["bloqueantes"].append("bandeja_inconsistente")
                                if _should_update(idx, total_filas):
                                    yield _json_progress(idx + 1, total_filas, 'processing')
                                continue
                        else:
                            pos_bandeja_to_bandeja[pos_bandeja_key] = bandeja

                        # Comprobar consistencia: una misma posición de caja no puede tener cajas distintas.
                        pos_key = (congelador, estante, posicion_rack_estante, rack, posicion_bandeja_rack, bandeja, posicion_caja_rack)
                        if pos_key in pos_to_caja:
                            if pos_to_caja[pos_key] != caja:
                                errores[fila_numero]["bloqueantes"].append("caja_inconsistente")
                                if _should_update(idx, total_filas):
                                    yield _json_progress(idx + 1, total_filas, 'processing')
                                continue
                        else:
                            pos_to_caja[pos_key] = caja

                        # Comprobar duplicado dentro del Excel: misma subposición completa ya usada.
                        subpos_key = (congelador, estante, posicion_rack_estante, rack, posicion_bandeja_rack, bandeja, posicion_caja_rack, caja, fila, columna)
                        if subpos_key in subposiciones_usadas:
                            errores[fila_numero]["bloqueantes"].append("subposicion_duplicada_excel")
                            if _should_update(idx, total_filas):
                                yield _json_progress(idx + 1, total_filas, 'processing')
                            continue
                        else:
                            subposiciones_usadas.add(subpos_key)
                        
                        # Comprobar si la posición del rack ya está ocupada por otro rack.
                        if Rack.objects.filter(
                            estante__congelador__congelador__iexact=congelador,
                            estante__numero__iexact=estante,
                            posicion_rack_estante__iexact=posicion_rack_estante
                        ).exclude(numero__iexact=rack).exists():
                            errores[fila_numero]["bloqueantes"].append("posicion_rack_ocupada")
                            if _should_update(idx, total_filas):
                                yield _json_progress(idx + 1, total_filas, 'processing')
                            continue
                        
                        # Comprobar si la posición de la caja ya está ocupada por otra caja.
                        if Bandeja.objects.filter(
                            rack__estante__congelador__congelador__iexact=congelador,
                            rack__estante__numero__iexact=estante,
                            rack__numero__iexact=rack,
                            posicion_bandeja_rack__iexact=posicion_bandeja_rack
                        ).exclude(numero__iexact=bandeja).exists():
                            errores[fila_numero]["bloqueantes"].append("posicion_bandeja_ocupada")
                            if _should_update(idx, total_filas):
                                yield _json_progress(idx + 1, total_filas, 'processing')
                            continue

                        if Caja.objects.filter(
                            bandeja__rack__estante__congelador__congelador__iexact=congelador,
                            bandeja__rack__estante__numero__iexact=estante,
                            bandeja__rack__numero__iexact=rack,
                            bandeja__numero__iexact=bandeja,
                            posicion_caja_rack__iexact=posicion_caja_rack
                        ).exclude(numero__iexact=caja).exists():
                            errores[fila_numero]["bloqueantes"].append("posicion_caja_ocupada")
                            if _should_update(idx, total_filas):
                                yield _json_progress(idx + 1, total_filas, 'processing')
                            continue
                        
                        # Comprobar si la localización ya existe (case-insensitive para textos).
                        if Subposicion.objects.filter(fila__iexact=fila,
                                                      columna__iexact=columna,
                                                      caja__numero__iexact=caja,
                                                      caja__bandeja__numero__iexact=bandeja,
                                                      caja__bandeja__rack__numero__iexact=rack,
                                                      caja__bandeja__rack__estante__numero=estante,
                                                      caja__bandeja__rack__estante__congelador__congelador__iexact=congelador).exists():
                            errores[fila_numero]["bloqueantes"].append("localizacion_duplicada")
                        else:
                            # Guardar fila válida.
                            filas_validas.append({
                                'congelador': congelador,
                                'estante': estante,
                                'posicion_rack_estante': posicion_rack_estante,
                                'rack': rack,
                                'posicion_bandeja_rack': posicion_bandeja_rack,
                                'bandeja': bandeja,
                                'posicion_caja_rack': posicion_caja_rack,
                                'caja': caja,
                                'fila': fila,
                                'columna': columna
                            })

                        if _should_update(idx, total_filas):
                            yield _json_progress(idx + 1, total_filas, 'processing')

                    # Guardar en sesión los resultados de la validación.
                    request.session['filas_validas'] = filas_validas
                    request.session['errores'] = errores

                    # Obtener configuración de mensajes para localizaciones.
                    msg_config = get_upload_messages('localizaciones')

                    # Build messages
                    mensajes = [{'level': 'info', 'text': f'{msg_config["titulo_inicial"]} {numero_registros} registros'}]

                    # Contar errores
                    numero_errores_bloqueantes = sum(1 for fila in errores if errores[fila]['bloqueantes'])

                    # Determinar si hay errores para mostrar la sección de Excel de errores.
                    errores_encontrados = (numero_errores_bloqueantes > 0) or extra_columns

                    # Generar mensajes según el estado.
                    if numero_errores_bloqueantes > 0:
                        msg = msg_config['con_bloqueantes'].format(count=numero_errores_bloqueantes)
                        mensajes.append({'level': 'error', 'text': msg})

                    if extra_columns:
                        num_extras = len([c.strip() for c in columnas_adicionales_str.split(',') if c.strip()])
                        msg = msg_config['columnas_extras'].format(count=num_extras, detalles=columnas_adicionales_str)
                        mensajes.append({'level': 'warning', 'text': msg})

                    if numero_errores_bloqueantes == 0 and not extra_columns:
                        mensajes.append({'level': 'success', 'text': msg_config['sin_errores']})

                    preview_context = {}
                    if numero_errores_bloqueantes == 0 and not extra_columns and congelador_context:
                        preview_context['preview_recurso'] = _build_import_preview(
                            congelador_context,
                            filas_validas,
                            simple_structure,
                        )

                    request.session['confirmacion_pendiente'] = {
                        'template': 'confirmacion_upload_localizacion.html',
                        'context': {
                            'errores_encontrados': errores_encontrados,
                            **preview_context,
                        },
                        'mensajes': mensajes
                    }
                    request.session.save()
                    redirect_confirmacion = request.path + '?mostrar_confirmacion=1'
                    if congelador_context:
                        redirect_confirmacion += f'&congelador_id={congelador_context.id}'
                    yield _json_progress(total_filas, total_filas, 'done', 'Validación completada', redirect_confirmacion)

                return StreamingHttpResponse(gen_validar_localizaciones(), content_type='application/x-ndjson')
            
        # Si se solicita un Excel de errores, se rellena con los fallos detectados.
        elif 'excel_errores' in request.POST:
            errores = request.session.get('errores', {})
            excel_bytes = base64.b64decode(request.session.get('excel_file_base64'))
            excel_file = io.BytesIO(excel_bytes)
            wb = openpyxl.load_workbook(excel_file)
            ws = wb.active
            
            # Definir los estilos para pintar el Excel usando configuración centralizada.
            colors = get_excel_colors()
            FILL_ERROR_CELL = PatternFill("solid", fgColor=colors['error_cell'])
            FILL_ERROR_COL = PatternFill("solid", fgColor=colors['extra_column'])
            
            # Diccionario de mensajes
            MENSAJES_ERROR = {
                "campo_obligatorio_vacio": "Campo obligatorio vacío",
                "localizacion_duplicada": "La localización ya existe en la base de datos",
                "subposicion_duplicada_excel": "La subposición aparece duplicada en el Excel",
                "formato_incorrecto": "Formato incorrecto (debe ser entero positivo)",
                "fecha_invalida": "Fecha inválida (Formato correcto: DD-MM-AAAA)",
                "caja_inconsistente": "Conflicto de caja en la misma posición",
                "rack_inconsistente": "Conflicto de rack en la misma posición",
                "bandeja_inconsistente": "Conflicto de bandeja en la misma posición",
                "posicion_rack_ocupada": "La posición del rack ya está ocupada por otro rack",
                "posicion_bandeja_ocupada": "La posición de la bandeja ya está ocupada por otra bandeja",
                "posicion_caja_ocupada": "La posición de la caja ya está ocupada por otra caja",
                "caracter_invalido_semicolon": "El carácter ';' no está permitido en este campo",
                "congelador_distinto_recurso": "El congelador indicado en el Excel no coincide con el recurso creado"
            }
            
            # Diccionario de columnas del excel
            columnas_excel = {}
            for cell in ws[1]:
                if cell.value in rename_columns:
                    columnas_excel[rename_columns[cell.value]] = cell.column
            
            # Añadir la columna de errores.
            col_errores = ws.max_column + 1
            ws.cell(row=1, column=col_errores, value="Errores")
            
            # Comprobar si hay columnas adicionales registradas en sesión.
            columnas_adicionales_str = request.session.get('columnas_adicionales', '')
            extra_columns_flag = bool(columnas_adicionales_str)

            # Si hay columnas adicionales, localizar los índices de las columnas inválidas.
            extra_col_indices = []
            if extra_columns_flag:
                extra_cols = [c.strip() for c in columnas_adicionales_str.split(',') if c.strip()]
                # Buscar índices de las columnas adicionales por nombre en la cabecera.
                for cell in ws[1]:
                    if cell.value in extra_cols:
                        extra_col_indices.append(cell.column)

            # Mapeo de errores sin campo a sus columnas específicas.
            error_campo_map_loc = {
                "localizacion_duplicada": "columna",
                "subposicion_duplicada_excel": "columna",
                "caja_inconsistente": "caja",
                "bandeja_inconsistente": "bandeja",
                "rack_inconsistente": "rack",
                "posicion_rack_ocupada": "posicion_rack_estante",
                "posicion_bandeja_ocupada": "posicion_bandeja_rack",
                "posicion_caja_ocupada": "posicion_caja_rack"
            }
            # Recorrer filas con errores
            for fila_numero, info in errores.items():
                has_error = bool(info.get("bloqueantes"))
                if not has_error:
                    continue

                # Colorear celdas específicas y construir mensajes.
                mensajes = []
                for err in info.get("bloqueantes", []):
                    if ":" in err:
                        tipo, campo = err.split(":")
                        msg = f"(Error) {MENSAJES_ERROR[tipo]}"
                        if msg not in mensajes:
                            mensajes.append(msg)
                        if campo in columnas_excel:
                            col = columnas_excel[campo]
                            celda = ws.cell(row=int(fila_numero), column=col)
                            celda.fill = FILL_ERROR_CELL
                    else:
                        campo = error_campo_map_loc.get(err)
                        msg = f"(Error) {MENSAJES_ERROR[err]}"
                        if msg not in mensajes:
                            mensajes.append(msg)
                        if campo and campo in columnas_excel:
                            col_err = columnas_excel[campo]
                            ws.cell(row=int(fila_numero), column=col_err).fill = FILL_ERROR_CELL

                celda_errores = ws.cell(row=int(fila_numero), column=col_errores)
                celda_errores.value = "\n".join(mensajes)
            
            # Pintar las columnas extra después del resto de errores.
            if extra_col_indices:
                for col_idx in extra_col_indices:
                    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
                        for cell in row:
                            cell.fill = FILL_ERROR_COL
            
            output = io.BytesIO()    
            wb.save(output)
            wb.close()
            response = HttpResponse(output.getvalue(), content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename="listado_errores.xlsx"'
            return response        
    else:
        form = UploadExcel()
    return render(request, 'localizacion_nueva.html', build_upload_context({'form': form})) 


@login_required
@permission_required('muestras.can_add_localizaciones_web')
def crear_recurso_congelador(request):
    """Crea un recurso definitivo con su estructura base ya definida."""
    if request.method == 'POST':
        form = CrearRecursoCongeladorForm(request.POST, request.FILES)
        if form.is_valid():
            congelador = form.save()
            return redirect('detalles_congelador', nombre_congelador=congelador.congelador)
    else:
        form = CrearRecursoCongeladorForm()
    return render(request, 'crear_recurso_congelador.html', {'form': form})


@login_required
@permission_required('muestras.can_view_localizaciones_web')
def detalles_congelador(request, nombre_congelador):
    """Muestra la ficha completa de un recurso y su historial de reparaciones."""
    freezer = Congelador.objects.prefetch_related('reparaciones').filter(congelador=nombre_congelador)
    if not freezer:
        messages.error(request, 'No se ha encontrado el recurso solicitado.')
        return redirect('localizaciones_todas')
    template = loader.get_template('detalles_congelador.html')
    return HttpResponse(template.render({'congelador': freezer[0]}, request))


@login_required
@permission_required('muestras.can_change_localizaciones_web')
def editar_congelador(request, nombre_congelador):
    """Edita los metadatos de un recurso sin permitir cambiar su estructura."""
    congelador = Congelador.objects.filter(congelador=nombre_congelador).first()
    if not congelador:
        messages.error(request, 'No se ha encontrado el recurso solicitado.')
        return redirect('localizaciones_todas')
    nombre_anterior = congelador.congelador
    if request.method == 'POST':
        form = CongeladorForm(request.POST, request.FILES, instance=congelador)
        if form.is_valid():
            nombre_nuevo = form.cleaned_data.get('congelador')

            _set_foreign_key_checks(False)

            try:
                with transaction.atomic():
                    if nombre_nuevo != nombre_anterior:
                        with connection.cursor() as cursor:
                            # Actualizar FK en estantes (to_field='congelador').
                            cursor.execute(
                                "UPDATE muestras_estante SET congelador_id = %s WHERE congelador_id = %s",
                                [nombre_nuevo, nombre_anterior]
                            )
                            # Actualizar el histórico textual de localizaciones.
                            cursor.execute(
                                "UPDATE muestras_localizacion SET congelador = %s WHERE congelador = %s",
                                [nombre_nuevo, nombre_anterior]
                            )
                    form.save()
            finally:
                _set_foreign_key_checks(True)

            return redirect('detalles_congelador', nombre_congelador = form.instance.congelador)
    else:
        form = CongeladorForm(instance=congelador)
    return render(request, 'editar_congelador.html', {'form': form, 'congelador': congelador})


@login_required
@permission_required('muestras.can_change_localizaciones_web')
def anadir_reparacion_recurso(request, nombre_congelador):
    recurso = Congelador.objects.filter(congelador=nombre_congelador).first()
    if not recurso:
        messages.error(request, 'No se ha encontrado el recurso solicitado.')
        return redirect('localizaciones_todas')

    if request.method == 'POST':
        form = ReparacionRecursoForm(request.POST, recurso=recurso)
        if form.is_valid():
            reparacion = form.save(commit=False)
            reparacion.recurso = recurso
            reparacion.save()
            return redirect('detalles_congelador', nombre_congelador=recurso.congelador)
    else:
        form = ReparacionRecursoForm(recurso=recurso)

    return render(request, 'editar_reparacion_recurso.html', {
        'form': form,
        'congelador': recurso,
        'modo': 'crear',
    })


@login_required
@permission_required('muestras.can_change_localizaciones_web')
def editar_reparacion_recurso(request, reparacion_id):
    reparacion = ReparacionRecurso.objects.select_related('recurso').filter(pk=reparacion_id).first()
    if not reparacion:
        messages.error(request, 'No se ha encontrado la reparación solicitada.')
        return redirect('localizaciones_todas')

    if request.method == 'POST':
        form = ReparacionRecursoForm(request.POST, instance=reparacion, recurso=reparacion.recurso)
        if form.is_valid():
            form.save()
            return redirect('detalles_congelador', nombre_congelador=reparacion.recurso.congelador)
    else:
        form = ReparacionRecursoForm(instance=reparacion, recurso=reparacion.recurso)

    return render(request, 'editar_reparacion_recurso.html', {
        'form': form,
        'congelador': reparacion.recurso,
        'reparacion': reparacion,
        'modo': 'editar',
    })


@login_required
@permission_required('muestras.can_change_localizaciones_web')
def eliminar_reparacion_recurso(request, reparacion_id):
    reparacion = ReparacionRecurso.objects.select_related('recurso').filter(pk=reparacion_id).first()
    if not reparacion:
        messages.error(request, 'No se ha encontrado la reparación solicitada.')
        return redirect('localizaciones_todas')

    nombre_recurso = reparacion.recurso.congelador
    if request.method == 'POST':
        reparacion.delete()
    return redirect('detalles_congelador', nombre_congelador=nombre_recurso)

@login_required
@permission_required('muestras.can_delete_localizaciones_web')
def eliminar_localizacion(request):
    # Vista para eliminar localizaciones con progreso AJAX, requiere permiso para eliminar localizaciones
    def gen_eliminar_loc():
        conflictos = []

        congelador_ids = request.POST.getlist('congelador')
        estante_ids = request.POST.getlist('estante')
        rack_ids = request.POST.getlist('rack')
        bandeja_ids = request.POST.getlist('bandeja')
        caja_ids = request.POST.getlist('caja')
        subposicion_ids = request.POST.getlist('subposicion')

        total_items = len(congelador_ids) + len(estante_ids) + len(rack_ids) + len(bandeja_ids) + len(caja_ids) + len(subposicion_ids)
        yield _json_progress(0, total_items, 'start')

        try:
            current = 0

            # Comprobar congeladores seleccionados
            for congelador_id in congelador_ids:
                try:
                    cong = Congelador.objects.get(id=congelador_id)
                except Congelador.DoesNotExist:
                    current += 1
                    continue
                subposicion_ocupada = (
                    Subposicion.objects.select_related(
                        'muestra',
                        'caja__bandeja',
                        'caja__rack__estante__congelador',
                    )
                    .filter(caja__rack__estante__congelador=cong, vacia=False)
                    .first()
                )
                if subposicion_ocupada:
                    conflictos.append(
                        _build_delete_conflict_message('recurso', cong.congelador, subposicion_ocupada)
                    )
                current += 1
                if _should_update(current - 1, total_items):
                    yield _json_progress(current, total_items, 'processing')

            # Comprobar estantes seleccionados
            for estante_id in estante_ids:
                try:
                    est = Estante.objects.get(id=estante_id)
                except Estante.DoesNotExist:
                    current += 1
                    continue
                subposicion_ocupada = (
                    Subposicion.objects.select_related(
                        'muestra',
                        'caja__bandeja',
                        'caja__rack__estante__congelador',
                    )
                    .filter(caja__rack__estante=est, vacia=False)
                    .first()
                )
                if subposicion_ocupada:
                    conflictos.append(
                        _build_delete_conflict_message('estante', est.numero, subposicion_ocupada)
                    )
                current += 1
                if _should_update(current - 1, total_items):
                    yield _json_progress(current, total_items, 'processing')

            # Verificar racks seleccionados
            for rack_id in rack_ids:
                try:
                    rack = Rack.objects.get(id=rack_id)
                except Rack.DoesNotExist:
                    current += 1
                    continue
                subposicion_ocupada = (
                    Subposicion.objects.select_related(
                        'muestra',
                        'caja__bandeja',
                        'caja__rack__estante__congelador',
                    )
                    .filter(caja__rack=rack, vacia=False)
                    .first()
                )
                if subposicion_ocupada:
                    rack_label = 'cajón' if _usa_estructura_simple(rack.estante.congelador) else 'rack'
                    conflictos.append(
                        _build_delete_conflict_message(rack_label, rack.numero, subposicion_ocupada)
                    )
                current += 1
                if _should_update(current - 1, total_items):
                    yield _json_progress(current, total_items, 'processing')

            # Verificar cajas seleccionadas
            for caja_id in caja_ids:
                try:
                    caja = Caja.objects.get(id=caja_id)
                except Caja.DoesNotExist:
                    current += 1
                    continue
                subposicion_ocupada = (
                    Subposicion.objects.select_related(
                        'muestra',
                        'caja__bandeja',
                        'caja__rack__estante__congelador',
                    )
                    .filter(caja=caja, vacia=False)
                    .first()
                )
                if subposicion_ocupada:
                    conflictos.append(
                        _build_delete_conflict_message('caja', caja.numero, subposicion_ocupada, article='la')
                    )
                current += 1
                if _should_update(current - 1, total_items):
                    yield _json_progress(current, total_items, 'processing')

            # Verificar bandejas seleccionadas
            for bandeja_id in bandeja_ids:
                try:
                    bandeja = Bandeja.objects.get(id=bandeja_id)
                except Bandeja.DoesNotExist:
                    current += 1
                    continue
                subposicion_ocupada = (
                    Subposicion.objects.select_related(
                        'muestra',
                        'caja__bandeja',
                        'caja__rack__estante__congelador',
                    )
                    .filter(caja__bandeja=bandeja, vacia=False)
                    .first()
                )
                if subposicion_ocupada:
                    conflictos.append(
                        _build_delete_conflict_message('bandeja', bandeja.numero, subposicion_ocupada, article='la')
                    )
                current += 1
                if _should_update(current - 1, total_items):
                    yield _json_progress(current, total_items, 'processing')

            # Verificar subposiciones seleccionadas
            for subposicion_id in subposicion_ids:
                try:
                    subposicion = Subposicion.objects.get(id=subposicion_id)
                except Subposicion.DoesNotExist:
                    current += 1
                    continue
                if not subposicion.vacia:
                    muestra_label = subposicion.muestra.nom_lab if subposicion.muestra_id else 'una muestra'
                    conflictos.append(
                        f'No se puede eliminar la subposición "{subposicion.numero}" porque está ocupada por la muestra "{muestra_label}".'
                    )
                current += 1
                if _should_update(current - 1, total_items):
                    yield _json_progress(current, total_items, 'processing')

            # Si hay posiciones ocupadas, mostrar error
            if conflictos:
                mensaje = ' '.join(conflictos[:4])
                if len(conflictos) > 4:
                    mensaje += f' Además, hay {len(conflictos) - 4} elemento(s) más con muestras almacenadas.'
                yield _json_progress(current, total_items, 'error', mensaje)
                return

            # Proceder con la eliminación.
            if subposicion_ids:
                Subposicion.objects.filter(id__in=subposicion_ids).delete()
            if caja_ids:
                Caja.objects.filter(id__in=caja_ids).delete()
            if bandeja_ids:
                Bandeja.objects.filter(id__in=bandeja_ids).delete()
            if rack_ids:
                Rack.objects.filter(id__in=rack_ids).delete()
            if estante_ids:
                Estante.objects.filter(id__in=estante_ids).delete()
            if congelador_ids:
                Congelador.objects.filter(id__in=congelador_ids).delete()

            yield _json_progress(total_items, total_items, 'done', f'{total_items} posiciones eliminadas correctamente')
        except Exception as e:
            yield _json_progress(0, total_items, 'error', str(e))

    return StreamingHttpResponse(gen_eliminar_loc(), content_type='application/x-ndjson')

@login_required
@permission_required('muestras.can_view_localizaciones_web')
def exportar_posiciones_libres(request):
    """Exporta las posiciones libres en un Excel independiente por recurso."""
    congeladores = _get_congeladores_exportables(request.GET.getlist('congelador'))
    if not congeladores:
        messages.error(request, 'No se ha seleccionado ningún recurso.')
        return redirect('localizaciones_todas')
    archivos = [
        (
            f"posiciones_libres_{_slug_recurso(congelador)}.xlsx",
            _build_posiciones_libres_excel(congelador),
        )
        for congelador in congeladores
    ]
    return _response_archivos_exportacion(archivos, 'posiciones_libres_recursos.zip')


def _slug_recurso(congelador):
    nombre = slugify(congelador.congelador or '')
    return nombre or f"recurso_{congelador.id}"


def _get_congeladores_exportables(congelador_ids):
    if not congelador_ids:
        return []
    return list(
        Congelador.objects
        .filter(id__in=congelador_ids)
        .order_by('congelador')
    )


def _write_workbook(headers, rows, sheet_title):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title
    header_font = Font(bold=True)

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = header_font

    for row_num, row in enumerate(rows, 2):
        for col_num, value in enumerate(row, 1):
            ws.cell(row=row_num, column=col_num, value=value)

    for col_num, header in enumerate(headers, 1):
        col_letter = openpyxl.utils.get_column_letter(col_num)
        ws.column_dimensions[col_letter].width = max(len(str(header)) + 4, 18)

    output = io.BytesIO()
    wb.save(output)
    wb.close()
    output.seek(0)
    return output.getvalue()


def _rows_exportacion_recurso(congelador):
    if _usa_estructura_simple(congelador):
        subposiciones = (
            Subposicion.objects
            .filter(caja__rack__estante__congelador=congelador, caja__bandeja__isnull=True)
            .select_related('caja__rack__estante__congelador', 'muestra')
            .order_by(
                'caja__rack__estante__numero',
                'caja__rack__posicion_rack_estante',
                'caja__rack__numero',
                'caja__numero',
                'fila',
                'columna',
            )
        )
        headers = [
            'Recurso',
            'Estante',
            'Posición del cajón en el estante',
            'Cajón',
            'Caja',
            'Fila',
            'Columna',
            'Estado',
            'Muestra',
        ]
        rows = []
        for sub in subposiciones:
            caja = sub.caja
            rack = caja.rack
            rows.append([
                congelador.congelador,
                rack.estante.numero,
                rack.posicion_rack_estante,
                rack.numero,
                caja.numero,
                sub.fila,
                sub.columna,
                'Libre' if sub.vacia else 'Ocupada',
                sub.muestra.nom_lab if sub.muestra else '',
            ])
        return headers, rows

    subposiciones = (
        Subposicion.objects
        .filter(caja__rack__estante__congelador=congelador)
        .select_related('caja__bandeja__rack__estante__congelador', 'muestra')
        .order_by(
            'caja__rack__estante__numero',
            'caja__rack__posicion_rack_estante',
            'caja__rack__numero',
            'caja__bandeja__posicion_bandeja_rack',
            'caja__bandeja__numero',
            'caja__posicion_caja_rack',
            'caja__numero',
            'fila',
            'columna',
        )
    )
    headers = [
        'Congelador',
        'Estante',
        'Posición rack en estante',
        'Rack',
        'Posición bandeja en rack',
        'Bandeja',
        'Posición caja en bandeja',
        'Caja',
        'Fila',
        'Columna',
        'Estado',
        'Muestra',
    ]
    rows = []
    for sub in subposiciones:
        caja = sub.caja
        bandeja = caja.bandeja
        rack = caja.rack
        estante = rack.estante
        rows.append([
            congelador.congelador,
            estante.numero,
            rack.posicion_rack_estante,
            rack.numero,
            bandeja.posicion_bandeja_rack if bandeja else '',
            bandeja.numero if bandeja else '',
            caja.posicion_caja_rack,
            caja.numero,
            sub.fila,
            sub.columna,
            'Libre' if sub.vacia else 'Ocupada',
            sub.muestra.nom_lab if sub.muestra else '',
        ])
    return headers, rows


def _build_recurso_excel(congelador):
    headers, rows = _rows_exportacion_recurso(congelador)
    return _write_workbook(headers, rows, 'Recurso')


def _rows_posiciones_libres_recurso(congelador):
    if _usa_estructura_simple(congelador):
        subposiciones_libres = (
            Subposicion.objects
            .filter(vacia=True, caja__rack__estante__congelador=congelador, caja__bandeja__isnull=True)
            .select_related('caja__rack__estante__congelador')
            .order_by(
                'caja__rack__estante__numero',
                'caja__rack__posicion_rack_estante',
                'caja__rack__numero',
                'caja__numero',
                'fila',
                'columna',
            )
        )
        headers = [
            'Recurso',
            'Estante',
            'Posición del cajón en el estante',
            'Cajón',
            'Caja',
            'Fila disponible',
            'Columna disponible',
        ]
        rows = []
        for sub in subposiciones_libres:
            caja = sub.caja
            rack = caja.rack
            rows.append([
                congelador.congelador,
                rack.estante.numero,
                rack.posicion_rack_estante,
                rack.numero,
                caja.numero,
                sub.fila,
                sub.columna,
            ])
        return headers, rows

    subposiciones_libres = (
        Subposicion.objects
        .filter(vacia=True, caja__rack__estante__congelador=congelador)
        .select_related('caja__bandeja__rack__estante__congelador')
        .order_by(
            'caja__rack__estante__numero',
            'caja__rack__posicion_rack_estante',
            'caja__rack__numero',
            'caja__bandeja__posicion_bandeja_rack',
            'caja__bandeja__numero',
            'caja__posicion_caja_rack',
            'caja__numero',
            'fila',
            'columna',
        )
    )
    headers = [
        'Congelador',
        'Estante',
        'Posición rack en estante',
        'Rack',
        'Posición bandeja en rack',
        'Bandeja',
        'Posición caja en bandeja',
        'Caja',
        'Fila disponible',
        'Columna disponible',
    ]
    rows = []
    for sub in subposiciones_libres:
        caja = sub.caja
        bandeja = caja.bandeja
        rack = caja.rack
        estante = rack.estante
        rows.append([
            congelador.congelador,
            estante.numero,
            rack.posicion_rack_estante,
            rack.numero,
            bandeja.posicion_bandeja_rack if bandeja else '',
            bandeja.numero if bandeja else '',
            caja.posicion_caja_rack,
            caja.numero,
            sub.fila,
            sub.columna,
        ])
    return headers, rows


def _build_posiciones_libres_excel(congelador):
    headers, rows = _rows_posiciones_libres_recurso(congelador)
    return _write_workbook(headers, rows, 'Posiciones libres')


def _response_archivos_exportacion(archivos, nombre_zip):
    if len(archivos) == 1:
        nombre_archivo, contenido = archivos[0]
        return HttpResponse(
            contenido,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={'Content-Disposition': f'attachment; filename="{nombre_archivo}"'},
        )

    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for nombre_archivo, contenido in archivos:
            zip_file.writestr(nombre_archivo, contenido)
    output.seek(0)
    return HttpResponse(
        output.getvalue(),
        content_type='application/zip',
        headers={'Content-Disposition': f'attachment; filename="{nombre_zip}"'},
    )


@login_required
@permission_required('muestras.can_view_localizaciones_web')
def exportar_congeladores(request):
    """Exporta todos los recursos agrupados por tipo de estructura."""
    congeladores = list(Congelador.objects.order_by('congelador'))
    if not congeladores:
        messages.error(request, 'No hay recursos para exportar.')
        return redirect('localizaciones_todas')
    archivos = []
    grupos = [
        (Congelador.ESTRUCTURA_CONGELADOR_80, 'recursos_congelador_-80.xlsx'),
        (Congelador.ESTRUCTURA_CONGELADOR_20, 'recursos_congelador_-20.xlsx'),
        (Congelador.ESTRUCTURA_NEVERA, 'recursos_neveras.xlsx'),
    ]
    for tipo_estructura, nombre_archivo in grupos:
        recursos_tipo = [c for c in congeladores if c.tipo_estructura == tipo_estructura]
        if not recursos_tipo:
            continue
        headers, rows = _rows_exportacion_recurso(recursos_tipo[0])
        for recurso in recursos_tipo[1:]:
            _, recurso_rows = _rows_exportacion_recurso(recurso)
            rows.extend(recurso_rows)
        archivos.append((nombre_archivo, _write_workbook(headers, rows, 'Recursos')))

    if not archivos:
        messages.error(request, 'No hay recursos con estructura definida para exportar.')
        return redirect('localizaciones_todas')
    return _response_archivos_exportacion(archivos, 'recursos_por_tipo.zip')


@login_required
@permission_required('muestras.can_view_localizaciones_web')
def exportar_congeladores_seleccionados(request):
    """Exporta cada recurso seleccionado en su propio Excel."""
    congeladores = _get_congeladores_exportables(request.GET.getlist('congelador'))
    if not congeladores:
        messages.error(request, 'No se ha seleccionado ningún recurso.')
        return redirect('localizaciones_todas')
    archivos = [
        (
            f"recurso_{_slug_recurso(congelador)}.xlsx",
            _build_recurso_excel(congelador),
        )
        for congelador in congeladores
    ]
    return _response_archivos_exportacion(archivos, 'recursos_seleccionados.zip')


@login_required
@permission_required('muestras.can_add_localizaciones_web')
def descargar_plantilla_localizaciones_recurso(request, congelador_id):
    """Descarga la plantilla de importación adaptada al tipo de recurso."""
    congelador = Congelador.objects.filter(pk=congelador_id).first()
    if not congelador:
        return HttpResponse("El recurso no se encuentra disponible.", status=404)

    config = _get_upload_config(congelador)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ubicaciones"

    headers = list(config['rename_columns'].keys())
    for idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=idx, value=header)
        cell.font = Font(bold=True)

    output = io.BytesIO()
    wb.save(output)
    wb.close()
    output.seek(0)

    nombre_archivo = f"plantilla_{congelador.congelador}.xlsx"
    return HttpResponse(
        output.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={'Content-Disposition': f'attachment; filename="{nombre_archivo}"'},
    )


@login_required
def descargar_plantilla(request, macro: int):
    base_path = os.path.join(
        settings.BASE_DIR,
        'geslab_project',
        'globalstaticfiles',
        'files',
    )

    mapping = {
        0: 'plantilla_localizaciones.xlsx',
        1: 'plantilla_localizaciones_macros.xlsm',
        2: 'plantilla_muestras.xlsx',
        3: 'plantilla_estudios.xlsx',
        4: 'plantilla_cambio_posicion.xlsx',
    }

    nombre = mapping.get(macro)
    if not nombre:
        return HttpResponse("La plantilla no se encuentra disponible.", status=404)

    plantilla_path = os.path.join(base_path, nombre)
    if not os.path.exists(plantilla_path):
        return HttpResponse("La plantilla no se encuentra disponible.", status=404)

    return FileResponse(open(plantilla_path, 'rb'), as_attachment=True, filename=nombre)

