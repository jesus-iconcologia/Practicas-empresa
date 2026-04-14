import base64
import io
import json
import os
from datetime import date

import openpyxl
import pandas as pd
from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required, permission_required
from django.forms import formset_factory
from django.core.paginator import EmptyPage, PageNotAnInteger, Paginator
from django.db import connection, transaction
from django.db.models import Q
from django.http import HttpResponse, StreamingHttpResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.template import loader
from django.utils import timezone
from openpyxl.styles import PatternFill

from ..forms import (
    AliquotaForm,
    DatosMuestrasComunesForm,
    DestruirMuestrasForm,
    IndividuoEstudioForm,
    MaterialForm,
    MuestraForm,
    UploadExcel,
)
from ..models import (
    Bandeja,
    Caja,
    Congelador,
    Estante,
    Estudio,
    Localizacion,
    Material,
    Muestra,
    Rack,
    Subposicion,
    historial_estudios,
    historial_localizaciones,
    registro_destruido,
)
from ..parameters_config import get_excel_colors, get_upload_messages
from .common import _json_progress, _should_update


def _localizacion_desde_subposicion(muestra, subposicion):
    bandeja = getattr(subposicion.caja, "bandeja", None)
    return Localizacion.objects.create(
        muestra=muestra,
        congelador=subposicion.caja.rack.estante.congelador.congelador,
        estante=subposicion.caja.rack.estante.numero,
        rack=subposicion.caja.rack.numero,
        bandeja=bandeja.numero if bandeja else None,
        caja=subposicion.caja.numero,
        fila=subposicion.fila,
        columna=subposicion.columna,
        subposicion=subposicion.numero,
    )


def _subposicion_cache_key(subposicion):
    bandeja = getattr(subposicion.caja, "bandeja", None)
    return (
        str(subposicion.caja.rack.estante.congelador.congelador).lower(),
        subposicion.caja.rack.estante.numero,
        str(subposicion.caja.rack.numero).lower(),
        str(bandeja.numero).lower() if bandeja else "",
        str(subposicion.caja.numero).lower(),
        str(subposicion.fila).lower(),
        str(subposicion.columna).lower(),
    )


def _material_config_payload(materiales):
    return {
        str(material.id): {
            "nombre": material.nombre,
            "usa_volumen": material.usa_volumen,
            "unidades_volumen": material.unidades_volumen_list(),
            "usa_concentracion": material.usa_concentracion,
            "unidades_concentracion": material.unidades_concentracion_list(),
            "usa_unidades_bloque": material.usa_unidades_bloque,
            "unidades_recuento": material.unidades_recuento_list(),
        }
        for material in materiales
    }


IMPORT_MODE_SIN_UBICACION = "sin_ubicacion"
IMPORT_MODE_CONGELADOR_80 = Congelador.ESTRUCTURA_CONGELADOR_80
IMPORT_MODE_CONGELADOR_20 = Congelador.ESTRUCTURA_CONGELADOR_20
IMPORT_MODE_NEVERA = Congelador.ESTRUCTURA_NEVERA

IMPORT_MODE_CHOICES = [
    (IMPORT_MODE_SIN_UBICACION, "Solo datos de la muestra"),
    (IMPORT_MODE_CONGELADOR_80, "Datos de la muestra + ubicación en congelador -80"),
    (IMPORT_MODE_CONGELADOR_20, "Datos de la muestra + ubicación en congelador -20"),
    (IMPORT_MODE_NEVERA, "Datos de la muestra + ubicación en nevera"),
]

IMPORT_MODE_UI = {
    IMPORT_MODE_SIN_UBICACION: {
        "label": "Solo datos de la muestra",
        "short_label": "solo datos de la muestra",
        "card_text": "Crea muestras sin ubicación física en el momento de la importación.",
        "description": "Esta plantilla sirve para crear muestras sin asignar ubicación física en el momento de la importación.",
        "location_rule": "No se incluyen columnas de ubicación.",
        "columns_note": "La plantilla contiene solo los datos de la muestra: ID individuo, ID muestra, Material, cantidades, fechas, observaciones, centro de procedencia, estado actual y estudio.",
    },
    IMPORT_MODE_CONGELADOR_80: {
        "label": "Datos de la muestra + ubicación en congelador -80",
        "short_label": "ubicación en congelador -80",
        "card_text": "Incluye Congelador, Estante, Rack, Bandeja, Caja, Fila y Columna.",
        "description": "Esta plantilla sirve para crear muestras y ubicarlas directamente en recursos con estructura completa de congelador -80.",
        "location_rule": "La ubicación es obligatoria y debe indicarse con Congelador, Estante, Rack, Bandeja, Caja, Fila y Columna.",
        "columns_note": "La plantilla mantiene la parte común de datos de muestra y añade la jerarquía completa de ubicación.",
    },
    IMPORT_MODE_CONGELADOR_20: {
        "label": "Datos de la muestra + ubicación en congelador -20",
        "short_label": "ubicación en congelador -20",
        "card_text": "Incluye Congelador, Estante, Cajón, Caja, Fila y Columna.",
        "description": "Esta plantilla sirve para crear muestras y ubicarlas directamente en recursos de congelador -20.",
        "location_rule": "La ubicación es obligatoria y debe indicarse con Congelador, Estante, Cajón, Caja, Fila y Columna.",
        "columns_note": "La plantilla mantiene la parte común de datos de muestra y añade la jerarquía simple de ubicación.",
    },
    IMPORT_MODE_NEVERA: {
        "label": "Datos de la muestra + ubicación en nevera",
        "short_label": "ubicación en nevera",
        "card_text": "Incluye Congelador, Estante, Cajón, Caja, Fila y Columna.",
        "description": "Esta plantilla sirve para crear muestras y ubicarlas directamente en recursos de tipo nevera.",
        "location_rule": "La ubicación es obligatoria y debe indicarse con Congelador, Estante, Cajón, Caja, Fila y Columna.",
        "columns_note": "La plantilla mantiene la parte común de datos de muestra y añade la jerarquía simple de ubicación.",
    },
}

IMPORT_COMMON_COLUMNS = [
    "ID individuo",
    "ID muestra",
    "Material",
    "Volumen actual",
    "Unidad de volumen",
    "Concentración actual",
    "Unidad de concentración",
    "Cantidad actual",
    "Unidad de cantidad",
    "Fecha de extracción",
    "Fecha de llegada",
    "Observaciones",
    "Centro de procedencia",
    "Estado actual",
    "Estudio",
]

IMPORT_LOCATION_COLUMNS = {
    IMPORT_MODE_SIN_UBICACION: [],
    IMPORT_MODE_CONGELADOR_80: ["Congelador", "Estante", "Rack", "Bandeja", "Caja", "Fila", "Columna"],
    IMPORT_MODE_CONGELADOR_20: ["Congelador", "Estante", "Cajón", "Caja", "Fila", "Columna"],
    IMPORT_MODE_NEVERA: ["Congelador", "Estante", "Cajón", "Caja", "Fila", "Columna"],
}


def _get_import_mode(mode):
    valid_modes = {choice[0] for choice in IMPORT_MODE_CHOICES}
    return mode if mode in valid_modes else IMPORT_MODE_SIN_UBICACION


def _get_import_mode_ui(import_mode):
    return IMPORT_MODE_UI[_get_import_mode(import_mode)]


def _get_import_mode_cards():
    cards = []
    for value, label in IMPORT_MODE_CHOICES:
        ui = IMPORT_MODE_UI[value]
        cards.append({
            "value": value,
            "label": label,
            "card_text": ui["card_text"],
        })
    return cards


def _get_import_headers(import_mode):
    return IMPORT_COMMON_COLUMNS + IMPORT_LOCATION_COLUMNS[_get_import_mode(import_mode)]


def _get_import_required_headers(import_mode):
    headers = ["ID muestra", "Material", "Estado actual"]
    headers.extend(IMPORT_LOCATION_COLUMNS[_get_import_mode(import_mode)])
    return headers


def _get_import_rename_columns(import_mode):
    rename_columns = {
        "ID individuo": "id_individuo",
        "ID muestra": "nom_lab",
        "Material": "id_material",
        "Volumen actual": "volumen_actual",
        "Unidad de volumen": "unidad_volumen",
        "Concentración actual": "concentracion_actual",
        "Unidad de concentración": "unidad_concentracion",
        "Cantidad actual": "masa_actual",
        "Unidad de cantidad": "unidad_masa",
        "Fecha de extracción": "fecha_extraccion",
        "Fecha de llegada": "fecha_llegada",
        "Observaciones": "observaciones",
        "Centro de procedencia": "centro_procedencia",
        "Estado actual": "estado_actual",
        "Estudio": "estudio",
    }
    mode = _get_import_mode(import_mode)
    if mode == IMPORT_MODE_CONGELADOR_80:
        rename_columns.update({
            "Congelador": "congelador",
            "Estante": "estante",
            "Rack": "rack",
            "Bandeja": "bandeja",
            "Caja": "caja",
            "Fila": "fila",
            "Columna": "columna",
        })
    elif mode in {IMPORT_MODE_CONGELADOR_20, IMPORT_MODE_NEVERA}:
        rename_columns.update({
            "Congelador": "congelador",
            "Estante": "estante",
            "Cajón": "rack",
            "Caja": "caja",
            "Fila": "fila",
            "Columna": "columna",
        })
    return rename_columns


def _build_import_template_workbook(import_mode):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Muestras"

    headers = _get_import_headers(import_mode)
    required_headers = set(_get_import_required_headers(import_mode))
    required_fill = PatternFill("solid", fgColor="FDE2E1")

    for column_index, header in enumerate(headers, start=1):
        cell = sheet.cell(row=1, column=column_index)
        cell.value = header
        if header in required_headers:
            cell.fill = required_fill

    widths = {
        "ID individuo": 18,
        "ID muestra": 24,
        "Material": 24,
        "Volumen actual": 16,
        "Unidad de volumen": 18,
        "Concentración actual": 18,
        "Unidad de concentración": 20,
        "Cantidad actual": 16,
        "Unidad de cantidad": 18,
        "Fecha de extracción": 18,
        "Fecha de llegada": 18,
        "Observaciones": 28,
        "Centro de procedencia": 24,
        "Estado actual": 18,
        "Estudio": 14,
        "Congelador": 18,
        "Estante": 12,
        "Rack": 12,
        "Cajón": 12,
        "Bandeja": 12,
        "Caja": 12,
        "Fila": 10,
        "Columna": 10,
    }
    for column_index, header in enumerate(headers, start=1):
        sheet.column_dimensions[openpyxl.utils.get_column_letter(column_index)].width = widths.get(header, 18)

    return workbook


def _crear_muestra_desde_alicuota(individuo_form, aliquota_form, user, datos_comunes=None):
    datos_comunes = datos_comunes or {}
    muestra = Muestra.objects.create(
        id_individuo=individuo_form.cleaned_data["id_individuo"],
        estudio=individuo_form.cleaned_data.get("estudio"),
        nom_lab=aliquota_form.cleaned_data["nom_lab"],
        material=aliquota_form.cleaned_data.get("material"),
        volumen_actual=aliquota_form.cleaned_data.get("volumen_actual"),
        unidad_volumen=aliquota_form.cleaned_data.get("unidad_volumen") or None,
        concentracion_actual=aliquota_form.cleaned_data.get("concentracion_actual"),
        unidad_concentracion=aliquota_form.cleaned_data.get("unidad_concentracion") or None,
        masa_actual=aliquota_form.cleaned_data.get("masa_actual"),
        unidad_masa=aliquota_form.cleaned_data.get("unidad_masa") or None,
        fecha_extraccion=datos_comunes.get("fecha_extraccion"),
        fecha_llegada=datos_comunes.get("fecha_llegada"),
        observaciones=datos_comunes.get("observaciones"),
        centro_procedencia=datos_comunes.get("centro_procedencia"),
        estado_actual=aliquota_form.cleaned_data.get("estado_actual") or "DISP",
    )

    if muestra.estudio:
        historial_estudios.objects.create(
            muestra=muestra,
            estudio=muestra.estudio,
            fecha_asignacion=timezone.now(),
            usuario_asignacion=user,
        )

    subposicion_id = aliquota_form.cleaned_data.get("subposicion_id")
    if not subposicion_id:
        return muestra

    subposicion = Subposicion.objects.select_for_update().get(id=subposicion_id)
    subposicion.muestra = muestra
    subposicion.vacia = False
    subposicion.save()

    localizacion = _localizacion_desde_subposicion(muestra, subposicion)
    historial_localizaciones.objects.create(
        muestra=muestra,
        localizacion=localizacion,
        fecha_asignacion=timezone.now(),
        usuario_asignacion=user,
    )
    return muestra


def _material_boolean_from_text(value):
    if value is None:
        return False
    text = str(value).strip().lower()
    return text in {"1", "true", "si", "sí", "x", "yes"}


def _excel_text_or_none(value):
    if value is None or pd.isna(value):
        return None
    text = str(value).strip()
    return text or None


def _parse_material_bool(value):
    text = _excel_text_or_none(value)
    if text is None:
        return None
    lowered = text.lower()
    if lowered in {"1", "true", "si", "sí", "x", "yes"}:
        return True
    if lowered in {"0", "false", "no", "n", "off"}:
        return False
    return "invalid"


@login_required
@permission_required('muestras.can_add_muestras_web')
def materiales_todos(request):
    materiales = Material.objects.all().order_by('nombre')

    nombre = request.GET.get('nombre', '').strip()
    if nombre:
        materiales = materiales.filter(nombre__icontains=nombre)

    usa_volumen = request.GET.get('usa_volumen', '').strip()
    if usa_volumen == 'si':
        materiales = materiales.filter(usa_volumen=True)
    elif usa_volumen == 'no':
        materiales = materiales.filter(usa_volumen=False)

    usa_concentracion = request.GET.get('usa_concentracion', '').strip()
    if usa_concentracion == 'si':
        materiales = materiales.filter(usa_concentracion=True)
    elif usa_concentracion == 'no':
        materiales = materiales.filter(usa_concentracion=False)

    usa_unidades = request.GET.get('usa_unidades', '').strip()
    if usa_unidades == 'si':
        materiales = materiales.filter(usa_unidades_bloque=True)
    elif usa_unidades == 'no':
        materiales = materiales.filter(usa_unidades_bloque=False)

    busqueda = request.GET.get('busqueda', '').strip()
    if busqueda:
        materiales = materiales.filter(nombre__icontains=busqueda)

    contador_total = materiales.count()
    items_por_pagina = request.GET.get('items_por_pagina', 25)
    if str(items_por_pagina) == 'todas':
        items_por_pagina = 'todas'
        paginator = Paginator(materiales, max(contador_total, 1))
    else:
        try:
            items_por_pagina = int(items_por_pagina)
            if items_por_pagina not in [10, 25, 50, 100]:
                items_por_pagina = 25
        except (ValueError, TypeError):
            items_por_pagina = 25
        paginator = Paginator(materiales, items_por_pagina)

    numero_pagina = request.GET.get('page', 1)
    try:
        materiales_page = paginator.page(numero_pagina)
    except PageNotAnInteger:
        materiales_page = paginator.page(1)
    except EmptyPage:
        materiales_page = paginator.page(paginator.num_pages)

    return render(request, 'materiales_todos.html', {
        'materiales': materiales_page.object_list,
        'materiales_page': materiales_page,
        'paginator': paginator,
        'contador_materiales': contador_total,
        'items_por_pagina': items_por_pagina,
        'busqueda': busqueda,
        'request': request,
    })


@login_required
@permission_required('muestras.can_add_muestras_web')
def acciones_materiales(request):
    if request.method != 'POST':
        return redirect('materiales_todos')

    materiales_ids = request.POST.getlist('material_id')
    queryset = Material.objects.filter(id__in=materiales_ids)

    if 'exportar_seleccionados' in request.POST:
        if not materiales_ids:
            messages.error(request, 'No has seleccionado ningún material para exportar.')
            return redirect('materiales_todos')

        response = HttpResponse(
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename="materiales_seleccionados.xlsx"'

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Materiales"
        headers = [
            "Nombre del material",
            "Usa volumen",
            "Unidades de volumen",
            "Usa concentración",
            "Unidades de concentración",
            "Usa unidades",
            "Unidades",
        ]
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col).value = header

        row_num = 2
        for material in queryset.order_by('nombre'):
            ws.cell(row=row_num, column=1).value = material.nombre
            ws.cell(row=row_num, column=2).value = 'Sí' if material.usa_volumen else 'No'
            ws.cell(row=row_num, column=3).value = material.unidades_volumen or ''
            ws.cell(row=row_num, column=4).value = 'Sí' if material.usa_concentracion else 'No'
            ws.cell(row=row_num, column=5).value = material.unidades_concentracion or ''
            ws.cell(row=row_num, column=6).value = 'Sí' if material.usa_unidades_bloque else 'No'
            ws.cell(row=row_num, column=7).value = material.unidades_recuento or ''
            row_num += 1

        wb.save(response)
        return response

    if 'eliminar' in request.POST:
        if not materiales_ids:
            messages.error(request, 'No has seleccionado ningún material para eliminar.')
            return redirect('materiales_todos')

        materiales_con_uso = list(queryset.filter(muestras__isnull=False).distinct())
        if materiales_con_uso:
            nombres = ", ".join(material.nombre for material in materiales_con_uso[:5])
            messages.error(request, f'No se pueden eliminar materiales que ya están asociados a muestras: {nombres}.')
            return redirect('materiales_todos')

        total = queryset.count()
        queryset.delete()
        messages.success(request, f'{total} materiales eliminados correctamente.')
        return redirect('materiales_todos')

    return redirect('materiales_todos')


@login_required
@permission_required('muestras.can_add_muestras_web')
def upload_excel_materiales(request):
    if request.GET.get('mostrar_confirmacion') and 'confirmacion_materiales_pendiente' in request.session:
        datos_conf = request.session.pop('confirmacion_materiales_pendiente')
        for msg in datos_conf.get('mensajes', []):
            getattr(messages, msg['level'])(request, msg['text'])
        return render(request, datos_conf['template'], datos_conf.get('context', {}))

    if request.method == 'POST':
        form = UploadExcel(request.POST, request.FILES)
        if 'confirmar' in request.POST:
            filas_validas = request.session.get('filas_validas_materiales', [])
            with transaction.atomic():
                for datos in filas_validas:
                    Material.objects.create(
                        nombre=datos['nombre'],
                        usa_volumen=datos['usa_volumen'],
                        unidades_volumen=datos['unidades_volumen'],
                        usa_concentracion=datos['usa_concentracion'],
                        unidades_concentracion=datos['unidades_concentracion'],
                        usa_unidades_bloque=datos['usa_unidades_bloque'],
                        unidades_recuento=datos['unidades_recuento'],
                        activo=True,
                    )
            request.session.pop('filas_validas_materiales', None)
            request.session.pop('errores_materiales', None)
            request.session.pop('excel_materiales_base64', None)
            request.session.pop('columnas_adicionales_materiales', None)
            messages.success(request, 'Los materiales se han creado correctamente.')
            return redirect('materiales_todos')

        if 'cancelar' in request.POST:
            request.session.pop('filas_validas_materiales', None)
            request.session.pop('errores_materiales', None)
            request.session.pop('excel_materiales_base64', None)
            request.session.pop('columnas_adicionales_materiales', None)
            messages.error(request, 'La importación de materiales se ha cancelado.')
            return redirect('upload_excel_materiales')

        if 'excel_errores' in request.POST:
            errores = request.session.get('errores_materiales', {})
            excel_b64 = request.session.get('excel_materiales_base64')
            if not excel_b64:
                messages.error(request, 'No hay un Excel validado para descargar.')
                return redirect('upload_excel_materiales')

            wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(excel_b64)))
            ws = wb.active
            colors = get_excel_colors()
            error_fill = PatternFill("solid", fgColor=colors['error_row'])

            for fila, detalle in errores.items():
                if detalle.get('bloqueantes'):
                    for cell in ws[fila]:
                        cell.fill = error_fill

            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename="materiales_con_errores.xlsx"'
            wb.save(response)
            return response

        if form.is_valid():
            excel_file = request.FILES['excel_file']
            excel_bytes = excel_file.read()
            request.session['excel_materiales_base64'] = base64.b64encode(excel_bytes).decode()
            df = pd.read_excel(io.BytesIO(excel_bytes))

            rename_columns = {
                'nombre del material': 'nombre',
                'usa volumen': 'usa_volumen',
                'unidades de volumen': 'unidades_volumen',
                'usa concentración': 'usa_concentracion',
                'unidades de concentración': 'unidades_concentracion',
                'usa unidades': 'usa_unidades_bloque',
                'unidades': 'unidades_recuento',
            }
            df.columns = df.columns.str.strip().str.lower()
            columnas_esperadas = set(rename_columns.keys())
            columnas_existentes = set(df.columns)
            columnas_faltantes = columnas_esperadas - columnas_existentes

            if columnas_faltantes:
                columnas_str = ', '.join(sorted(columnas_faltantes))
                return render(request, 'upload_excel_materiales.html', {'form': form, 'error': f'Error de formato: faltan columnas esperadas: {columnas_str}'})

            if df.empty or len(df) == 0:
                return render(request, 'upload_excel_materiales.html', {'form': form, 'error': 'Error de formato: el archivo Excel está vacío o no contiene filas de datos.'})

            columnas_adicionales = columnas_existentes - columnas_esperadas
            columnas_adicionales_str = ', '.join(sorted(columnas_adicionales))
            request.session['columnas_adicionales_materiales'] = columnas_adicionales_str

            df = df.rename(columns=rename_columns)
            errores = {}
            filas_validas = []
            nombres_excel = set()
            nombres_existentes = {nombre.lower() for nombre in Material.objects.values_list('nombre', flat=True)}

            for idx, row in df.iterrows():
                fila = idx + 2
                errores[fila] = {'bloqueantes': []}

                nombre = _excel_text_or_none(row.get('nombre'))
                usa_volumen = _parse_material_bool(row.get('usa_volumen'))
                usa_concentracion = _parse_material_bool(row.get('usa_concentracion'))
                usa_unidades_bloque = _parse_material_bool(row.get('usa_unidades_bloque'))

                datos = {
                    'nombre': nombre,
                    'usa_volumen': bool(usa_volumen is True),
                    'unidades_volumen': _excel_text_or_none(row.get('unidades_volumen')),
                    'usa_concentracion': bool(usa_concentracion is True),
                    'unidades_concentracion': _excel_text_or_none(row.get('unidades_concentracion')),
                    'usa_unidades_bloque': bool(usa_unidades_bloque is True),
                    'unidades_recuento': _excel_text_or_none(row.get('unidades_recuento')),
                }

                if not nombre:
                    errores[fila]['bloqueantes'].append('campo_obligatorio_vacio:nombre')
                else:
                    nombre_lower = nombre.lower()
                    if nombre_lower in nombres_existentes:
                        errores[fila]['bloqueantes'].append('material_existente_bd')
                    if nombre_lower in nombres_excel:
                        errores[fila]['bloqueantes'].append('material_duplicado_excel')
                    else:
                        nombres_excel.add(nombre_lower)

                if usa_volumen == "invalid":
                    errores[fila]['bloqueantes'].append('valor_invalido:usa_volumen')
                if usa_concentracion == "invalid":
                    errores[fila]['bloqueantes'].append('valor_invalido:usa_concentracion')
                if usa_unidades_bloque == "invalid":
                    errores[fila]['bloqueantes'].append('valor_invalido:usa_unidades')

                if not errores[fila]['bloqueantes']:
                    filas_validas.append(datos)

            request.session['filas_validas_materiales'] = filas_validas
            request.session['errores_materiales'] = errores

            numero_errores_bloqueantes = sum(1 for detalle in errores.values() if detalle['bloqueantes'])
            errores_encontrados = numero_errores_bloqueantes > 0 or bool(columnas_adicionales)
            mensajes = [{'level': 'info', 'text': f'El excel contiene {len(df)} registros.'}]
            if numero_errores_bloqueantes == 0 and not columnas_adicionales:
                mensajes.append({'level': 'success', 'text': 'No tiene errores en ningún campo.'})
            else:
                if numero_errores_bloqueantes > 0:
                    mensajes.append({'level': 'error', 'text': f'Contiene {numero_errores_bloqueantes} filas con errores graves.'})
                if columnas_adicionales:
                    mensajes.append({'level': 'warning', 'text': f'Contiene {len(columnas_adicionales)} columnas extras: {columnas_adicionales_str}'})

            request.session['confirmacion_materiales_pendiente'] = {
                'template': 'confirmacion_upload_materiales.html',
                'context': {
                    'errores_encontrados': errores_encontrados,
                    'numero_errores_bloqueantes': numero_errores_bloqueantes,
                    'tiene_columnas_extras': bool(columnas_adicionales),
                },
                'mensajes': mensajes,
            }
            request.session.save()
            return redirect(f"{request.path}?mostrar_confirmacion=1")
    else:
        form = UploadExcel()

    return render(request, 'upload_excel_materiales.html', {'form': form})


@login_required
@permission_required('muestras.can_add_muestras_web')
def descargar_plantilla_materiales(request):
    plantilla_path = os.path.join(
        settings.BASE_DIR,
        'geslab_project',
        'globalstaticfiles',
        'files',
        'plantilla_materiales.xlsx',
    )
    if not os.path.exists(plantilla_path):
        return HttpResponse("La plantilla no se encuentra disponible.", status=404)

    with open(plantilla_path, 'rb') as plantilla_file:
        response = HttpResponse(
            plantilla_file.read(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename="plantilla_materiales.xlsx"'
        return response


@login_required
@permission_required('muestras.can_add_muestras_web')
def nuevo_material(request):
    if request.method == 'POST':
        form = MaterialForm(request.POST)
        if form.is_valid():
            form.save()
            messages.success(request, 'Se ha creado el material con exito.')
            return redirect('materiales_todos')
    else:
        form = MaterialForm()
    return render(request, 'nuevo_material.html', {'form': form})


@login_required
@permission_required('muestras.can_add_muestras_web')
def editar_material(request, material_id):
    material = get_object_or_404(Material, pk=material_id)
    if request.method == 'POST':
        form = MaterialForm(request.POST, instance=material)
        if form.is_valid():
            form.save()
            messages.success(request, 'Material actualizado correctamente.')
            return redirect('materiales_todos')
    else:
        form = MaterialForm(instance=material)
    return render(request, 'editar_material.html', {'form': form, 'material': material})


@login_required
@permission_required('muestras.can_add_muestras_web')
def eliminar_material(request, material_id):
    if request.method != 'POST':
        return redirect('materiales_todos')
    material = get_object_or_404(Material, pk=material_id)
    if material.muestras.exists():
        messages.error(request, 'No se puede eliminar un material que ya está asociado a muestras.')
        return redirect('materiales_todos')
    material.delete()
    return redirect('materiales_todos')


@login_required
def principal(request):
    # Vista principal de la aplicaciÃ³n, muestra una pÃ¡gina de bienvenida
    template = loader.get_template('principal.html')
    return HttpResponse(template.render(request=request))


# Vistas para Muestras
@login_required
@permission_required('muestras.can_view_muestras_web')
def muestras_todas(request):
    # Vista que muestra todas las muestras y su localizaciÃ³n asociada, requiere que el usuario estÃ© autenticado
    muestras = (Muestra.objects.select_related('subposicion__caja__bandeja__rack__estante__congelador'))

    # Filtrado de muestras si se proporcionan parÃ¡metros de bÃºsqueda en los filtros del template
    field_names = [f.name for f in Muestra._meta.local_fields if f.name not in ('id','estudio')]
    field_names_readable = ['Id del individuo','Nombre dado por el laboratorio','Material','Volumen actual','Unidad de volumen','ConcentraciÃ³n actual','Unidad de concentraciÃ³n','Masa actual','Unidad de masa','Fecha de extracciÃ³n','Fecha de llegada','Observaciones','Estado inicial','Centro de procedencia','Lugar de procedencia','Estado actual']
    field_names_readable_dict = {k:v for (k,v) in zip(field_names,field_names_readable)}
    
    # Filtrado de campos de texto normales
    for field in field_names:
        val = request.GET.get(field)
        if val:
            # Separar por punto y coma (;) para permitir mÃºltiples filtros
            valores = [v.strip() for v in val.split(';') if v.strip()]
            if valores:
                # Para IDs (exact match)
                if field in ['id_individuo', 'nom_lab']:
                    q_filters = Q()
                    for valor in valores:
                        if valor == 'null':
                            # Incluir tanto NULL como cadenas vacÃ­as
                            q_filters |= Q(**{f"{field}__isnull": True}) | Q(**{f"{field}": ""})
                        else:
                            q_filters |= Q(**{f"{field}__iexact": valor})
                    muestras = muestras.filter(q_filters)
                # Para campos que son dropdowns, usar bÃºsqueda exacta
                elif field in ['id_material', 'centro_procedencia', 'estado_actual']:
                    q_filters = Q()
                    for valor in valores:
                        if valor == 'null':
                            # Incluir tanto NULL como cadenas vacÃ­as
                            q_filters |= Q(**{f"{field}__isnull": True}) | Q(**{f"{field}": ""})
                        else:
                            q_filters |= Q(**{f"{field}": valor})
                    muestras = muestras.filter(q_filters)
                else:
                    # Para campos de texto: bÃºsqueda icontains para cada valor
                    q_filters = Q()
                    for valor in valores:
                        q_filters |= Q(**{f"{field}__icontains": valor})
                    muestras = muestras.filter(q_filters)
    
    # Filtrado de las muestras en base al estudio especÃ­fico (por ID)
    if request.GET.get('estudio'):
        estudio_val = request.GET['estudio']
        if estudio_val:
            # Separar por punto y coma para mÃºltiples estudios
            estudios_ids = [e.strip() for e in estudio_val.split(';') if e.strip()]
            if estudios_ids:
                q_filters = Q()
                for estudio_id in estudios_ids:
                    if estudio_id == 'null':
                        q_filters |= Q(estudio__isnull=True)
                    else:
                        try:
                            q_filters |= Q(estudio__id=int(estudio_id))
                        except (ValueError, TypeError):
                            pass
                muestras = muestras.filter(q_filters)

    # Filtrado de material por texto libre (icontains, separado por ;)
    if request.GET.get('id_material_texto'):
        valores = [v.strip() for v in request.GET['id_material_texto'].split(';') if v.strip()]
        if valores:
            q_filters = Q()
            for valor in valores:
                q_filters |= Q(id_material__icontains=valor)
            muestras = muestras.filter(q_filters)

    # Filtrado de estudio por texto libre (nombre del estudio, icontains, separado por ;)
    if request.GET.get('estudio_texto'):
        valores = [v.strip() for v in request.GET['estudio_texto'].split(';') if v.strip()]
        if valores:
            q_filters = Q()
            for valor in valores:
                if valor.lower() == 'null' or valor.lower() == 'sin estudio':
                    q_filters |= Q(estudio__isnull=True)
                else:
                    q_filters |= Q(estudio__nombre_estudio__icontains=valor)
            muestras = muestras.filter(q_filters)

    # Filtrado de centro de procedencia por texto libre (icontains, separado por ;)
    if request.GET.get('centro_procedencia_texto'):
        valores = [v.strip() for v in request.GET['centro_procedencia_texto'].split(';') if v.strip()]
        if valores:
            q_filters = Q()
            for valor in valores:
                q_filters |= Q(centro_procedencia__icontains=valor)
            muestras = muestras.filter(q_filters)

    # Filtrado de estado actual por texto libre (icontains, separado por ;)
    if request.GET.get('estado_actual_texto'):
        valores = [v.strip() for v in request.GET['estado_actual_texto'].split(';') if v.strip()]
        if valores:
            q_filters = Q()
            for valor in valores:
                q_filters |= Q(estado_actual__icontains=valor)
            muestras = muestras.filter(q_filters)
    
    # Filtrado por localizaciones
    if request.GET.get('cajon'):
        cajon_val = request.GET['cajon']
        cajones = [c.strip() for c in cajon_val.split(';') if c.strip()]
        if cajones:
            q_filters = Q()
            for cajon in cajones:
                if cajon == 'null':
                    q_filters |= Q(subposicion__isnull=True)
                else:
                    q_filters |= Q(
                        subposicion__caja__rack__numero__iexact=cajon,
                        subposicion__caja__rack__estante__congelador__tipo_estructura__in=['congelador_20', 'nevera'],
                    )
            muestras = muestras.filter(q_filters)

    if request.GET.get('congelador'):
        congelador_val = request.GET['congelador']
        congeladores = [c.strip() for c in congelador_val.split(';') if c.strip()]
        if congeladores:
            q_filters = Q()
            for congelador in congeladores:
                if congelador == 'null':
                    q_filters |= Q(subposicion__isnull=True)
                else:
                    q_filters |= Q(subposicion__caja__rack__estante__congelador__congelador__iexact=congelador)
            muestras = muestras.filter(q_filters)
    
    if request.GET.get('estante'):
        estante_val = request.GET['estante']
        estantes = [e.strip() for e in estante_val.split(';') if e.strip()]
        if estantes:
            q_filters = Q()
            for estante in estantes:
                if estante == 'null':
                    q_filters |= Q(subposicion__isnull=True)
                else:
                    try:
                        q_filters |= Q(subposicion__caja__rack__estante__numero=int(estante))
                    except (ValueError, TypeError):
                        pass
            muestras = muestras.filter(q_filters)
    
    if request.GET.get('rack'):
        rack_val = request.GET['rack']
        racks = [r.strip() for r in rack_val.split(';') if r.strip()]
        if racks:
            q_filters = Q()
            for rack in racks:
                if rack == 'null':
                    q_filters |= Q(subposicion__isnull=True)
                else:
                    q_filters |= Q(subposicion__caja__rack__numero__iexact=rack)
            muestras = muestras.filter(q_filters)
    
    if request.GET.get('caja'):
        caja_val = request.GET['caja']
        cajas = [c.strip() for c in caja_val.split(';') if c.strip()]
        if cajas:
            q_filters = Q()
            for caja in cajas:
                if caja == 'null':
                    q_filters |= Q(subposicion__isnull=True)
                else:
                    q_filters |= Q(subposicion__caja__numero__iexact=caja)
            muestras = muestras.filter(q_filters)

    # Filtrado de las muestras a mostrar si el perfil es de un investigador, mostrando solo las asociadas a sus estudios
    if request.user.groups.filter(name='Investigadores').exists():
        muestras = muestras.filter(estudio__investigadores_asociados=request.user)
    
    # BÃºsqueda general (server-side): filtra en todos los campos de texto relevantes
    busqueda_general = request.GET.get('busqueda', '').strip()
    if busqueda_general:
        q_busqueda = Q()
        q_busqueda |= Q(id_individuo__icontains=busqueda_general)
        q_busqueda |= Q(nom_lab__icontains=busqueda_general)
        q_busqueda |= Q(id_material__icontains=busqueda_general)
        q_busqueda |= Q(observaciones__icontains=busqueda_general)
        q_busqueda |= Q(centro_procedencia__icontains=busqueda_general)
        q_busqueda |= Q(estado_actual__icontains=busqueda_general)
        q_busqueda |= Q(estudio__nombre_estudio__icontains=busqueda_general)
        muestras = muestras.filter(q_busqueda)

    '''
    # Crear un PDF con las muestras filtradas
    if request.GET.get('crear_pdf'):    
        buffer = BytesIO()
        p = canvas.Canvas(buffer)
        y = 800
        p.setFont("Helvetica", 16)
        p.drawString(30,y, "Listado de Muestras")
        p.setFont("Helvetica", 12)
        y -= 30
        p.drawString(30, y, "ID Individuo")
        p.drawString(150, y, "Nombre Laboratorio")
        p.drawString(300, y, "LocalizaciÃ³n")
        y-= 30
        for muestra in muestras:
            p.drawString(30, y, muestra.id_individuo)
            p.drawString(150, y, muestra.nom_lab)
            p.drawString(300, y, str(muestra.localizacion.first()) if muestra.localizacion.exists() else 'No archivada')
            y -= 20
            if y < 50:
                p.showPage()
                y = 800
        p.save()
        buffer.seek(0)
        return FileResponse(buffer, as_attachment=True, filename='listado_muestras.pdf')
    '''



     # Crear un Excel con las muestras filtradas 
    if request.GET.get('exportar_excel'):
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="listado_muestras.xlsx"'
        wb = openpyxl.load_workbook(os.path.join(settings.BASE_DIR, 'geslab_project', 'globalstaticfiles', 'files', 'listado_muestras.xlsx'))
        ws = wb.active
        row_num = 2
        for muestra in muestras:
            col_num = 1
            
            for field in field_names:
                value = muestra.__dict__[field]
                if value is None:
                    value = ''
                ws.cell(row_num, col_num).value= str(value)
                col_num += 1
            value = muestra.estudio.nombre_estudio if muestra.estudio else ''
            ws.cell(row_num, col_num).value= str(value)
            col_num += 1
            value = muestra.posicion_completa()
            if value is None:
                value = ''
            else:
                value = value.split("-")
                for columna in value:
                    ws.cell(row_num, col_num).value= str(columna)
                    col_num += 1
            row_num += 1
    
        wb.save(response)
        return response
    
    # Obtener opciones para los dropdowns
    opciones_materiales = Muestra.objects.values_list('id_material', flat=True).distinct().exclude(id_material__isnull=True).exclude(id_material='')
    opciones_centros = Muestra.objects.values_list('centro_procedencia', flat=True).distinct().exclude(centro_procedencia__isnull=True).exclude(centro_procedencia='')
    opciones_estudios = Estudio.objects.all()
    opciones_congeladores = Congelador.objects.values_list('congelador', flat=True).distinct()
    opciones_estantes = Estante.objects.values_list('numero', flat=True).distinct().order_by('numero')
    opciones_cajones = Rack.objects.filter(
        estante__congelador__tipo_estructura__in=['congelador_20', 'nevera']
    ).values_list('numero', flat=True).distinct().order_by('numero')
    opciones_racks = Rack.objects.values_list('numero', flat=True).distinct().order_by('numero')
    opciones_cajas = Caja.objects.values_list('numero', flat=True).distinct().order_by('numero')
    

    # PaginaciÃ³n
    contador_total = muestras.count()
    items_por_pagina = request.GET.get('items_por_pagina', 25)
    if str(items_por_pagina) == 'todas':
        items_por_pagina = 'todas'
        paginator = Paginator(muestras, max(contador_total, 1))
    else:
        try:
            items_por_pagina = int(items_por_pagina)
            if items_por_pagina not in [10, 25, 50, 100]:
                items_por_pagina = 25
        except (ValueError, TypeError):
            items_por_pagina = 25
        paginator = Paginator(muestras, items_por_pagina)
    numero_pagina = request.GET.get('page', 1)
    
    try:
        muestras_pagina = paginator.page(numero_pagina)
    except PageNotAnInteger:
        muestras_pagina = paginator.page(1)
    except EmptyPage:
        muestras_pagina = paginator.page(paginator.num_pages)

    # Cargar el template y pasar las muestras y los campos de filtro al mismo
    template = loader.get_template('muestras_todas.html')
    context = {    
        'muestras': muestras,
        'contador_muestras': muestras.count(),
        'field_names': field_names,
        'field_names_readable_dict': field_names_readable_dict,
        'opciones_materiales': opciones_materiales,
        'opciones_centros': opciones_centros,
        'opciones_estudios': opciones_estudios,
        'opciones_congeladores': opciones_congeladores,
        'opciones_estantes': opciones_estantes,
        'opciones_cajones': opciones_cajones,
        'opciones_racks': opciones_racks,
        'opciones_cajas': opciones_cajas,
        'muestras_page': muestras_pagina,
        'paginator': paginator,
        'items_por_pagina': items_por_pagina,
        'busqueda': busqueda_general,
        'muestras_ids_json': json.dumps(list(muestras.values_list('id', flat=True))),
    }
    return HttpResponse(template.render(context, request))
@login_required
@permission_required('muestras.can_view_muestras_web')
def acciones_post(request):
    # Vista que redirigue la peticiÃ³n del usuario segÃºn el botÃ³n de acciÃ³n que haya pulsado y las muestras que haya seleccionado
    if request.method=="POST":
        muestras_seleccionadas = request.POST.getlist('muestra_id')
        if 'estudio' in request.POST:
            # Se guardan las muestras seleccionadas en la sesiÃ³n y se redirigue al usuario a la selecciÃ³n de un estudio
            if muestras_seleccionadas:
                request.session['muestras_estudio']=muestras_seleccionadas
                return redirect('seleccionar_estudio')
        elif 'eliminar' in request.POST:
            # Se eliminan las muestras seleccionadas con progreso AJAX
            if muestras_seleccionadas:
                def gen_eliminar():
                    muestras_a_procesar = list(Muestra.objects.filter(id__in=muestras_seleccionadas))
                    total = len(muestras_a_procesar)
                    yield _json_progress(0, total, 'start')
                    try:
                        for i, muestra in enumerate(muestras_a_procesar):
                            if Subposicion.objects.filter(muestra=muestra).exists():
                                subposicion = Subposicion.objects.get(muestra=muestra)
                                subposicion.muestra = None
                                subposicion.vacia = True
                                subposicion.save()
                            muestra.delete()
                            if _should_update(i, total):
                                yield _json_progress(i + 1, total, 'processing')
                        yield _json_progress(total, total, 'done', f'{total} muestras eliminadas correctamente')
                    except Exception as e:
                        yield _json_progress(0, total, 'error', str(e))
                return StreamingHttpResponse(gen_eliminar(), content_type='application/x-ndjson')
        elif 'envio' in request.POST:
            # Se guardan las muestras seleccionadas en la sesiÃ³n y se redirige al usuario a la agenda de envÃ­os
            if 'muestras_envio' in request.session:
                del request.session['muestras_envio']
            permitidas = []
            for mid in muestras_seleccionadas:
                estado = Muestra.objects.filter(id=mid).values_list('estado_actual', flat=True).first()
                if estado != 'DEST':
                    permitidas.append(mid)

            request.session['muestras_envio'] = permitidas
            return redirect('agenda')
        elif 'destruir' in request.POST:
            if muestras_seleccionadas:
                request.session['muestras_destruir'] = muestras_seleccionadas
                return redirect('destruir_muestras')
        elif 'cambio_posicion' in request.POST:
            # Se redirigue al usuario a la vista de cambio de posiciÃ³n de muestras
            # Permitir cambio de posiciÃ³n sin selecciÃ³n previa
            if 'muestras_cambio_posicion' in request.session:
                del request.session['muestras_cambio_posicion']
            if not muestras_seleccionadas:
                request.session['muestras_cambio_posicion'] = []
            else:
                request.session['muestras_cambio_posicion'] = muestras_seleccionadas
            return redirect('cambio_posicion')
        elif 'exportar_seleccionadas' in request.POST:
            # Exportar a Excel solo las muestras seleccionadas
            if muestras_seleccionadas:
                field_names = [f.name for f in Muestra._meta.local_fields if f.name not in ('id','estudio')]
                muestras_export = Muestra.objects.filter(id__in=muestras_seleccionadas)
                response = HttpResponse(content_type='application/ms-excel')
                response['Content-Disposition'] = 'attachment; filename="muestras_seleccionadas.xlsx"'
                wb = openpyxl.load_workbook(os.path.join(settings.BASE_DIR, 'geslab_project', 'globalstaticfiles', 'files', 'listado_muestras.xlsx'))
                ws = wb.active
                row_num = 2
                for muestra in muestras_export:
                    col_num = 1
                    for field in field_names:
                        value = muestra.__dict__[field]
                        if value is None:
                            value = ''
                        ws.cell(row_num, col_num).value = str(value)
                        col_num += 1
                    value = muestra.estudio.nombre_estudio if muestra.estudio else ''
                    ws.cell(row_num, col_num).value = str(value)
                    col_num += 1
                    value = muestra.posicion_completa()
                    if value is None:
                        value = ''
                    else:
                        value = value.split("-")
                        for columna in value:
                            ws.cell(row_num, col_num).value = str(columna)
                            col_num += 1
                    row_num += 1
                wb.save(response)
                return response             
    return redirect('muestras_todas')    
@login_required
@permission_required('muestras.can_view_muestras_web')
def detalles_muestra(request, nom_lab):
    # Vista que muestra los detalles de una muestra especÃ­fica, requiere permiso para ver muestras
    muestra = Muestra.objects.get(nom_lab=nom_lab)
    localizacion_actual = muestra.localizacion.order_by('-id').first()
    subposicion_actual = getattr(muestra, "subposicion", None)
    caja_layout = None

    if subposicion_actual and subposicion_actual.caja_id:
        caja = (
            Subposicion.objects
            .select_related("caja__rack__estante__congelador", "caja__bandeja")
            .get(pk=subposicion_actual.pk)
            .caja
        )
        subposiciones = list(caja.subposiciones.select_related("muestra").order_by("fila", "columna", "numero"))
        filas = sorted({sub.fila for sub in subposiciones if sub.fila})
        columnas = sorted(
            {sub.columna for sub in subposiciones if sub.columna},
            key=lambda value: (len(str(value)), str(value)),
        )
        celdas = {
            (str(sub.fila), str(sub.columna)): {
                "id": sub.id,
                "ocupada": not sub.vacia,
                "es_muestra_actual": sub.muestra_id == muestra.id,
                "muestra_nom_lab": sub.muestra.nom_lab if sub.muestra else "",
            }
            for sub in subposiciones
            if sub.fila and sub.columna
        }
        grid_rows = []
        for fila in filas:
            cells = []
            for columna in columnas:
                cell = celdas.get((str(fila), str(columna)), {
                    "ocupada": False,
                    "es_muestra_actual": False,
                    "muestra_nom_lab": "",
                })
                cells.append({
                    "fila": fila,
                    "columna": columna,
                    **cell,
                })
            grid_rows.append({
                "fila": fila,
                "cells": cells,
            })
        caja_layout = {
            "congelador": caja.rack.estante.congelador.congelador,
            "estante": caja.rack.estante.numero,
            "rack": caja.rack.numero,
            "bandeja": caja.bandeja.numero if caja.bandeja_id else None,
            "caja": caja.numero,
            "fila": subposicion_actual.fila,
            "columna": subposicion_actual.columna,
            "columnas": columnas,
            "grid_rows": grid_rows,
        }

    template = loader.get_template('detalles_muestra.html')
    context = {
        'muestra': muestra,
        'localizacion_actual': localizacion_actual,
        'caja_layout': caja_layout,
    }
    return HttpResponse(template.render(context, request))


@login_required
@permission_required('muestras.can_add_muestras_web')
@transaction.atomic
def añadir_muestras(request):
    # Vista para añadir varias alícuotas asociadas a un mismo individuo
    AliquotaFormSet = formset_factory(AliquotaForm, extra=0, can_delete=True)

    if request.method == 'POST':
        individuo_form = IndividuoEstudioForm(request.POST, prefix='individuo')
        datos_muestras_form = DatosMuestrasComunesForm(request.POST, prefix='datos')
        aliquota_formset = AliquotaFormSet(request.POST, prefix='aliquotas')

        if individuo_form.is_valid() and datos_muestras_form.is_valid() and aliquota_formset.is_valid():
            formularios_activos = [
                form for form in aliquota_formset
                if form.cleaned_data and not form.cleaned_data.get('DELETE')
            ]

            if not formularios_activos:
                messages.error(request, 'Añade al menos una alícuota antes de guardar.')
            else:
                nombres_vistos = set()
                subposiciones_vistas = set()
                hay_errores = False

                for form in formularios_activos:
                    nom_lab = form.cleaned_data['nom_lab'].strip().lower()
                    if nom_lab in nombres_vistos:
                        form.add_error('nom_lab', 'Este ID muestra está duplicado dentro del formulario.')
                        hay_errores = True
                    nombres_vistos.add(nom_lab)

                    subposicion_id = form.cleaned_data.get('subposicion_id')
                    if not subposicion_id:
                        continue
                    if subposicion_id in subposiciones_vistas:
                        form.add_error(None, 'La misma ubicación no se puede asignar a dos alícuotas.')
                        hay_errores = True
                        continue
                    subposiciones_vistas.add(subposicion_id)

                    try:
                        subposicion = Subposicion.objects.select_for_update().get(id=subposicion_id)
                        if not subposicion.vacia:
                            form.add_error(None, 'La ubicación seleccionada ya está ocupada.')
                            hay_errores = True
                    except Subposicion.DoesNotExist:
                        form.add_error(None, 'La ubicación seleccionada no existe.')
                        hay_errores = True

                if not hay_errores:
                    for form in formularios_activos:
                        _crear_muestra_desde_alicuota(
                            individuo_form,
                            form,
                            request.user,
                            datos_muestras_form.cleaned_data,
                        )
                    messages.success(request, 'Las alícuotas se han añadido correctamente.')
                    return redirect('muestras_todas')
    else:
        individuo_form = IndividuoEstudioForm(prefix='individuo')
        datos_muestras_form = DatosMuestrasComunesForm(prefix='datos')
        aliquota_formset = AliquotaFormSet(
            prefix='aliquotas',
            initial=[{'estado_actual': 'DISP'}],
        )

    materiales = list(Material.objects.filter(activo=True).order_by('nombre'))
    congeladores = (
        Congelador.objects
        .values('id', 'congelador', 'tipo_estructura')
        .order_by('congelador')
    )

    return render(request, 'añadir_muestras.html', {
        'individuo_form': individuo_form,
        'datos_muestras_form': datos_muestras_form,
        'aliquota_formset': aliquota_formset,
        'congeladores': congeladores,
        'materiales': materiales,
        'materiales_json': json.dumps(_material_config_payload(materiales)),
        'materiales_count': len(materiales),
    })

@login_required
@permission_required('muestras.can_delete_muestras_web')
def eliminar_muestra(request, nom_lab):
    # Vista para eliminar una muestra, requiere permiso para eliminar muestras
    muestra = get_object_or_404(Muestra, nom_lab=nom_lab)
    if Subposicion.objects.filter(muestra = muestra).exists():
        # Liberar la subposiciÃ³n asociada a la muestra antes de eliminarla
        subposicion = Subposicion.objects.get(muestra = muestra)
        subposicion.muestra = None
        subposicion.vacia = True
        subposicion.save()
    muestra.delete()
    messages.success(request,'Muestras eliminadas correctamente')
    return redirect('muestras_todas')
@login_required
@permission_required('muestras.can_add_muestras_web')
def descargar_plantilla_importacion_muestras(request, import_mode):
    mode = _get_import_mode(import_mode)
    workbook = _build_import_template_workbook(mode)
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    response = HttpResponse(
        output.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response['Content-Disposition'] = f'attachment; filename="plantilla_muestras_{mode}.xlsx"'
    return response


@login_required
@permission_required('muestras.can_add_muestras_web')
def upload_excel(request):
    # Mostrar confirmaciÃ³n despuÃ©s de validaciÃ³n con progreso
    if request.GET.get('mostrar_confirmacion') and 'confirmacion_pendiente' in request.session:
        datos_conf = request.session.pop('confirmacion_pendiente')
        for msg in datos_conf.get('mensajes', []):
            getattr(messages, msg['level'])(request, msg['text'])
        return render(request, datos_conf['template'], datos_conf.get('context', {}))
    # Vista para subir un archivo Excel con mÃºltiples muestras y asociarlas a un estudio y subposiciÃ³n desde el excel, requiere permiso para aÃ±adir muestras
    raw_import_mode = request.POST.get("import_mode") or request.GET.get("import_mode")
    import_mode = _get_import_mode(raw_import_mode)
    has_selected_mode = raw_import_mode in {choice[0] for choice in IMPORT_MODE_CHOICES}
    if request.method=="POST":
        form = UploadExcel(request.POST, request.FILES)
        # Si el usuario confirma, se crean en la base de datos los registros validos
        if 'confirmar' in request.POST:
            filas_validas = request.session.get('filas_validas',[])
            errores_sesion = request.session.get('errores', {})
            total = len(filas_validas)

            def gen_confirmar_muestras():
                yield _json_progress(0, total, 'start')
                try:
                    with transaction.atomic():
                        for i, datos in enumerate(filas_validas):
                            # Antes de crear, eliminar del dict los campos marcados como advertencia
                            fila_num = datos.get('fila')
                            advertencias = []
                            if fila_num:
                                advertencias = errores_sesion.get(fila_num, {}).get('advertencias', [])
                            for warn in advertencias:
                                if warn == 'fecha_incoherente':
                                    datos['fecha_extraccion'] = None
                                    datos['fecha_llegada'] = None
                                elif ':' in warn:
                                    _, campo = warn.split(':', 1)
                                    if campo in datos:
                                        datos[campo] = None

                            fecha_extraccion = None
                            fecha_llegada = None
                            if datos.get("fecha_extraccion"):
                                try:
                                    fecha_extraccion = date.fromisoformat(datos["fecha_extraccion"])
                                except Exception:
                                    fecha_extraccion = None
                            if datos.get("fecha_llegada"):
                                try:
                                    fecha_llegada = date.fromisoformat(datos["fecha_llegada"])
                                except Exception:
                                    fecha_llegada = None
                            estudio = None
                            if datos.get("estudio"):
                                estudio = Estudio.objects.get(id=datos["estudio"])

                            muestra = Muestra.objects.create(
                                id_individuo=datos.get("id_individuo"),
                                nom_lab=datos["nom_lab"],
                                material=datos.get("material_obj"),
                                volumen_actual=datos.get("volumen_actual"),
                                unidad_volumen=datos.get("unidad_volumen"),
                                concentracion_actual=datos.get("concentracion_actual"),
                                unidad_concentracion=datos.get("unidad_concentracion"),
                                masa_actual=datos.get("masa_actual"),
                                unidad_masa=datos.get("unidad_masa"),
                                fecha_extraccion=fecha_extraccion,
                                fecha_llegada=fecha_llegada,
                                observaciones=datos.get("observaciones"),
                                centro_procedencia=datos.get("centro_procedencia"),
                                estado_actual=datos.get("estado_actual") or None,
                                estudio=estudio,
                            )

                            subposicion_id = datos.get("subposicion_id")
                            if subposicion_id:
                                try:
                                    subposicion = Subposicion.objects.select_for_update().get(id=subposicion_id)
                                except Subposicion.DoesNotExist:
                                    subposicion = None
                                if subposicion:
                                    localizacion = _localizacion_desde_subposicion(muestra, subposicion)
                                    subposicion.vacia = False
                                    subposicion.muestra = muestra
                                    subposicion.save()
                                    historial_localizaciones.objects.create(
                                        muestra=muestra,
                                        localizacion=localizacion,
                                        fecha_asignacion=timezone.now(),
                                        usuario_asignacion=request.user
                                    )

                            if datos.get("estudio"):
                                historial_estudios.objects.create(
                                    muestra=muestra,
                                    estudio=estudio,
                                    fecha_asignacion=timezone.now(),
                                    usuario_asignacion=request.user
                                )
                            if _should_update(i, total):
                                yield _json_progress(i + 1, total, 'processing')
                    yield _json_progress(total, total, 'done', 'Muestras importadas correctamente', '/muestras/')
                except Exception as e:
                    yield _json_progress(0, total, 'error', str(e))

            return StreamingHttpResponse(gen_confirmar_muestras(), content_type='application/x-ndjson')
        # Si el usuario cancela, no se aÃ±ade nada a la base de datos
        elif 'cancelar' in request.POST:
            messages.error(request,'Las muestras no se han aÃ±adido')
            return redirect('muestras_todas')
        
        # Si hay un archivo excel subido, se procesa
        elif 'excel_file' in request.FILES:
            # Limpiar sesiÃ³n residual de uploads anteriores
            if 'columnas_adicionales' in request.session:
                del request.session['columnas_adicionales']
            if form.is_valid():
                # Leer excel y preparar columnas 
                excel_file = request.FILES['excel_file']
                excel_bytes = excel_file.read()
                request.session['excel_file_name'] = excel_file.name
                request.session['excel_file_base64']= base64.b64encode(excel_bytes).decode()
                request.session['import_mode'] = import_mode
                excel_stream = io.BytesIO(excel_bytes)
                
                # Intentar leer el archivo Excel
                try:
                    df = pd.read_excel(excel_stream)
                except Exception as e:
                    return render(request, 'upload_excel_details.html', {
                        'form': form,
                        'error': f'âŒ Error al leer el archivo Excel: {str(e)}',
                        'import_mode': import_mode,
                        'import_mode_ui': _get_import_mode_ui(import_mode),
                    })
                
                # Validar que tenga al menos una fila de datos
                if df.empty or len(df) == 0:
                    return render(request, 'upload_excel_details.html', {
                        'form': form,
                        'error': 'âŒ Error de formato: El archivo Excel estÃ¡ vacÃ­o o no contiene filas de datos.',
                        'import_mode': import_mode,
                        'import_mode_ui': _get_import_mode_ui(import_mode),
                    })
                
                rename_columns = _get_import_rename_columns(import_mode)
                # Validar columnas
                columnas_esperadas = set(_get_import_headers(import_mode))
                columnas_existentes = set(df.columns)
                columnas_faltantes = columnas_esperadas - columnas_existentes
                if columnas_faltantes:
                    columnas_str = ', '.join(sorted(columnas_faltantes))
                    return render(request, 'upload_excel_details.html', {
                        'form': form,
                        'error': f'âŒ Error de formato: El archivo Excel no contiene las siguientes columnas esperadas: {columnas_str}',
                        'import_mode': import_mode,
                        'import_mode_ui': _get_import_mode_ui(import_mode),
                    })
                
                columnas_adicionales = columnas_existentes - columnas_esperadas
                if columnas_adicionales:
                    columnas_adicionales_str = ', '.join(sorted(columnas_adicionales))
                    request.session['columnas_adicionales'] = columnas_adicionales_str
                
                df.rename(columns=rename_columns, inplace=True)
                # Funciones para normalizar las columnas del excel
                def norm(value):
                    if value is None or pd.isna(value):
                        return None

                    if isinstance(value, str):
                        value = value.strip()
                        return value if value != "" else None

                    return value
                
                def norm_code(value):
                    if value is None or pd.isna(value):
                        return None

                    if isinstance(value, float) and value.is_integer():
                        return str(int(value))

                    return str(value).strip()

                # Definir campos obligatorios para una muestra
                obligatorios = ["nom_lab", "id_material", "estado_actual"]
                if import_mode == IMPORT_MODE_CONGELADOR_80:
                    obligatorios.extend(["congelador", "estante", "rack", "bandeja", "caja", "fila", "columna"])
                elif import_mode in {IMPORT_MODE_CONGELADOR_20, IMPORT_MODE_NEVERA}:
                    obligatorios.extend(["congelador", "estante", "rack", "caja", "fila", "columna"])

                cache = {
                    'subposiciones': {
                        _subposicion_cache_key(c): c
                        for c in Subposicion.objects.select_related('caja__bandeja__rack__estante__congelador')
                    },

                    'estudios':{e:Estudio.objects.get(id=e)
                                for e in Estudio.objects.values_list('id', flat=True)},

                    'materiales': {
                        nombre.casefold(): material
                        for material in Material.objects.filter(activo=True)
                        for nombre in [material.nombre]
                    },
                    'tipos_recurso': {
                        nombre.casefold(): tipo
                        for nombre, tipo in Congelador.objects.values_list('congelador', 'tipo_estructura')
                    },

                    'muestras_existentes': set(Muestra.objects.values_list('nom_lab',flat=True))
                }

                total_filas = len(df)

                def gen_validar_muestras():
                    errores = {}
                    filas_validas = []
                    numero_registros = 0
                    nom_lab_excel = set()

                    yield _json_progress(0, total_filas, 'start')

                    # Recorrer el df para detectar errores y normalizar
                    for idx, row in df.iterrows():
                        numero_registros += 1
                        fila = idx + 2 
                        errores[fila]={"bloqueantes":[],"advertencias":[]}
                        datos = {
                            "id_individuo":norm(row['id_individuo']),
                            "nom_lab":norm(row['nom_lab']),
                            "id_material":norm(row['id_material']),
                            "volumen_actual":norm_code(row['volumen_actual']),
                            "unidad_volumen":norm(row['unidad_volumen']),
                            "concentracion_actual":norm_code(row['concentracion_actual']),
                            "unidad_concentracion":norm(row['unidad_concentracion']),
                            "masa_actual":norm_code(row['masa_actual']),
                            "unidad_masa":norm(row['unidad_masa']),
                            "fecha_extraccion":norm(row['fecha_extraccion']),
                            "fecha_llegada":norm(row['fecha_llegada']),
                            "observaciones":norm(row['observaciones']),
                            "centro_procedencia":norm(row['centro_procedencia']),
                            "estado_actual":norm(row['estado_actual']),
                            "congelador":(norm_code(row['congelador']) or None) if 'congelador' in row else None,
                            "estante":norm_code(row['estante']) if 'estante' in row else None,
                            "rack":(norm_code(row['rack']) or None) if 'rack' in row else None,
                            "bandeja":(norm_code(row['bandeja']) or None) if 'bandeja' in row else None,
                            "caja":(norm_code(row['caja']) or None) if 'caja' in row else None,
                            "fila":(norm_code(row['fila']) or None) if 'fila' in row else None,
                            "columna":(norm_code(row['columna']) or None) if 'columna' in row else None,
                            "estudio":norm_code(row['estudio'])  
                        }
                        datos["material_obj"] = None
                        # Normalizar a minÃºsculas los campos textuales relevantes
                        if datos['congelador'] is not None:
                            datos['congelador'] = str(datos['congelador']).lower()
                        if datos['rack'] is not None:
                            datos['rack'] = str(datos['rack']).lower()
                        if datos['bandeja'] is not None:
                            datos['bandeja'] = str(datos['bandeja']).lower()
                        if datos['caja'] is not None:
                            datos['caja'] = str(datos['caja']).lower()
                        if datos['fila'] is not None:
                            datos['fila'] = str(datos['fila']).lower()
                        if datos['columna'] is not None:
                            datos['columna'] = str(datos['columna']).lower()

                        for campo in obligatorios:
                            if not datos.get(campo):
                                errores[fila]["bloqueantes"].append(f"campo_obligatorio_vacio:{campo}")

                        # Validar que ningÃºn campo contenga el carÃ¡cter punto y coma (;)
                        campos_a_validar = ['id_individuo', 'nom_lab', 'id_material', 'unidad_volumen', 
                                           'unidad_concentracion', 'unidad_masa', 'observaciones',
                                           'centro_procedencia', 'estado_actual', 'congelador',
                                           'rack', 'bandeja', 'caja', 'fila', 'columna']
                        for campo in campos_a_validar:
                            valor = datos.get(campo)
                            if valor and isinstance(valor, str) and ';' in valor:
                                errores[fila]["bloqueantes"].append(f"caracter_invalido_semicolon:{campo}")

                        # Comprobar si los campos estan en el formato correcto (numÃ©ricos)
                        for campo in ['volumen_actual', 'concentracion_actual', 'masa_actual']:
                            if datos[campo] != None:
                                try:
                                    valor_numerico = float(datos[campo])
                                    # Validar que sea positivo o cero (â‰¥0)
                                    if valor_numerico < 0:
                                        errores[fila]["advertencias"].append(f"valor_negativo:{campo}")
                                        datos[campo] = None  # No asignar valores negativos
                                    else:
                                        datos[campo] = valor_numerico
                                except (TypeError, ValueError):
                                    errores[fila]["advertencias"].append(f"formato_incorrecto:{campo}")
                                    datos[campo] = None  # No asignar valores con formato incorrecto
                        
                        # Validar estado_actual: debe ser uno de los valores permitidos (case-insensitive)
                        estado_actual_validos = {
                            'disponible': 'DISP',
                            'enviada': 'ENV',
                            'parcialmente enviada': 'PENV',
                            'enviada parcialmente': 'PENV',
                            'destruida': 'DEST',
                            'disp': 'DISP',
                            'env': 'ENV',
                            'penv': 'PENV',
                            'dest': 'DEST'
                        }
                        if datos.get('estado_actual'):
                            estado_normalizado = str(datos['estado_actual']).strip().lower()
                            if estado_normalizado in estado_actual_validos:
                                datos['estado_actual'] = estado_actual_validos[estado_normalizado]
                            else:
                                errores[fila]["advertencias"].append("estado_actual_invalido")
                                datos['estado_actual'] = None

                        material_nombre = datos.get("id_material")
                        if material_nombre:
                            material_obj = cache["materiales"].get(str(material_nombre).casefold())
                            if material_obj:
                                datos["material_obj"] = material_obj
                                datos["id_material"] = material_obj.nombre
                            else:
                                errores[fila]["bloqueantes"].append("material_no_existe")

                        material_obj = datos.get("material_obj")
                        if material_obj:
                            if material_obj.usa_volumen:
                                if datos.get("unidad_volumen") and datos["unidad_volumen"] not in material_obj.unidades_volumen_list():
                                    errores[fila]["bloqueantes"].append("unidad_no_valida:unidad_volumen")
                            elif datos.get("volumen_actual") is not None or datos.get("unidad_volumen"):
                                errores[fila]["bloqueantes"].append("campo_no_permitido:volumen_actual")
                                if datos.get("unidad_volumen"):
                                    errores[fila]["bloqueantes"].append("campo_no_permitido:unidad_volumen")

                            if material_obj.usa_concentracion:
                                if datos.get("unidad_concentracion") and datos["unidad_concentracion"] not in material_obj.unidades_concentracion_list():
                                    errores[fila]["bloqueantes"].append("unidad_no_valida:unidad_concentracion")
                            elif datos.get("concentracion_actual") is not None or datos.get("unidad_concentracion"):
                                errores[fila]["bloqueantes"].append("campo_no_permitido:concentracion_actual")
                                if datos.get("unidad_concentracion"):
                                    errores[fila]["bloqueantes"].append("campo_no_permitido:unidad_concentracion")

                            if material_obj.usa_unidades_bloque:
                                if datos.get("unidad_masa") and datos["unidad_masa"] not in material_obj.unidades_recuento_list():
                                    errores[fila]["bloqueantes"].append("unidad_no_valida:unidad_masa")
                            elif datos.get("masa_actual") is not None or datos.get("unidad_masa"):
                                errores[fila]["bloqueantes"].append("campo_no_permitido:masa_actual")
                                if datos.get("unidad_masa"):
                                    errores[fila]["bloqueantes"].append("campo_no_permitido:unidad_masa")

                        if datos.get("congelador"):
                            tipo_recurso = cache["tipos_recurso"].get(str(datos["congelador"]).casefold())
                            if tipo_recurso and import_mode != IMPORT_MODE_SIN_UBICACION and tipo_recurso != import_mode:
                                errores[fila]["bloqueantes"].append("tipo_recurso_incorrecto")
                            
                        # Validar y normalizar fechas (esperado formato dÃ­a-mes-aÃ±o o fecha Excel)
                        for campo in ['fecha_extraccion', 'fecha_llegada']:
                            if datos[campo] != None:
                                try:
                                    # Forzar interpretaciÃ³n dÃ­a/mes/aÃ±o cuando sea una cadena
                                    fecha = pd.to_datetime(datos[campo], dayfirst=True, errors='raise')
                                    datos[campo] = fecha.date().isoformat()
                                except Exception:
                                    errores[fila]["bloqueantes"].append(f"fecha_invalida:{campo}")

                        # Si ambas fechas estÃ¡n presentes, comprobar coherencia: llegada >= extracciÃ³n
                        if datos.get('fecha_extraccion') and datos.get('fecha_llegada'):
                            try:
                                fe_extr = date.fromisoformat(datos['fecha_extraccion'])
                                fe_lleg = date.fromisoformat(datos['fecha_llegada'])
                                if fe_lleg < fe_extr:
                                    errores[fila]["bloqueantes"].append('fecha_incoherente')
                            except Exception:
                                # Si por alguna razÃ³n no se pueden convertir, marcar como fecha invÃ¡lida y bloquear
                                errores[fila]["bloqueantes"].append('fecha_invalida:fecha_extraccion')
                                    
                        # Comprobar si hay duplicados entre las muestras dentro del excel o en la base de datos
                        nom_lab = datos["nom_lab"]

                        if nom_lab in cache["muestras_existentes"]:
                            errores[fila]["bloqueantes"].append("muestra_duplicada_bd")
                        if nom_lab in nom_lab_excel:
                            errores[fila]["bloqueantes"].append("muestra_duplicada_excel")
                        else:
                            nom_lab_excel.add(nom_lab)
                        
                        # Comprobar si el estudio existe en la base de datos
                        estudio_id = datos.get("estudio")
                        if estudio_id:
                            try:
                                estudio_id = int(estudio_id)  # Convertir a integer para buscar en BD
                                estudio = cache["estudios"].get(estudio_id)
                                if not estudio:
                                    errores[fila]["advertencias"].append("estudio_no_existe")
                                    datos["estudio"] = None
                                else:
                                    datos["estudio"] = estudio.id
                            except (ValueError, TypeError):
                                errores[fila]["advertencias"].append("estudio_no_existe")
                                datos["estudio"] = None

                        # Comprobar si la localizaciÃ³n estÃ¡ ocupada o existe Ãºnicamente si se han proporcionado datos de posiciÃ³n
                        if import_mode != IMPORT_MODE_SIN_UBICACION and (
                            datos.get("congelador") or datos.get("estante") or datos.get("rack") or datos.get("bandeja") or datos.get("caja") or datos.get("fila") or datos.get("columna")
                        ):
                            key = (
                                datos["congelador"],
                                datos["estante"],
                                datos["rack"],
                                datos["bandeja"] or "",
                                datos["caja"],
                                datos["fila"],
                                datos["columna"],
                            )

                            subposicion = cache["subposiciones"].get(key)

                            if not subposicion:
                                errores[fila]["bloqueantes"].append("localizacion_no_existe")
                            elif not subposicion.vacia:
                                errores[fila]["bloqueantes"].append("localizacion_ocupada")
                            else:
                                datos["subposicion_id"] = subposicion.id
                        else:
                            # No se proporcionÃ³ posiciÃ³n; la muestra quedarÃ¡ sin localizaciÃ³n asignada
                            datos["subposicion_id"] = None

                        # Detectar campos opcionales vacios (solo 'nom_lab' es obligatorio)
                        opcionales = [
                            'id_individuo',
                            'id_material',
                            'volumen_actual',
                            'unidad_volumen',
                            'concentracion_actual',
                            'unidad_concentracion',
                            'masa_actual',
                            'unidad_masa',
                            'fecha_extraccion',
                            'fecha_llegada',
                            'observaciones',
                            'estado_actual',
                            'estudio',
                            'centro_procedencia',
                            # PosiciÃ³n (tratar como opcional; si se proporcionan, se validan arriba)
                            'congelador',
                            'estante',
                            'rack',
                            'bandeja',
                            'caja',
                            'fila',
                            'columna'
                        ]
                        for campo in opcionales:
                            if datos.get(campo) is None:
                                errores[fila]["advertencias"].append(f"campo_vacio:{campo}")

                        # Registrar filas validas
                        if not errores[fila]["bloqueantes"]:
                            datos['fila'] = fila
                            filas_validas.append(datos)

                        if _should_update(idx, total_filas):
                            yield _json_progress(idx + 1, total_filas, 'processing')
                    
                    # Guardar en la sesiÃ³n las filas validas y los errores detectados
                    request.session['filas_validas']=filas_validas
                    request.session['errores'] = errores

                    # Obtener configuraciÃ³n de mensajes para muestras
                    msg_config = get_upload_messages('muestras')
                    
                    # Contar errores
                    numero_errores_bloqueantes = 0
                    numero_errores_advertencia = 0
                    for fila in errores:
                        if errores[fila]['bloqueantes']:
                            numero_errores_bloqueantes += 1
                        if errores[fila]["advertencias"]:
                            numero_errores_advertencia += 1

                    # Construir lista de mensajes para confirmaciÃ³n
                    mensajes = []
                    mensajes.append({'level': 'info', 'text': f'{msg_config["titulo_inicial"]} {numero_registros} registros.'})

                    if numero_errores_bloqueantes == 0 and numero_errores_advertencia == 0:
                        mensajes.append({'level': 'success', 'text': msg_config['sin_errores']})
                    else:
                        if numero_errores_advertencia > 0:
                            msg = msg_config['con_advertencias'].format(count=numero_errores_advertencia)
                            mensajes.append({'level': 'warning', 'text': msg})
                        if numero_errores_bloqueantes > 0:
                            msg = msg_config['con_bloqueantes'].format(count=numero_errores_bloqueantes)
                            mensajes.append({'level': 'error', 'text': msg})
                    
                    # Mostrar mensaje de columnas extras si existen
                    columnas_extras_str = request.session.get('columnas_adicionales', '')
                    tiene_columnas_extras = bool(columnas_extras_str)
                    numero_columnas_extras = len(columnas_extras_str.split(', ')) if columnas_extras_str else 0
                    if tiene_columnas_extras:
                        msg = msg_config['columnas_extras'].format(count=numero_columnas_extras, detalles=columnas_extras_str)
                        mensajes.append({'level': 'warning', 'text': msg})

                    # Guardar datos de confirmaciÃ³n en sesiÃ³n para mostrar tras redirecciÃ³n
                    request.session['confirmacion_pendiente'] = {
                        'template': 'confirmacion_upload.html',
                        'context': {
                            'numero_errores_bloqueantes': numero_errores_bloqueantes,
                            'numero_errores_advertencia': numero_errores_advertencia,
                            'tiene_columnas_extras': tiene_columnas_extras,
                            'numero_columnas_extras': numero_columnas_extras,
                            'columnas_extras_str': columnas_extras_str
                        },
                        'mensajes': mensajes
                    }
                    request.session.save()

                    yield _json_progress(total_filas, total_filas, 'done', 'ValidaciÃ³n completada', request.path + '?mostrar_confirmacion=1')

                return StreamingHttpResponse(gen_validar_muestras(), content_type='application/x-ndjson')
        # Si se solicita un excel de errores, este se rellena en base a los errores detectados durante la validaciÃ³n 
        elif 'excel_errores' in request.POST:
                    # Leer los errores y el excel de la sesiÃ³n
                    errores = request.session.get('errores',[])
                    columnas_adicionales_str = request.session.get('columnas_adicionales', '')
                    # Convertir string de columnas adicionales a un set para procesamiento
                    columnas_adicionales = set()
                    if columnas_adicionales_str:
                        columnas_adicionales = set(col.strip() for col in columnas_adicionales_str.split(','))
                    
                    excel_bytes = base64.b64decode(request.session.get('excel_file_base64'))
                    excel_file = io.BytesIO(excel_bytes)
                    wb = openpyxl.load_workbook(excel_file)
                    ws = wb.active
                    # Definir los estilos para pintar el excel usando configuraciÃ³n centralizada
                    colors = get_excel_colors()
                    FILL_ERROR_CELL = PatternFill("solid", fgColor=colors['error_cell'])
                    FILL_WARN_CELL  = PatternFill("solid", fgColor=colors['warning_cell'])
                    FILL_EXTRA_COL  = PatternFill("solid", fgColor=colors['extra_column'])
                    # Diccionario de mensajes
                    MENSAJES_ERROR = {
                        "campo_obligatorio_vacio": "Campo obligatorio vacÃ­o",
                        "campo_no_permitido": "Este campo no aplica al material indicado",
                        "formato_incorrecto": "Formato incorrecto",
                        "valor_negativo": "El valor debe ser positivo (â‰¥0)",
                        "unidad_no_valida": "La unidad indicada no es válida para el material",
                        "estado_actual_invalido": "Estado debe ser: Disponible, Enviada, Enviada parcialmente o Destruida",
                        "material_no_existe": "El material indicado no existe o no estÃ¡ activo",
                        "tipo_recurso_incorrecto": "El recurso indicado no coincide con el tipo de plantilla seleccionado",
                        "fecha_invalida": "Fecha invÃ¡lida (Formato correcto: DD-MM-AAAA)",
                        "fecha_incoherente": "Fecha llegada anterior a fecha de extracciÃ³n",
                        "muestra_duplicada_bd": "La muestra ya existe en la base de datos",
                        "muestra_duplicada_excel": "Muestra duplicada dentro del Excel",
                        "localizacion_ocupada": "La subposiciÃ³n ya estÃ¡ ocupada",
                        "localizacion_no_existe": "La localizaciÃ³n no existe",
                        "campo_vacio": "Campo opcional vacÃ­o",
                        "estudio_no_existe": "El estudio no existe (se puede asignar despuÃ©s)",
                        "caracter_invalido_semicolon": "El carÃ¡cter ';' no estÃ¡ permitido en este campo",
                    }
                    # Diccionario de columnas del excel
                    columnas_excel = {}
                    import_mode = _get_import_mode(request.session.get('import_mode'))
                    rename_columns = _get_import_rename_columns(import_mode)
                    for cell in ws[1]:
                        if cell.value in rename_columns:
                            columnas_excel[rename_columns[cell.value]] = cell.column
                    # AÃ±adir la columna de errores
                    col_errores = ws.max_column + 1
                    ws.cell(row=1, column=col_errores, value="Errores")
                    # Mapear errores sin campo a sus columnas especÃ­ficas (para muestras)
                    error_campo_map = {
                        "muestra_duplicada_bd": "nom_lab",
                        "muestra_duplicada_excel": "nom_lab",
                        "localizacion_no_existe": "columna",
                        "localizacion_ocupada": "columna"
                    }
                    # Recorrer filas con errores 
                    for fila, info in errores.items():
                        has_error = bool(info.get("bloqueantes", []))
                        has_warn = bool(info.get("advertencias", []))
                        if not has_error and not has_warn:
                            continue

                        # Colorear celdas especÃ­ficas y construir mensajes
                        mensajes = []
                        for err in info.get("bloqueantes", []):
                            if ":" in err:
                                tipo, campo = err.split(":")
                                msg = f"(Error) {MENSAJES_ERROR[tipo]}"
                                if msg not in mensajes:
                                    mensajes.append(msg)
                                if campo in columnas_excel:
                                    col = columnas_excel[campo]
                                    celda = ws.cell(row=int(fila), column=col)
                                    celda.fill = FILL_ERROR_CELL
                            else:
                                campo = error_campo_map.get(err)
                                msg = f"(Error) {MENSAJES_ERROR[err]}"
                                if msg not in mensajes:
                                    mensajes.append(msg)
                                if campo and campo in columnas_excel:
                                    col_err = columnas_excel[campo]
                                    ws.cell(row=int(fila), column=col_err).fill = FILL_ERROR_CELL
                        for warn in info.get("advertencias", []):
                            if ":" in warn:
                                tipo, campo = warn.split(":")
                                mensaje_warn = MENSAJES_ERROR.get(tipo, tipo)
                                msg = f"(Advertencia) {mensaje_warn}"
                                if msg not in mensajes:
                                    mensajes.append(msg)
                                if campo in columnas_excel:
                                    col = columnas_excel[campo]
                                    celda = ws.cell(row=int(fila), column=col)
                                    celda.fill = FILL_WARN_CELL
                            else:
                                mensaje_warn = MENSAJES_ERROR.get(warn, warn)
                                msg = f"(Advertencia) {mensaje_warn}"
                                if msg not in mensajes:
                                    mensajes.append(msg)
                        ws.cell(row=int(fila), column=col_errores, value="\n".join(mensajes))

                    # Pintar columnas extras con color de columna extra
                    if columnas_adicionales:
                        for col_name in columnas_adicionales:
                            # Encontrar el nÃºmero de columna del Excel original para esta columna extra
                            for cell in ws[1]:
                                if cell.value == col_name:
                                    col_num = cell.column
                                    # Pintar el encabezado
                                    header_cell = ws.cell(row=1, column=col_num)
                                    header_cell.fill = FILL_EXTRA_COL
                                    # Pintar todas las celdas de datos en la columna
                                    for row in range(2, ws.max_row + 1):
                                        ws.cell(row=row, column=col_num).fill = FILL_EXTRA_COL
                                    break

                    # Rertornar el excel de errores    
                    output = io.BytesIO()    
                    wb.save(output)
                    wb.close()
                    response = HttpResponse(output.getvalue(),content_type='application/ms-excel')
                    response['Content-Disposition'] = 'attachment; filename="listado_errores.xlsx"'
                    return response        
 
    else:
        form = UploadExcel()

    if not has_selected_mode:
        return render(request, 'upload_excel_select_type.html', {
            'import_mode_cards': _get_import_mode_cards(),
        })

    return render(request, 'upload_excel_details.html', {
        'form': form,
        'import_mode': import_mode,
        'import_mode_ui': _get_import_mode_ui(import_mode),
    })
@login_required
@permission_required('muestras.can_change_muestras_web')
def cambio_posicion(request):
    # Mostrar confirmaciÃ³n despuÃ©s de validaciÃ³n con progreso
    if request.GET.get('mostrar_confirmacion') and 'confirmacion_pendiente' in request.session:
        datos_conf = request.session.pop('confirmacion_pendiente')
        for msg in datos_conf.get('mensajes', []):
            getattr(messages, msg['level'])(request, msg['text'])
        return render(request, datos_conf['template'], datos_conf.get('context', {}))
    # Vista para cambiar la posiciÃ³n de mÃºltiples muestras a partir de un archivo Excel, requiere permiso para cambiar muestras
    if request.method=="POST":
        form = UploadExcel(request.POST, request.FILES)
        # Si el usuario confirma, se guardan las muestras en una nueva posiciÃ³n, vaciando la posicion antigua
        if 'confirmar' in request.POST:
            filas_validas = request.session.get('filas_validas',[])
            total = len(filas_validas)

            def gen_confirmar_cambio():
                yield _json_progress(0, total, 'start')
                try:
                    with transaction.atomic():
                        for i, datos in enumerate(filas_validas):
                            subposicion = Subposicion.objects.select_for_update().get(id=datos["subposicion_id"])
                            muestra = Muestra.objects.get(nom_lab=datos['nom_lab'])

                            localizacion = _localizacion_desde_subposicion(muestra, subposicion)

                            antigua_id = datos.get('subposicion_antigua')
                            if antigua_id:
                                try:
                                    subposicion_antigua = Subposicion.objects.select_for_update().get(id=antigua_id)
                                    subposicion_antigua.vacia = True
                                    subposicion_antigua.muestra = None
                                    subposicion_antigua.save()
                                except Subposicion.DoesNotExist:
                                    pass

                            subposicion.vacia = False
                            subposicion.muestra = muestra
                            subposicion.save()

                            historial_localizaciones.objects.create(
                                muestra=muestra,
                                localizacion=localizacion,
                                fecha_asignacion=timezone.now(),
                                usuario_asignacion=request.user
                            )
                            if _should_update(i, total):
                                yield _json_progress(i + 1, total, 'processing')
                    yield _json_progress(total, total, 'done', 'Posiciones actualizadas correctamente', '/muestras/')
                except Exception as e:
                    yield _json_progress(0, total, 'error', str(e))

            return StreamingHttpResponse(gen_confirmar_cambio(), content_type='application/x-ndjson')

        # Si el usuario cancela, no se hace nada
        elif 'cancelar' in request.POST:
            messages.error(request,'Las muestras no se han cambiado de posiciÃ³n.')
            return redirect('muestras_todas')

        elif 'descargar_excel_cambio_posicion' in request.POST:
            # Descargar plantilla de cambio de posiciÃ³n con IDs precargados si hay selecciÃ³n
            muestras = request.session.get('muestras_cambio_posicion', [])
            plantilla_path = os.path.join(settings.BASE_DIR, 'geslab_project', 'globalstaticfiles', 'files', 'plantilla_cambio_posicion.xlsx')
            if not os.path.exists(plantilla_path):
                return HttpResponse("La plantilla no se encuentra disponible.", status=404)
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename="plantilla_cambio_posicion.xlsx"'
            wb = openpyxl.load_workbook(plantilla_path)
            ws = wb.active
            row_num = 2
            for muestra_id in muestras:
                sample = Muestra.objects.filter(id=muestra_id).first()
                if not sample:
                    continue
                ws.cell(row_num, 1).value = str(sample.nom_lab)
                row_num += 1
            wb.save(response)
            return response
        
        elif 'excel_file' in request.FILES:
            # Limpiar sesiÃ³n residual de uploads anteriores
            if 'columnas_adicionales' in request.session:
                del request.session['columnas_adicionales']
            if form.is_valid():
                # Leer excel y preparar columnas 
                excel_file = request.FILES['excel_file']
                excel_bytes = excel_file.read()
                request.session['excel_file_name'] = excel_file.name
                request.session['excel_file_base64']= base64.b64encode(excel_bytes).decode()
                excel_stream = io.BytesIO(excel_bytes)
                
                # Intentar leer el archivo Excel
                try:
                    df = pd.read_excel(excel_stream)
                except Exception as e:
                    return render(request, 'upload_excel_cambio_posicion.html', {'form': form, 'error': f'âŒ Error al leer el archivo Excel: {str(e)}'})
                
                rename_columns = {
                    'Nombre Laboratorio': 'nom_lab',
                    'Congelador': 'congelador', 
                    'Estante': 'estante',
                    'PosiciÃ³n del rack en el estante': 'posicion_rack_estante',
                    'Rack': 'rack',
                    'PosiciÃ³n de la bandeja en el rack': 'posicion_bandeja_rack',
                    'Bandeja': 'bandeja',
                    'PosiciÃ³n de la caja en la bandeja': 'posicion_caja_rack',
                    'Caja': 'caja',
                    'Fila': 'fila',
                    'Columna': 'columna',
                }
                
                # Normalizar nombres de columnas a minÃºsculas para comparaciÃ³n insensible a mayÃºsculas/minÃºsculas
                df.columns = df.columns.str.lower()
                rename_columns_normalized = {k.lower(): v for k, v in rename_columns.items()}
                
                # Validar que el Excel tenga las columnas esperadas
                columnas_esperadas = set(rename_columns_normalized.keys())
                columnas_existentes = set(df.columns)
                columnas_faltantes = columnas_esperadas - columnas_existentes
                
                if columnas_faltantes:
                    columnas_str = ', '.join(sorted(columnas_faltantes))
                    return render(request, 'upload_excel_cambio_posicion.html', {'form': form, 'error': f'âŒ Error de formato: El archivo Excel no contiene las siguientes columnas esperadas: {columnas_str}'})
                
                # Validar que el Excel no estÃ© vacÃ­o
                if df.empty or len(df) == 0:
                    return render(request, 'upload_excel_cambio_posicion.html', {'form': form, 'error': 'âŒ Error de formato: El archivo Excel estÃ¡ vacÃ­o o no contiene filas de datos.'})
                
                # Validar columnas adicionales
                columnas_adicionales = columnas_existentes - columnas_esperadas
                extra_columns = False
                columnas_adicionales_str = ''
                if columnas_adicionales:
                    columnas_adicionales_str = ', '.join(sorted(columnas_adicionales))
                    request.session['columnas_adicionales'] = columnas_adicionales_str
                    extra_columns = True
                
                df.rename(columns=rename_columns_normalized, inplace=True)
                # Funciones para normalizar las columnas del excel
                def norm(value):
                    if value is None or pd.isna(value):
                        return None

                    if isinstance(value, str):
                        value = value.strip()
                        return value if value != "" else None

                    return value
                
                def norm_code(value):
                    if value is None or pd.isna(value):
                        return None

                    if isinstance(value, float) and value.is_integer():
                        return str(int(value)).lower()

                    return str(value).strip().lower()
                # Carga de datos previos y creaciÃ³n de estructuras 
                cache = {
                    'subposiciones': {
                        _subposicion_cache_key(c): c
                        for c in Subposicion.objects.select_related('caja__bandeja__rack__estante__congelador')
                    },
                    'posiciones_actuales': {
                        (p.muestra.nom_lab.lower()): p.id 
                        for p in Subposicion.objects.all() if p.muestra != None
                    },

                    'muestras_existentes': set(m.lower() for m in Muestra.objects.values_list('nom_lab',flat=True))
                }

                total_filas = len(df)

                def gen_validar_cambio():
                    filas_validas = []
                    errores = {}
                    nom_lab_excel = set()
                    numero_registros = 0

                    yield _json_progress(0, total_filas, 'start')

                    # Recorrer el df para detectar errores y normalizar
                    for idx, row in df.iterrows():
                        numero_registros += 1
                        fila = idx + 2 
                        errores[fila]={"bloqueantes":[]}
                        datos = {
                            "nom_lab":norm(row['nom_lab']),
                            "congelador":(norm_code(row['congelador']) or None),
                            "estante":norm_code(row['estante']),
                            "posicion_rack_estante":norm_code(row['posicion_rack_estante']),
                            "rack":(norm_code(row['rack']) or None),
                            "posicion_bandeja_rack":norm_code(row['posicion_bandeja_rack']),
                            "bandeja":(norm_code(row['bandeja']) or None),
                            "posicion_caja_rack":norm_code(row['posicion_caja_rack']),
                            "caja":(norm_code(row['caja']) or None),
                            "fila":(norm_code(row['fila']) or None),
                            "columna":(norm_code(row['columna']) or None), 
                        }
                        # Comprobar si los campos obligatorios estÃ¡n rellenados
                        obligatorios = ["nom_lab", "congelador", "estante", "posicion_rack_estante", "rack", "posicion_bandeja_rack", "bandeja", "caja", "posicion_caja_rack", "fila", "columna"]

                        for campo in obligatorios:
                            if not datos.get(campo):
                                errores[fila]["bloqueantes"].append(f"campo_obligatorio_vacio:{campo}")
                        
                        # Comprobar si hay duplicados entre las muestras dentro del excel o si la muestra no estÃ¡ en la base de datos
                        nom_lab = datos["nom_lab"].lower()

                        if nom_lab not in cache["muestras_existentes"]:
                            errores[fila]["bloqueantes"].append("muestra_no_existe_bd")
                        if nom_lab in nom_lab_excel:
                            errores[fila]["bloqueantes"].append("muestra_duplicada_excel")
                        else:
                            nom_lab_excel.add(nom_lab)

                        # Comprobar si la localizaciÃ³n estÃ¡ ocupada o existe
                        key = (
                            datos["congelador"],
                            datos["estante"],
                            datos["rack"],
                            datos["bandeja"],
                            datos["caja"],
                            datos["fila"],
                            datos["columna"]
                        )

                        subposicion = cache["subposiciones"].get(key)

                        if not subposicion:
                            errores[fila]["bloqueantes"].append("localizacion_no_existe")
                        elif not subposicion.vacia:
                            errores[fila]["bloqueantes"].append("localizacion_ocupada")
                        else:
                            datos["subposicion_id"] = subposicion.id
                            datos["subposicion_antigua"] = cache['posiciones_actuales'].get(datos['nom_lab'].lower())
                        
                        # Registrar filas validas
                        if not errores[fila]["bloqueantes"]:
                            datos['fila'] = fila
                            filas_validas.append(datos)

                        if _should_update(idx, total_filas):
                            yield _json_progress(idx + 1, total_filas, 'processing')

                    # Guardar en la sesiÃ³n las filas validas y los errores detectados
                    request.session['filas_validas'] = filas_validas
                    request.session['errores'] = errores

                    # Mensajes de informaciÃ³n de la subida
                    upload_msgs = get_upload_messages('cambio_posicion')
                    mensajes = [{'level': 'info', 'text': f"{upload_msgs['titulo_inicial']} {numero_registros} registros."}]
                    numero_errores_bloqueantes = 0
                    for fila in errores:
                        if errores[fila]['bloqueantes']:
                            numero_errores_bloqueantes += 1
                    
                    errores_encontrados = numero_errores_bloqueantes > 0 or extra_columns
                    
                    if numero_errores_bloqueantes == 0 and not extra_columns:
                        mensajes.append({'level': 'success', 'text': upload_msgs['sin_errores']})
                    else:
                        if numero_errores_bloqueantes > 0:
                            msg_bloqueantes = upload_msgs['con_bloqueantes'].format(count=numero_errores_bloqueantes)
                            mensajes.append({'level': 'error', 'text': msg_bloqueantes})
                        if extra_columns:
                            msg_extras = upload_msgs['columnas_extras'].format(count=len(columnas_adicionales), detalles=columnas_adicionales_str)
                            mensajes.append({'level': 'warning', 'text': msg_extras})

                    # Guardar datos de confirmaciÃ³n en sesiÃ³n para mostrar tras redirecciÃ³n
                    request.session['confirmacion_pendiente'] = {
                        'template': 'confirmacion_upload_cambio_posicion.html',
                        'context': {'errores_encontrados': errores_encontrados},
                        'mensajes': mensajes
                    }
                    request.session.save()
                    yield _json_progress(total_filas, total_filas, 'done', 'ValidaciÃ³n completada', request.path + '?mostrar_confirmacion=1')

                return StreamingHttpResponse(gen_validar_cambio(), content_type='application/x-ndjson')
            
        # Si se solicita un excel de errores, este se rellena en base a los errores detectados durante la validaciÃ³n 
        elif 'excel_errores' in request.POST:
                    # Leer los errores y el excel de la sesiÃ³n
                    errores = request.session.get('errores',[])
                    columnas_adicionales_str = request.session.get('columnas_adicionales', '')
                    columnas_adicionales = set()
                    if columnas_adicionales_str:
                        columnas_adicionales = set(col.strip() for col in columnas_adicionales_str.split(','))
                    
                    excel_bytes = base64.b64decode(request.session.get('excel_file_base64'))
                    excel_file = io.BytesIO(excel_bytes)
                    wb = openpyxl.load_workbook(excel_file)
                    ws = wb.active
                    # Definir los estilos para pintar el excel usando configuraciÃ³n centralizada
                    colors = get_excel_colors()
                    FILL_ERROR_CELL = PatternFill("solid", fgColor=colors['error_cell'])
                    FILL_EXTRA_COL  = PatternFill("solid", fgColor=colors['extra_column'])
                    # Diccionario de mensajes
                    MENSAJES_ERROR = {
                        "campo_obligatorio_vacio": "Campo obligatorio vacÃ­o",
                        "muestra_no_existe_bd": "La muestra no existe en la base de datos",
                        "muestra_duplicada_excel": "Muestra duplicada dentro del Excel",
                        "localizacion_ocupada": "La subposiciÃ³n ya estÃ¡ ocupada",
                        "localizacion_no_existe": "La localizaciÃ³n no existe",
                        "fecha_invalida": "Fecha invÃ¡lida (Formato correcto: DD-MM-AAAA)",
                        "caracter_invalido_semicolon": "El carÃ¡cter ';' no estÃ¡ permitido en este campo"
                    }
                    # Diccionario de columnas del excel
                    columnas_excel = {}
                    rename_columns = {
                        'Nombre Laboratorio': 'nom_lab',
                        'Congelador': 'congelador', 
                        'Estante': 'estante',
                        'PosiciÃ³n del rack en el estante': 'posicion_rack_estante',
                        'Rack': 'rack',
                        'PosiciÃ³n de la bandeja en el rack': 'posicion_bandeja_rack',
                        'Bandeja': 'bandeja',
                        'PosiciÃ³n de la caja en la bandeja': 'posicion_caja_rack',
                        'Caja': 'caja',
                        'Fila': 'fila',
                        'Columna': 'columna',
                    }
                    for cell in ws[1]:
                        if cell.value in rename_columns:
                            columnas_excel[rename_columns[cell.value]] = cell.column
                    # AÃ±adir la columna de errores
                    col_errores = ws.max_column + 1
                    ws.cell(row=1, column=col_errores, value="Errores")
                    # Mapear errores sin campo a sus columnas especÃ­ficas (cambio posiciÃ³n)
                    error_campo_map_cambio = {
                        "muestra_no_existe_bd": "nom_lab",
                        "muestra_duplicada_excel": "nom_lab",
                        "localizacion_ocupada": "columna",
                        "localizacion_no_existe": "congelador"
                    }
                    # Recorrer filas con errores 
                    for fila, info in errores.items():
                        has_error = bool(info["bloqueantes"])
                        if not has_error:
                            continue

                        # Colorear celdas especÃ­ficas y construir mensajes
                        mensajes = []
                        for err in info["bloqueantes"]:
                            if ":" in err:
                                tipo, campo = err.split(":")
                                msg = f"(Error) {MENSAJES_ERROR[tipo]}"
                                if msg not in mensajes:
                                    mensajes.append(msg)
                                col = columnas_excel[campo]
                                celda = ws.cell(row=int(fila), column=col)
                                celda.fill = FILL_ERROR_CELL
                            else:
                                campo = error_campo_map_cambio.get(err)
                                msg = f"(Error) {MENSAJES_ERROR[err]}"
                                if msg not in mensajes:
                                    mensajes.append(msg)
                                if campo and campo in columnas_excel:
                                    col_err = columnas_excel[campo]
                                    ws.cell(row=int(fila), column=col_err).fill = FILL_ERROR_CELL
                        ws.cell(row=int(fila), column=col_errores, value="\n".join(mensajes))

                    # Pintar columnas extras con color de columna extra
                    if columnas_adicionales:
                        # ComparaciÃ³n case-insensitive porque la validaciÃ³n normaliza a minÃºsculas
                        extras_lower = {c.lower() for c in columnas_adicionales}
                        for cell in ws[1]:
                            if cell.value and str(cell.value).lower() in extras_lower:
                                col_num = cell.column
                                # Pintar el encabezado
                                ws.cell(row=1, column=col_num).fill = FILL_EXTRA_COL
                                # Pintar todas las celdas de datos en la columna
                                for row in range(2, ws.max_row + 1):
                                    ws.cell(row=row, column=col_num).fill = FILL_EXTRA_COL

                    # Rertornar el excel de errores    
                    output = io.BytesIO()    
                    wb.save(output)
                    wb.close()
                    response = HttpResponse(output.getvalue(),content_type='application/ms-excel')
                    response['Content-Disposition'] = 'attachment; filename="listado_errores.xlsx"'
                    return response        
    else:
        form = UploadExcel()
    return render(request, 'upload_excel_cambio_posicion.html', {'form': form}) 

@login_required
@permission_required('muestras.can_change_muestras_web')
def editar_muestra(request, nom_lab):
    # Vista para editar una muestra existente, requiere permiso para cambiar muestras
    muestra = Muestra.objects.get(nom_lab=nom_lab)
    if request.method == 'POST':
        # Guardar el estudio anterior ANTES de is_valid(), ya que is_valid() modifica la instancia
        estudio_anterior = muestra.estudio
        form = MuestraForm(request.POST, instance=muestra)
        if form.is_valid():
            # Verificar si nom_lab ha sido cambiado
            nom_lab_anterior = muestra.nom_lab
            nom_lab_nuevo = form.cleaned_data.get('nom_lab')
            
            with connection.cursor() as cursor:
                # Desactivar las restricciones de clave forÃ¡nea
                cursor.execute("SET FOREIGN_KEY_CHECKS=0")
            
            try:
                with transaction.atomic():
                    # Si nom_lab cambiÃ³ y hay localizaciones, actualizar todas las referencias
                    if nom_lab_nuevo != nom_lab_anterior and muestra.localizacion.exists():
                        with connection.cursor() as cursor:
                            # Actualizar todas las localizaciones con el nuevo nom_lab
                            cursor.execute(
                                "UPDATE muestras_localizacion SET muestra_id = %s WHERE muestra_id = %s",
                                [nom_lab_nuevo, nom_lab_anterior]
                            )
                    
                    # Preservar fechas si el usuario las dejÃ³ vacÃ­as pero habÃ­a valores anteriores
                    muestra_guardada = form.save(commit=False)
                    if not form.cleaned_data.get('fecha_extraccion') and muestra.fecha_extraccion:
                        muestra_guardada.fecha_extraccion = muestra.fecha_extraccion
                    if not form.cleaned_data.get('fecha_llegada') and muestra.fecha_llegada:
                        muestra_guardada.fecha_llegada = muestra.fecha_llegada
                    muestra_guardada.save()
                    
                    # Registrar en historial de estudios si el estudio ha cambiado
                    estudio_nuevo = muestra_guardada.estudio
                    if estudio_nuevo != estudio_anterior:
                        historial_estudios.objects.create(
                            muestra=muestra_guardada,
                            estudio=estudio_nuevo,
                            fecha_asignacion=timezone.now(),
                            usuario_asignacion=request.user
                        )
                    
                    # Manejar cambio de localizaciÃ³n si se proporcionÃ³
                    subposicion_id = request.POST.get("subposicion_id")
                    if subposicion_id:
                        try:
                            subposicion = Subposicion.objects.select_for_update().get(id=subposicion_id)
                            
                            if subposicion.vacia:
                                # Vaciar la subposiciÃ³n antigua si existe
                                if Subposicion.objects.filter(muestra=muestra).exists():
                                    subposicion_antigua = Subposicion.objects.select_for_update().get(muestra=muestra)
                                    subposicion_antigua.muestra = None
                                    subposicion_antigua.vacia = True
                                    subposicion_antigua.save()

                                # Asignar la nueva subposiciÃ³n
                                subposicion.muestra = muestra
                                subposicion.vacia = False
                                subposicion.save()

                                # Crear localizaciÃ³n
                                localizacion = _localizacion_desde_subposicion(muestra, subposicion)

                                historial_localizaciones.objects.create(
                                    muestra=muestra,
                                    localizacion=localizacion,
                                    fecha_asignacion=timezone.now(),
                                    usuario_asignacion=request.user
                                )
                                messages.success(request, 'Muestra editada correctamente')
                            else:
                                messages.error(request, 'La subposiciÃ³n estÃ¡ ocupada por otra muestra.')
                        except Subposicion.DoesNotExist:
                            messages.error(request, 'La subposiciÃ³n seleccionada no existe.')
                    else:
                        messages.success(request, 'Muestra editada correctamente')
            finally:
                # Reactivar las restricciones de clave forÃ¡nea
                with connection.cursor() as cursor:
                    cursor.execute("SET FOREIGN_KEY_CHECKS=1")
            
            return redirect('muestras_todas')
    else:
        form = MuestraForm(instance=muestra)
    
    # Obtener datos para los menÃºs desplegables
    congeladores = Congelador.objects.values('id', 'congelador').order_by('congelador')
    
    # Obtener localizaciÃ³n actual si existe. No acceder directamente a
    # `muestra.subposicion` porque para relaciones OneToOneDescriptor el acceso
    # puede lanzar RelatedObjectDoesNotExist cuando no existe la relaciÃ³n.
    localizacion_actual = None
    subposicion_actual = Subposicion.objects.filter(muestra=muestra).select_related(
        'caja__bandeja__rack__estante__congelador'
    ).first()
    if subposicion_actual:
        localizacion_actual = {
            'congelador_id': subposicion_actual.caja.rack.estante.congelador.id,
            'congelador_nombre': subposicion_actual.caja.rack.estante.congelador.congelador,
            'estante_id': subposicion_actual.caja.rack.estante.id,
            'estante_numero': subposicion_actual.caja.rack.estante.numero,
            'rack_id': subposicion_actual.caja.rack.id,
            'rack_numero': subposicion_actual.caja.rack.numero,
            'bandeja_id': subposicion_actual.caja.bandeja.id if subposicion_actual.caja.bandeja else None,
            'bandeja_numero': subposicion_actual.caja.bandeja.numero if subposicion_actual.caja.bandeja else None,
            'caja_id': subposicion_actual.caja.id,
            'caja_numero': subposicion_actual.caja.numero,
            'subposicion_id': subposicion_actual.id,
            'subposicion_numero': subposicion_actual.numero,
        }
    
    context = {
        'form': form,
        'muestra': muestra,
        'congeladores': congeladores,
        'localizacion_actual': localizacion_actual,
        'localizacion_actual_json': json.dumps(localizacion_actual),
    }
    
    return render(request, 'editar_muestra.html', context)


@login_required
@permission_required('muestras.can_change_muestras_web')
def destruir_muestras(request):
    # Vista para destruir mÃºltiples muestras con formulario detallado
    if 'muestras_destruir' not in request.session:
        messages.error(request, 'No hay muestras seleccionadas para destruir.')
        return redirect('muestras_todas')
    
    muestras_ids = request.session['muestras_destruir']
    qs = Muestra.objects.filter(id__in=muestras_ids)
    # Solo se puede destruir si la muestra estÃ¡ Disponible (DISP) o Parcialmente enviada (PENV)
    muestras_a_destruir = list(qs.filter(estado_actual__in=['DISP', 'PENV']))
    
    if not muestras_a_destruir:
        messages.error(request, 'Ninguna de las muestras seleccionadas puede ser destruida (deben estar Disponible o Parcialmente enviada).')
        del request.session['muestras_destruir']
        return redirect('muestras_todas')
    
    if request.method == 'POST':
        form = DestruirMuestrasForm(request.POST)
        if form.is_valid():
            # Procesar la destrucciÃ³n
            motivo = form.cleaned_data['motivo_destruccion']
            metodo = form.cleaned_data['metodo_destruccion']
            lugar = form.cleaned_data['lugar_destruccion']
            responsable = form.cleaned_data['responsable_autoriza']
            tecnico = form.cleaned_data['tecnico_realiza']
            fecha_destruccion = form.cleaned_data['fecha_destruccion']
            observaciones = form.cleaned_data['observaciones_destruccion']
            
            numero_muestras_destruidas = 0
            for sample in muestras_a_destruir:
                sample.estado_actual = 'DEST'
                sample.volumen_actual = 0
                sample.concentracion_actual = 0
                sample.save()
                # Liberar la subposiciÃ³n asociada
                if Subposicion.objects.filter(muestra=sample).exists():
                    subposicion = Subposicion.objects.get(muestra=sample)
                    subposicion.vacia = True
                    subposicion.muestra = None
                    subposicion.save()
                if Localizacion.objects.filter(muestra=sample).exists():
                    Localizacion.objects.filter(muestra=sample).update(muestra=None)
                # Registrar en trazabilidad
                registro_destruido.objects.create(
                    muestra=sample,
                    fecha=fecha_destruccion,
                    usuario=request.user,
                    motivo=motivo,
                    metodo=metodo,
                    lugar=lugar,
                    responsable=responsable,
                    tecnico=tecnico,
                    observaciones=observaciones
                )
                numero_muestras_destruidas += 1
            
            messages.success(request, f'{numero_muestras_destruidas} muestras destruidas correctamente.')
            del request.session['muestras_destruir']
            return redirect('muestras_todas')
    else:
        form = DestruirMuestrasForm()
    
    context = {
        'form': form,
        'muestras_a_destruir': muestras_a_destruir,
    }
    return render(request, 'destruir_muestras.html', context)


@login_required
@permission_required('muestras.can_view_muestras_web')
def historial_localizaciones_muestra(request,muestra_id):
    # Vista para ver el historial de localizaciones de una muestra especÃ­fica
    muestra = Muestra.objects.get(id=muestra_id)
    historiales = historial_localizaciones.objects.filter(muestra=muestra).order_by('-fecha_asignacion')
    if muestra.estado_actual=='DEST':
        estado_destruccion = registro_destruido.objects.filter(muestra=muestra).first()
    else:
        estado_destruccion = None
    template = loader.get_template('historial_localizaciones.html')
    return HttpResponse(template.render({'historiales':historiales, 'muestra':muestra, 'estado_destruccion':estado_destruccion},request))


@login_required
@permission_required('muestras.can_view_muestras_web')
def historial_destruccion(request, muestra_id):
    # Vista para ver el historial de destrucciÃ³n de una muestra especÃ­fica
    muestra = Muestra.objects.get(id=muestra_id)
    registros = registro_destruido.objects.filter(muestra=muestra).order_by('-fecha', '-id')
    template = loader.get_template('historial_destruccion.html')
    return HttpResponse(template.render({'muestra': muestra, 'registros': registros}, request))

