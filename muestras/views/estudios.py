import base64
import io
import json
import zipfile
from datetime import date

from django.conf import settings
import os

import openpyxl
import pandas as pd
from django.contrib import messages
from django.contrib.auth.decorators import login_required, permission_required
from django.core.paginator import EmptyPage, PageNotAnInteger, Paginator
from django.db import transaction
from django.db.models import Count, Q
from django.http import FileResponse, HttpResponse, StreamingHttpResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.template import loader
from django.utils import timezone
from django.utils.safestring import mark_safe
from django.utils.text import slugify
from openpyxl.styles import PatternFill

from ..forms import EstudioForm, DocumentoForm, UploadExcel
from ..models import Documento, Estudio, Muestra, historial_estudios
from ..parameters_config import get_excel_colors, get_upload_messages
from ..services.estudio_service import renombrar_estudio_y_migrar_muestras
from .common import _json_progress, _should_update
from django.contrib.auth.models import User


def _estudios_visibles_para_usuario(user):
    """Devuelve el alcance de estudios permitido para el usuario actual."""
    queryset = Estudio.objects.all()
    if user.groups.filter(name='Investigadores').exists():
        queryset = queryset.filter(investigadores_asociados=user)
    return queryset


def _get_estudio_accesible(request, id_estudio):
    """Obtiene un estudio solo si entra dentro del alcance visible del usuario."""
    return get_object_or_404(_estudios_visibles_para_usuario(request.user), id=id_estudio)


def _clear_estudio_session_selection(request):
    """Limpia la selección temporal de muestras usada por las acciones de estudio."""
    request.session.pop('muestras_estudio', None)

# Vistas relacionadas con el modelo estudio
@login_required
@permission_required('muestras.can_view_estudios_web')
def estudios_todos(request):
    # Vista para ver todos los estudios, los investigadores solo ven los suyos asociados
    scope_queryset = _estudios_visibles_para_usuario(request.user)
    queryset = scope_queryset.annotate(num_muestras=Count('muestra'))

    # Definición de campos filtrables
    field_names = [
        'referencia_estudio',
        'nombre_estudio',
        'descripcion_estudio',
        'investigador_principal',
    ]
    field_names_readable = [
        'Referencia del estudio',
        'Nombre del estudio',
        'Descripción / palabras clave',
        'Investigador principal',
    ]
    field_names_readable_dict = {k: v for k, v in zip(field_names, field_names_readable)}

    # Filtros por campo (multi-valor con ';' + icontains)
    for field in field_names:
        val = request.GET.get(field)
        if val:
            valores = [v.strip() for v in val.split(';') if v.strip()]
            if valores:
                q_filters = Q()
                for valor in valores:
                    q_filters |= Q(**{f"{field}__icontains": valor})
                queryset = queryset.filter(q_filters)

    # Filtros por rango de fechas (inicio/fin)
    fecha_inicio_desde = request.GET.get('fecha_inicio_desde')
    if fecha_inicio_desde:
        queryset = queryset.filter(fecha_inicio_estudio__gte=fecha_inicio_desde)

    fecha_inicio_hasta = request.GET.get('fecha_inicio_hasta')
    if fecha_inicio_hasta:
        queryset = queryset.filter(fecha_inicio_estudio__lte=fecha_inicio_hasta)

    fecha_fin_desde = request.GET.get('fecha_fin_desde')
    if fecha_fin_desde:
        queryset = queryset.filter(fecha_fin_estudio__gte=fecha_fin_desde)

    fecha_fin_hasta = request.GET.get('fecha_fin_hasta')
    if fecha_fin_hasta:
        queryset = queryset.filter(fecha_fin_estudio__lte=fecha_fin_hasta)

    # Búsqueda general (server-side)
    busqueda_general = request.GET.get('busqueda', '').strip()
    if busqueda_general:
        q_busqueda = Q()
        q_busqueda |= Q(referencia_estudio__icontains=busqueda_general)
        q_busqueda |= Q(nombre_estudio__icontains=busqueda_general)
        q_busqueda |= Q(descripcion_estudio__icontains=busqueda_general)
        q_busqueda |= Q(investigador_principal__icontains=busqueda_general)
        queryset = queryset.filter(q_busqueda)
    

    # Paginación
    queryset = queryset.order_by('id')
    contador_total = queryset.count()
    items_por_pagina = request.GET.get('items_por_pagina', 25)
    if str(items_por_pagina) == 'todas':
        items_por_pagina = 'todas'
        paginator = Paginator(queryset, max(contador_total, 1))
    else:
        try:
            items_por_pagina = int(items_por_pagina)
            if items_por_pagina not in [10, 25, 50, 100]:
                items_por_pagina = 25
        except Exception:
            items_por_pagina = 25
        paginator = Paginator(queryset, items_por_pagina)

    numero_pagina = request.GET.get('page', 1)
    try:
        estudios_page = paginator.page(numero_pagina)
    except PageNotAnInteger:
        estudios_page = paginator.page(1)
    except EmptyPage:
        estudios_page = paginator.page(paginator.num_pages)

    # Listas para los multiselect de filtros
    opciones_referencias = (
        scope_queryset
        .exclude(referencia_estudio__isnull=True)
        .exclude(referencia_estudio__exact='')
        .values_list('referencia_estudio', flat=True)
        .distinct()
        .order_by('referencia_estudio')
    )

    opciones_investigadores = (
        scope_queryset
        .exclude(investigador_principal__isnull=True)
        .exclude(investigador_principal__exact='')
        .values_list('investigador_principal', flat=True)
        .distinct()
        .order_by('investigador_principal')
    )

    template = loader.get_template('estudios_todos.html')
    context = {
        'estudios': estudios_page.object_list,
        'paginator': paginator,
        'muestras_page': estudios_page,
        'contador_muestras': contador_total,
        'items_por_pagina': items_por_pagina,
        'busqueda': busqueda_general,
        'estudios_ids_json': json.dumps(list(queryset.values_list('id', flat=True))),
        'opciones_referencias': opciones_referencias,
        'opciones_investigadores': opciones_investigadores,
        'field_names': field_names,
        'field_names_readable_dict': field_names_readable_dict,
        'request': request,
    }
    return HttpResponse(template.render(context, request))

@login_required
@permission_required('muestras.can_change_estudios_web')
def acciones_estudios(request):
    """
    Acciones masivas sobre estudios: exportar, eliminar, editar, documentación,
    importar (redirección) y nuevo estudio (redirección).
    """
    if request.method == "POST":
        estudios_seleccionados = request.POST.getlist('estudio_id')

        # Exportar seleccionados a Excel
        if 'exportar_seleccionados' in request.POST:
            if not estudios_seleccionados:
                messages.error(request, "No has seleccionado ningún estudio para exportar.")
                return redirect('estudios_todos')

            queryset = _estudios_visibles_para_usuario(request.user).filter(id__in=estudios_seleccionados)
            response = HttpResponse(
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            response['Content-Disposition'] = 'attachment; filename="estudios_seleccionados.xlsx"'

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Estudios"

            headers = [
                "ID",
                "Referencia",
                "Nombre",
                "Descripción",
                "Fecha inicio",
                "Fecha fin",
                "Investigador principal",
            ]
            for col, h in enumerate(headers, 1):
                ws.cell(row=1, column=col).value = h

            row_num = 2
            for e in queryset:
                ws.cell(row_num, 1).value = e.id
                ws.cell(row_num, 2).value = e.referencia_estudio
                ws.cell(row_num, 3).value = e.nombre_estudio
                ws.cell(row_num, 4).value = e.descripcion_estudio
                ws.cell(row_num, 5).value = str(e.fecha_inicio_estudio or "")
                ws.cell(row_num, 6).value = str(e.fecha_fin_estudio or "")
                ws.cell(row_num, 7).value = e.investigador_principal
                row_num += 1

            wb.save(response)
            return response

        # Eliminar (solo estudios sin muestras asociadas)
        elif 'eliminar' in request.POST:
            if not estudios_seleccionados:
                messages.error(request, "No has seleccionado estudios para eliminar.")
                return redirect('estudios_todos')

            eliminados = 0
            bloqueados = 0
            for est_id in estudios_seleccionados:
                estudio = _estudios_visibles_para_usuario(request.user).filter(id=est_id).first()
                if not estudio:
                    continue
                if Muestra.objects.filter(estudio=estudio).exists():
                    bloqueados += 1
                    continue
                estudio.delete()
                eliminados += 1

            if eliminados:
                messages.success(request, f"{eliminados} estudio(s) eliminado(s) correctamente.")
            if bloqueados:
                messages.error(request, f"{bloqueados} estudio(s) no se pudieron eliminar por tener muestras asociadas.")
            return redirect('estudios_todos')

        # Editar (requiere exactamente 1 seleccionado)
        elif 'editar' in request.POST:
            if len(estudios_seleccionados) != 1:
                messages.error(request, "Selecciona exactamente un estudio para editarlo.")
                return redirect('estudios_todos')
            return redirect('editar_estudio', id_estudio=estudios_seleccionados[0])

        # Documentación (requiere exactamente 1 seleccionado)
        elif 'documentacion' in request.POST:
            if len(estudios_seleccionados) != 1:
                messages.error(request, "Selecciona exactamente un estudio para ver su documentación.")
                return redirect('estudios_todos')
            return redirect('repositorio_estudio', id_estudio=estudios_seleccionados[0])

        # Importar (redirige)
        elif 'importar' in request.POST:
            return redirect('excel_estudios')

        # Nuevo estudio (redirige)
        elif 'nuevo_estudio' in request.POST:
            return redirect('nuevo_estudio')

        # Acción no reconocida
        messages.error(request, "Acción no reconocida.")
        return redirect('estudios_todos')

    # Si no es POST, vuelve al listado
    return redirect('estudios_todos')
@login_required
@permission_required('muestras.can_add_estudios_web')
def excel_estudios(request):
    # Mostrar confirmación después de validación con progreso
    if request.GET.get('mostrar_confirmacion') and 'confirmacion_pendiente' in request.session:
        datos_conf = request.session.pop('confirmacion_pendiente')
        for msg in datos_conf.get('mensajes', []):
            getattr(messages, msg['level'])(request, msg['text'])
        return render(request, datos_conf['template'], datos_conf.get('context', {}))
    # Vista para subir estudios desde un archivo Excel, requiere permiso para añadir estudios
    if request.method=="POST":
        form = UploadExcel(request.POST, request.FILES)
        # Si el usuario confirma, se guardan los estudios en la base de datos
        if 'confirmar' in request.POST:
            filas_validas = request.session.get('filas_validas',[])
            total = len(filas_validas)

            def gen_confirmar_estudios():
                yield _json_progress(0, total, 'start')
                try:
                    with transaction.atomic():
                        for i, datos in enumerate(filas_validas):
                            fecha_inicio_estudio = date.fromisoformat(datos["fecha_inicio_estudio"]) if datos["fecha_inicio_estudio"] else None
                            fecha_fin_estudio = date.fromisoformat(datos["fecha_fin_estudio"]) if datos["fecha_fin_estudio"] else None
                            Estudio.objects.create(
                                referencia_estudio=datos['referencia_estudio'],
                                nombre_estudio=datos['nombre_estudio'],
                                descripcion_estudio=datos['descripcion_estudio'],
                                fecha_inicio_estudio=fecha_inicio_estudio,
                                fecha_fin_estudio=fecha_fin_estudio,
                                investigador_principal=datos['investigador_principal']
                            )
                            if _should_update(i, total):
                                yield _json_progress(i + 1, total, 'processing')
                    yield _json_progress(total, total, 'done', 'Estudios importados correctamente', '/estudios/')
                except Exception as e:
                    yield _json_progress(0, total, 'error', str(e))

            return StreamingHttpResponse(gen_confirmar_estudios(), content_type='application/x-ndjson')
        # Si el usuario cancela, no se hace nada
        elif 'cancelar' in request.POST:
            messages.error(request,'Los estudios no se han añadido')
            return redirect('estudios_todos')
        # Si se sube un archivo excel, se procesa y valida
        elif 'excel_file' in request.FILES:
            if form.is_valid():
                # Leer excel y preparar columnas
                excel_file = request.FILES['excel_file']
                excel_bytes = excel_file.read()
                request.session['excel_file_name'] = excel_file.name
                request.session['excel_file_base64']= base64.b64encode(excel_bytes).decode()
                excel_stream = io.BytesIO(excel_bytes)
                try:
                    df = pd.read_excel(excel_stream)
                except Exception as e:
                    return render(request, 'upload_excel_estudios.html', {'form': form, 'error': f'❌ Error al leer el archivo: {str(e)}'})
                if df.empty:
                    return render(request, 'upload_excel_estudios.html', {'form': form, 'error': '❌ Error de formato: El archivo Excel está vacío o no contiene filas de datos.'})
                rename_columns = {
                    'Referencia del estudio': 'referencia_estudio', 
                    'Nombre del estudio': 'nombre_estudio',
                    'Descripción': 'descripcion_estudio',
                    'Fecha de inicio': 'fecha_inicio_estudio',
                    'Fecha de fin': 'fecha_fin_estudio',
                    'Investigador principal': 'investigador_principal',
                }
                # Validar columnas
                columnas_esperadas = set(rename_columns.keys())
                columnas_existentes = set(df.columns)
                columnas_faltantes = columnas_esperadas - columnas_existentes
                if columnas_faltantes:
                    columnas_str = ', '.join(sorted(columnas_faltantes))
                    return render(request, 'upload_excel_estudios.html', {'form': form, 'error': f'❌ Error de formato: El archivo Excel no contiene las siguientes columnas esperadas: {columnas_str}'})
                columnas_adicionales = columnas_existentes - columnas_esperadas
                if columnas_adicionales:
                    columnas_adicionales_str = ', '.join(sorted(columnas_adicionales))
                    request.session['columnas_adicionales'] = columnas_adicionales_str
                else:
                    request.session.pop('columnas_adicionales', None)
                df.rename(columns=rename_columns, inplace=True)
                 # Funciones para normalizar las columnas del excel
                def norm(value):
                    if value is None or pd.isna(value):
                        return None

                    if isinstance(value, str):
                        value = value.strip()
                        return value if value != "" else None

                    return value
                '''
                def norm_code(value):
                    if value is None or pd.isna(value):
                        return None

                    if isinstance(value, float) and value.is_integer():
                        return str(int(value))

                    return str(value).strip()
                def convertir_fecha(value):
                        if pd.isna(value) or value is None:
                            return None
                        fecha = pd.to_datetime(value)
                        return fecha.date()
                '''
                # Crear estructuras previas del excel
                total_filas = len(df)

                def gen_validar_estudios():
                    filas_validas = []
                    errores = {}
                    nombre_estudios_excel = set()
                    numero_registros = 0
                    cache_inner = {
                        'estudios_existentes_lower': set(n.lower() for n in Estudio.objects.values_list('nombre_estudio', flat=True).distinct() if n),
                        'referencias_existentes_lower': set(str(x).strip().lower() for x in Estudio.objects.values_list('referencia_estudio', flat=True).distinct() if x is not None)
                    }
                    referencias_excel = set()

                    yield _json_progress(0, total_filas, 'start')

                    for idx, row in df.iterrows():
                        # Recorrer el df para detectar errores y normalizar
                        numero_registros += 1
                        fila = idx + 2 
                        errores[fila]={"advertencias":[], "bloqueantes":[]}
                        datos = {
                            "nombre_estudio":norm(row['nombre_estudio']),
                            'referencia_estudio': norm(row['referencia_estudio']),
                            'descripcion_estudio': norm(row['descripcion_estudio']),
                            'fecha_inicio_estudio':norm(row['fecha_inicio_estudio']),
                            'fecha_fin_estudio': norm(row['fecha_fin_estudio']),
                            'investigador_principal': norm(row['investigador_principal'])
                        }
                        # Detectar campos vacios
                        optativos = ["referencia_estudio", "descripcion_estudio", "fecha_inicio_estudio", "fecha_fin_estudio", "investigador_principal"]
                        for campo in optativos:
                            if not datos.get(campo):
                                errores[fila]["advertencias"].append(f"campo_optativo_vacio:{campo}")
                        obligatorios = ["nombre_estudio"]
                        for campo in obligatorios:
                            if not datos.get(campo):
                                errores[fila]["bloqueantes"].append(f"campo_obligatorio_vacio:{campo}")
                        
                        # Validar que ningún campo contenga el carácter punto y coma (;)
                        campos_a_validar = ['referencia_estudio', 'nombre_estudio', 'descripcion_estudio', 'investigador_principal']
                        for campo in campos_a_validar:
                            valor = datos.get(campo)
                            if valor and isinstance(valor, str) and ';' in valor:
                                errores[fila]["bloqueantes"].append(f"caracter_invalido_semicolon:{campo}")
                        
                        # Detectar formato de fecha incorrecto (advertencia)
                        for campo in ['fecha_inicio_estudio', 'fecha_fin_estudio']:
                            if datos[campo] != None:
                                try:
                                    # Si es un Timestamp de pandas o datetime, usar directamente
                                    if isinstance(datos[campo], (pd.Timestamp, type(pd.NaT))):
                                        if pd.isna(datos[campo]):
                                            datos[campo] = None
                                        else:
                                            # Convertir Timestamp/datetime a ISO string
                                            datos[campo] = datos[campo].date().isoformat()
                                    elif isinstance(datos[campo], date):
                                        # Si es un objeto date, convertir directamente a ISO
                                        datos[campo] = datos[campo].isoformat()
                                    else:
                                        # Si es string, parsear con formato DD-MM-AAAA
                                        fecha_str = str(datos[campo]).strip()
                                        partes = fecha_str.split('-')
                                        if len(partes) == 3 and all(p.isdigit() for p in partes):
                                            fecha = pd.to_datetime(fecha_str, format='%d-%m-%Y')
                                            datos[campo] = fecha.date().isoformat()
                                        else:
                                            errores[fila]["bloqueantes"].append(f"fecha_invalida:{campo}")
                                            datos[campo] = None
                                except Exception:
                                    errores[fila]["bloqueantes"].append(f"fecha_invalida:{campo}")
                                    datos[campo] = None

                        # Validar que fecha_fin >= fecha_inicio si ambas están informadas
                        fecha_inicio = datos.get('fecha_inicio_estudio')
                        fecha_fin = datos.get('fecha_fin_estudio')
                        if fecha_inicio and fecha_fin:
                            # Ambas fechas están informadas y son válidas
                            if fecha_fin < fecha_inicio:
                                errores[fila]["bloqueantes"].append("fecha_fin_menor_que_inicio")

                        # Detectar si el estudio ya existe
                        nombre_estudio = datos['nombre_estudio']
                        if nombre_estudio:
                            # Normalizar a string en caso de que sea numérico
                            nombre_estudio_str = str(nombre_estudio).strip()
                            nombre_estudio_lower = nombre_estudio_str.lower()
                            if nombre_estudio_lower in cache_inner['estudios_existentes_lower']:
                                errores[fila]["bloqueantes"].append(f"estudio_existente")
                            if nombre_estudio_lower in nombre_estudios_excel:
                                errores[fila]["bloqueantes"].append("estudio_duplicado_excel")
                            else:
                                nombre_estudios_excel.add(nombre_estudio_lower)
                        else:
                            nombre_estudio_lower = ''
                        # Validar referencia_estudio: si existe, no puede coincidir con otras en DB ni duplicarse en el Excel
                        referencia = datos.get('referencia_estudio')
                        if referencia:
                            # Normalizar a string antes de usar lower() para admitir valores numéricos en Excel
                            ref_str = str(referencia).strip()
                            if ref_str:
                                ref_lower = ref_str.lower()
                                if ref_lower in cache_inner['referencias_existentes_lower']:
                                    errores[fila]["bloqueantes"].append("referencia_existente")
                                if ref_lower in referencias_excel:
                                    errores[fila]["bloqueantes"].append("referencia_duplicada_excel")
                                else:
                                    referencias_excel.add(ref_lower)
                       
                        # Registrar filas validas
                        if not errores[fila]["bloqueantes"]:
                            filas_validas.append(datos)

                        if _should_update(idx, total_filas):
                            yield _json_progress(idx + 1, total_filas, 'processing')

                    request.session['filas_validas'] = filas_validas
                    request.session['errores'] = errores

                    # Obtener configuración de mensajes para estudios
                    msg_config = get_upload_messages('estudios')

                    # Contar errores
                    numero_errores_bloqueantes = 0
                    numero_errores_advertencia = 0
                    for fila in errores:
                        if errores[fila]['bloqueantes']:
                            numero_errores_bloqueantes += 1
                        if errores[fila]["advertencias"]:
                            numero_errores_advertencia += 1

                    mensajes = []

                    # Mensaje inicial
                    mensajes.append({'level': 'info', 'text': f'{msg_config["titulo_inicial"]} {numero_registros} registros.'})

                    # Generar mensajes según el estado
                    if numero_errores_advertencia > 0:
                        msg = msg_config['con_advertencias'].format(count=numero_errores_advertencia)
                        mensajes.append({'level': 'warning', 'text': msg})
                    if numero_errores_bloqueantes > 0:
                        msg = msg_config['con_bloqueantes'].format(count=numero_errores_bloqueantes)
                        mensajes.append({'level': 'error', 'text': msg})
                    if numero_errores_bloqueantes == 0 and numero_errores_advertencia == 0:
                        mensajes.append({'level': 'success', 'text': msg_config['sin_errores']})

                    # Manejar columnas extras
                    columnas_extras_str = request.session.get('columnas_adicionales', '')
                    tiene_columnas_extras = bool(columnas_extras_str)
                    numero_columnas_extras = len(columnas_extras_str.split(', ')) if columnas_extras_str else 0
                    if tiene_columnas_extras:
                        msg = msg_config['columnas_extras'].format(count=numero_columnas_extras, detalles=columnas_extras_str)
                        mensajes.append({'level': 'warning', 'text': msg})

                    request.session['confirmacion_pendiente'] = {
                        'template': 'confirmacion_upload_estudios.html',
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
                    yield _json_progress(total_filas, total_filas, 'done', 'Validación completada', request.path + '?mostrar_confirmacion=1')

                return StreamingHttpResponse(gen_validar_estudios(), content_type='application/x-ndjson')
            
        # Si se solicita un excel de errores, este se rellena en base a los errores detectados durante la validación
        elif 'excel_errores' in request.POST:
                    errores = request.session.get('errores',[])
                    excel_bytes = base64.b64decode(request.session.get('excel_file_base64'))
                    excel_file = io.BytesIO(excel_bytes)
                    wb = openpyxl.load_workbook(excel_file)
                    ws = wb.active
                    # Definir los estilos para pintar el excel usando configuración centralizada
                    colors = get_excel_colors()
                    FILL_ERROR_CELL = PatternFill("solid", fgColor=colors['error_cell'])
                    FILL_WARN_CELL  = PatternFill("solid", fgColor=colors['warning_cell'])
                    FILL_EXTRA_COL  = PatternFill("solid", fgColor=colors['extra_column'])
                    # Diccionario de mensajes
                    MENSAJES_ERROR = {
                        "campo_obligatorio_vacio": "Campo obligatorio vacío",
                        "fecha_invalida": "Fecha inválida (Formato correcto: DD-MM-AAAA)",
                        "fecha_fin_menor_que_inicio": "Fecha de fin anterior a fecha de inicio",
                        "estudio_existente": "El estudio ya existe en la base de datos",
                        "estudio_duplicado_excel": "Estudio duplicado dentro del Excel",
                        "referencia_existente": "Referencia ya existe en la base de datos",
                        "referencia_duplicada_excel": "Referencia duplicada dentro del Excel",
                        "campo_optativo_vacio": "Campo opcional vacío",
                        "caracter_invalido_semicolon": "El carácter ';' no está permitido en este campo"
                    }
                    # Diccionario de columnas del excel
                    columnas_excel = {}
                    rename_columns = {
                    'Referencia del estudio': 'referencia_estudio', 
                    'Nombre del estudio': 'nombre_estudio',
                    'Descripción': 'descripcion_estudio',
                    'Fecha de inicio': 'fecha_inicio_estudio',
                    'Fecha de fin': 'fecha_fin_estudio',
                    'Investigador principal': 'investigador_principal',
                    }
                    for cell in ws[1]:
                        columnas_excel[rename_columns.get(cell.value, cell.value)] = cell.column
                    # Añadir la columna de errores
                    col_errores = ws.max_column + 1
                    ws.cell(row=1, column=col_errores, value="Errores")
                    # Mapeo de errores sin campo a sus columnas específicas (estudios)
                    error_campo_map_study = {
                        "estudio_existente": "nombre_estudio",
                        "estudio_duplicado_excel": "nombre_estudio",
                        "referencia_existente": "referencia_estudio",
                        "referencia_duplicada_excel": "referencia_estudio"
                    }
                    # Recorrer filas con errores 
                    for fila, info in errores.items():
                        has_error = bool(info.get("bloqueantes", []))
                        has_warn = bool(info.get("advertencias", []))
                        if not has_error and not has_warn:
                            continue

                        # Colorear celdas específicas y construir mensajes
                        mensajes = []
                        for err in info.get("bloqueantes", []):
                            if ":" in err:
                                tipo, campo = err.split(":")
                                msg = f"(Error) {MENSAJES_ERROR[tipo]}"
                                if msg not in mensajes:
                                    mensajes.append(msg)
                                col = columnas_excel[campo]
                                celda = ws.cell(row=int(fila), column=col)
                                celda.fill = FILL_ERROR_CELL
                            else:
                                campo = error_campo_map_study.get(err)
                                msg = f"(Error) {MENSAJES_ERROR[err]}"
                                if msg not in mensajes:
                                    mensajes.append(msg)
                                if campo and campo in columnas_excel:
                                    col_err = columnas_excel[campo]
                                    ws.cell(row=int(fila), column=col_err).fill = FILL_ERROR_CELL
                        for warn in info.get("advertencias", []): 
                            if ":" in warn:
                                tipo, campo = warn.split(":", 1)
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
                    expected_renamed = set(rename_columns.values())
                    for col_name, col_num in columnas_excel.items():
                        if col_name not in expected_renamed:
                            # Pintar el encabezado
                            header_cell = ws.cell(row=1, column=col_num)
                            header_cell.fill = FILL_EXTRA_COL
                            # Pintar todas las celdas de datos en la columna
                            for row in range(2, ws.max_row + 1):
                                ws.cell(row=row, column=col_num).fill = FILL_EXTRA_COL

                    output = io.BytesIO()    
                    wb.save(output)
                    wb.close()
                    response = HttpResponse(output.getvalue(),content_type='application/ms-excel')
                    response['Content-Disposition'] = 'attachment; filename="listado_errores.xlsx"'
                    return response      
    else: 
        form= UploadExcel()
    return render(request, 'upload_excel_estudios.html', {'form': form}) 
@login_required
@permission_required('muestras.can_delete_estudios_web')
def eliminar_estudio(request, id_estudio):
    # Vista para eliminar un estudio existente
    estudio = _get_estudio_accesible(request, id_estudio)
    if Muestra.objects.filter(estudio=estudio).exists():
        messages.error(request, mark_safe(f'No se puede eliminar el estudio "{estudio.nombre_estudio}" porque tiene muestras asociadas. Desasocia las muestras primero.'))
        return redirect('estudios_todos')
    estudio.delete()
    messages.success(request,'Estudio eliminado correctamente')
    return redirect('estudios_todos')
@login_required
@permission_required('muestras.can_change_estudios_web')
def seleccionar_estudio(request):
    # Vista para seleccionar un estudio al que añadir muestras
    estudios = _estudios_visibles_para_usuario(request.user).order_by('nombre_estudio')
    template = loader.get_template('seleccionar_estudio.html')
    return HttpResponse(template.render({'estudios':estudios},request))
@login_required
@permission_required('muestras.can_change_estudios_web')
def añadir_muestras_estudio(request):
    # Vista para añadir muestras a un estudio seleccionado (progreso AJAX)
    if request.method == 'POST':
        # Obtener las muestras de la sesión
        muestras = request.session.get('muestras_estudio', [])
        muestras = list(Muestra.objects.filter(id__in=muestras))

        # Desasociar muestras de sus estudios si se selecciona esa opción
        if len(request.POST.getlist('desasociar')) == 1:
            total = len(muestras)
            def gen_desasociar():
                yield _json_progress(0, total, 'start')
                try:
                    for i, muestra in enumerate(muestras):
                        if muestra.estado_actual != 'DEST':
                            muestra.estudio = None
                            muestra.save()
                            historial = historial_estudios.objects.create(
                                muestra=muestra,
                                estudio=None,
                                fecha_asignacion=timezone.now(),
                                usuario_asignacion=request.user
                            )
                            historial.save()
                        if _should_update(i, total):
                            yield _json_progress(i + 1, total, 'processing')
                    if 'muestras_estudio' in request.session:
                        del request.session['muestras_estudio']
                    yield _json_progress(total, total, 'done', f'{total} muestras desasociadas correctamente')
                except Exception as e:
                    yield _json_progress(0, total, 'error', str(e))
            return StreamingHttpResponse(gen_desasociar(), content_type='application/x-ndjson')

        # Obtener los estudios seleccionados y asociar las muestras
        ids_estudios = request.POST.getlist('estudio_nombre')
        total = len(muestras) * len(ids_estudios)
        def gen_asociar():
            yield _json_progress(0, total, 'start')
            try:
                current = 0
                estudios_permitidos = {
                    estudio.nombre_estudio: estudio
                    for estudio in _estudios_visibles_para_usuario(request.user).filter(nombre_estudio__in=ids_estudios)
                }
                for study in ids_estudios:
                    studio = estudios_permitidos.get(study)
                    if not studio:
                        raise Estudio.DoesNotExist(f"Estudio no accesible: {study}")
                    for muestra in muestras:
                        if muestra.estado_actual != 'DEST':
                            muestra.estudio = studio
                            muestra.save()
                            if not historial_estudios.objects.filter(muestra=muestra, estudio=studio).exists():
                                historial = historial_estudios.objects.create(
                                    muestra=muestra,
                                    estudio=studio,
                                    fecha_asignacion=timezone.now(),
                                    usuario_asignacion=request.user
                                )
                                historial.save()
                        current += 1
                        if _should_update(current - 1, total):
                            yield _json_progress(current, total, 'processing')
                if 'muestras_estudio' in request.session:
                    del request.session['muestras_estudio']
                yield _json_progress(total, total, 'done', 'Muestras añadidas correctamente a los estudios')
            except Exception as e:
                yield _json_progress(0, total, 'error', str(e))
        return StreamingHttpResponse(gen_asociar(), content_type='application/x-ndjson')
    return redirect('muestras_todas')

# Implementaciones normalizadas del módulo de estudios.
@login_required
@permission_required('muestras.can_add_estudios_web')
def nuevo_estudio(request):
    """Crea un estudio nuevo usando la validación centralizada del formulario."""
    if request.method == 'POST':
        form = EstudioForm(request.POST)
        if form.is_valid():
            form.save()
            messages.success(request, 'Estudio añadido correctamente')
            return redirect('estudios_todos')
    else:
        form = EstudioForm()
    template = loader.get_template('nuevo_estudio.html')
    return HttpResponse(template.render({'form': form}, request))


@login_required
@permission_required('muestras.can_change_estudios_web')
def editar_estudio(request, id_estudio):
    """Edita un estudio y sincroniza sus referencias en muestras al renombrarlo."""
    estudio = _get_estudio_accesible(request, id_estudio)
    if request.method == 'POST':
        form = EstudioForm(request.POST, instance=estudio)
        if form.is_valid():
            renombrar_estudio_y_migrar_muestras(estudio, form.cleaned_data['nombre_estudio'])
            form.save()
            messages.info(request, 'El estudio se ha modificado correctamente')
            return redirect('estudios_todos')
    else:
        form = EstudioForm(instance=estudio)
    return render(request, 'editar_estudio.html', {'form': form, 'estudio': estudio})


@login_required
@permission_required('muestras.can_view_estudios_web')
def historial_estudios_muestra(request, muestra_id):
    """Muestra el historial de asignaciones de estudio de una muestra."""
    muestra = get_object_or_404(Muestra, id=muestra_id)
    historiales = historial_estudios.objects.filter(muestra=muestra).order_by('-fecha_asignacion')
    template = loader.get_template('historial_estudios.html')
    return HttpResponse(template.render({'historiales': historiales, 'muestra': muestra}, request))


@login_required
@permission_required('muestras.can_view_estudios_web')
def repositorio_estudio(request, id_estudio):
    """Muestra el repositorio documental del estudio con filtros y paginación."""
    estudio = _get_estudio_accesible(request, id_estudio)
    documentos = (
        Documento.objects
        .filter(estudio=estudio, eliminado=False)
        .select_related('usuario_subida', 'estudio')
        .order_by('-fecha_subida', '-id')
    )
    request.session['id'] = id_estudio
    usuarios = (
        User.objects
        .filter(documento__estudio=estudio, documento__eliminado=False)
        .distinct()
        .order_by('username')
    )

    usuario = request.GET.get('usuario')
    if usuario:
        usuarios_list = [u.strip() for u in usuario.split(';') if u.strip()]
        if usuarios_list:
            q_filters = Q()
            for user in usuarios_list:
                q_filters |= Q(usuario_subida__username=user)
            documentos = documentos.filter(q_filters)

    categoria = request.GET.get('categoria')
    if categoria:
        categorias_list = [c.strip() for c in categoria.split(';') if c.strip()]
        if categorias_list:
            q_filters = Q()
            for cat in categorias_list:
                q_filters |= Q(categoria__icontains=cat)
            documentos = documentos.filter(q_filters)

    busqueda_general = request.GET.get('busqueda', '').strip()
    if busqueda_general:
        q_busqueda = Q()
        q_busqueda |= Q(categoria__icontains=busqueda_general)
        q_busqueda |= Q(descripcion__icontains=busqueda_general)
        q_busqueda |= Q(usuario_subida__username__icontains=busqueda_general)
        q_busqueda |= Q(archivo__icontains=busqueda_general)
        documentos = documentos.filter(q_busqueda)

    contador_total = documentos.count()
    items_por_pagina = request.GET.get('items_por_pagina', 25)
    if str(items_por_pagina) == 'todas':
        items_por_pagina = 'todas'
        paginator = Paginator(documentos, max(contador_total, 1))
    else:
        try:
            items_por_pagina = int(items_por_pagina)
            if items_por_pagina not in [10, 25, 50, 100]:
                items_por_pagina = 25
        except Exception:
            items_por_pagina = 25
        paginator = Paginator(documentos, items_por_pagina)

    numero_pagina = request.GET.get('page', 1)
    try:
        documentos_page = paginator.page(numero_pagina)
    except PageNotAnInteger:
        documentos_page = paginator.page(1)
    except EmptyPage:
        documentos_page = paginator.page(paginator.num_pages)

    template = loader.get_template('repositorio_estudio.html')
    context = {
        'documentos': documentos_page.object_list,
        'paginator': paginator,
        'muestras_page': documentos_page,
        'contador_muestras': contador_total,
        'items_por_pagina': items_por_pagina,
        'busqueda': busqueda_general,
        'id': estudio.id,
        'estudio': estudio,
        'usuarios': usuarios,
        'request': request,
    }
    return HttpResponse(template.render(context, request))


@login_required
@permission_required('muestras.can_change_estudios_web')
def subir_documento(request, id_estudio):
    """Sube un documento al repositorio de un estudio concreto."""
    estudio = _get_estudio_accesible(request, id_estudio)
    if request.method == 'POST':
        form = DocumentoForm(request.POST, request.FILES)
        if form.is_valid():
            doc = form.save(commit=False)
            doc.usuario_subida = request.user
            doc.estudio = estudio
            doc.save()
            return redirect('repositorio_estudio', id_estudio=estudio.id)
        messages.error(request, 'Hubo un error al subir el documento.')
    else:
        form = DocumentoForm()
    template = loader.get_template('subir_documento.html')
    return HttpResponse(template.render({'form': form, 'estudio': estudio}, request))


@login_required
@permission_required('muestras.can_view_estudios_web')
def descargar_documento(request, id_estudio, documento_id):
    """Descarga un documento verificando que pertenece al estudio solicitado."""
    estudio = _get_estudio_accesible(request, id_estudio)
    doc = get_object_or_404(Documento, pk=documento_id, estudio=estudio, eliminado=False)
    if not doc.archivo or not os.path.exists(doc.archivo.path):
        messages.error(request, 'El archivo solicitado no se encuentra disponible.')
        return redirect('repositorio_estudio', id_estudio=estudio.id)
    return FileResponse(doc.archivo.open('rb'), as_attachment=True, filename=os.path.basename(doc.archivo.name))


@login_required
@permission_required('muestras.can_view_estudios_web')
def descargar_documentos_seleccionados(request):
    """Genera un ZIP con los documentos seleccionados del estudio activo."""
    ids_documento = request.POST.getlist('doc_id')
    id_estudio = request.POST.get('id_estudio') or request.session.get('id')
    if not ids_documento or not id_estudio:
        if id_estudio:
            return redirect('repositorio_estudio', id_estudio=id_estudio)
        return redirect('estudios_todos')

    estudio = _get_estudio_accesible(request, id_estudio)
    documentos = list(
        Documento.objects.filter(
            pk__in=ids_documento,
            eliminado=False,
            estudio=estudio,
        ).select_related('estudio')
    )
    if not documentos:
        return redirect('repositorio_estudio', id_estudio=estudio.id)

    zip_buffer = io.BytesIO()
    zip_filename = f"documentos_estudio_{slugify(estudio.nombre_estudio) or estudio.id}.zip"
    archivos_anadidos = 0

    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        used_names = set()
        for doc in documentos:
            if not doc.archivo or not os.path.exists(doc.archivo.path):
                continue

            original_name = os.path.basename(doc.archivo.name)
            final_name = original_name
            counter = 1
            while final_name in used_names:
                base_name, extension = os.path.splitext(original_name)
                final_name = f"{base_name}_{counter}{extension}"
                counter += 1
            used_names.add(final_name)
            zip_file.write(doc.archivo.path, arcname=final_name)
            archivos_anadidos += 1

    if archivos_anadidos == 0:
        messages.error(request, 'Ninguno de los documentos seleccionados esta disponible para descargar.')
        return redirect('repositorio_estudio', id_estudio=estudio.id)

    zip_buffer.seek(0)
    response = HttpResponse(zip_buffer.getvalue(), content_type='application/zip')
    response['Content-Disposition'] = f'attachment; filename="{zip_filename}"'
    return response


@login_required
@permission_required('muestras.can_change_estudios_web')
def eliminar_documento(request, *_args):
    """Elimina documentos seleccionados del estudio activo y vuelve al repositorio."""
    ids_documento = request.POST.getlist('doc_id')
    id_estudio = request.POST.get('id_estudio') or request.session.get('id')
    if not id_estudio:
        return redirect('estudios_todos')

    if not ids_documento:
        messages.info(request, 'No se ha seleccionado ningún documento.')
        return redirect('repositorio_estudio', id_estudio=id_estudio)

    estudio = _get_estudio_accesible(request, id_estudio)
    eliminados = 0
    errores = 0
    for element in ids_documento:
        try:
            doc = Documento.objects.get(pk=element, eliminado=False, estudio=estudio)
            doc.delete()
            eliminados += 1
        except Documento.DoesNotExist:
            errores += 1
        except Exception:
            errores += 1
    if eliminados:
        messages.success(request, f'{eliminados} documento(s) eliminado(s) correctamente.')
    if errores:
        messages.error(request, f'No se pudieron eliminar {errores} documento(s).')
    return redirect('repositorio_estudio', id_estudio=id_estudio)
