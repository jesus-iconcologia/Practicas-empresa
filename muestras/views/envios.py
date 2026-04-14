import base64
import io
import os

import openpyxl
import pandas as pd
from django.conf import settings
from django.contrib import messages
from django.contrib.auth.decorators import login_required, permission_required
from django.db import transaction
from django.http import HttpResponse, StreamingHttpResponse
from django.shortcuts import redirect, render
from django.template import loader
from django.utils import timezone
from openpyxl.styles import PatternFill

from ..forms import CentroForm, UploadExcel
from ..models import Envio, Localizacion, Muestra, Subposicion, agenda_envio
from .common import _json_progress, _should_update

# Vistas relacionadas con el envio de muestras
@login_required
@permission_required('muestras.can_change_muestras_web')
def formulario_envios(request,centro):
    # Vista para mostrar el formulario de envíos de muestras a un centro específico
    muestras_envio = request.session.get('muestras_envio', [])
    centro_envio = agenda_envio.objects.get(id=centro)
    muestras = Muestra.objects.filter(id__in=muestras_envio, volumen_actual__gt=0)
    template = loader.get_template('formulario_envios.html')
    return HttpResponse(template.render({'muestras':muestras,'centro':centro_envio},request))

@login_required
@permission_required('muestras.can_change_muestras_web')
def registrar_envio(request,centro):
    # Vista para registrar el envío de muestras a un centro específico desde el formulario de envíos
    if request.method=='POST':
        # Obtener los datos del formulario, guardados en la sesión y registrar los envíos
        centro_envio = agenda_envio.objects.get(id=centro)
        muestras = request.POST.getlist('muestra_id')
        volumen_enviado_form = request.POST.getlist('volumen_enviado')
        concentracion_enviada_form = request.POST.getlist('concentracion_enviada')
        centro_destino_form = centro_envio.centro
        lugar_destino_form = centro_envio.lugar
        if not muestras:
            messages.error(request, 'No hay muestras para registrar el envío.')
            return redirect('formulario_envios', centro=centro)
        if len(volumen_enviado_form) != len(muestras) or len(concentracion_enviada_form) != len(muestras):
            messages.error(request, 'Faltan datos de volumen o concentración para alguna muestra.')
            return redirect('formulario_envios', centro=centro)

        for idx, muestra in enumerate(muestras):
            instancia_muestra = Muestra.objects.get(id=muestra)
            envio = Envio.objects.create(
                muestra=instancia_muestra,
                fecha_envio=timezone.now(),
                volumen_enviado = volumen_enviado_form[idx],
                unidad_volumen_enviado = instancia_muestra.unidad_volumen,
                concentracion_enviada = concentracion_enviada_form[idx],
                unidad_concentracion_enviada = instancia_muestra.unidad_concentracion,
                centro_destino = centro_destino_form,
                lugar_destino=lugar_destino_form,
                usuario_envio = request.user
            )
            envio.save()
            # Actualizar el estado, la posición y el volumen de la muestra tras el envío
            if float(volumen_enviado_form[idx]) >= instancia_muestra.volumen_actual:
                instancia_muestra.volumen_actual = 0
                instancia_muestra.concentracion_actual = 0
                instancia_muestra.estado_actual = 'ENV'
                instancia_muestra.save()
                if Localizacion.objects.filter(muestra=instancia_muestra).exists():
                    loc = Localizacion.objects.get(muestra=instancia_muestra)
                    loc.muestra = None
                    loc.save()
                if Subposicion.objects.filter(muestra=instancia_muestra).exists():
                    sub = Subposicion.objects.get(muestra=instancia_muestra)
                    sub.muestra = None
                    sub.vacia = True
                    sub.save()
            else:
                instancia_muestra.volumen_actual -= float(volumen_enviado_form[idx])
                instancia_muestra.estado_actual = 'PENV'
                instancia_muestra.save()
        if 'muestras_envio' in request.session:
            del request.session['muestras_envio']
        return redirect('muestras_todas')
    return redirect('formulario_envios')

@login_required
@permission_required('muestras.can_change_muestras_web')
def upload_excel_envios(request,centro):
    # Mostrar confirmación después de validación con progreso
    if request.GET.get('mostrar_confirmacion') and 'confirmacion_pendiente' in request.session:
        datos_conf = request.session.pop('confirmacion_pendiente')
        for msg in datos_conf.get('mensajes', []):
            getattr(messages, msg['level'])(request, msg['text'])
        return render(request, datos_conf['template'], datos_conf.get('context', {}))
    # Vista para subir un archivo Excel con los datos de envío de muestras
    centro_envio = agenda_envio.objects.get(id=centro)
    if request.method=='POST':
        form = UploadExcel(request.POST, request.FILES)
        if 'confirmar' in request.POST:
            # Si el usuario confirma, se registran los envíos en la base de datos
            filas_validas = request.session.get('filas_validas',[])
            total = len(filas_validas)

            def gen_confirmar_envios():
                yield _json_progress(0, total, 'start')
                try:
                    with transaction.atomic():
                        for i, datos in enumerate(filas_validas):
                            muestra = Muestra.objects.get(nom_lab=datos['nom_lab'])
                            Envio.objects.create(
                                muestra=muestra,
                                volumen_enviado=datos['volumen_enviado'],
                                unidad_volumen_enviado=datos['unidad_volumen_enviado'],
                                concentracion_enviada=datos['concentracion_enviada'],
                                unidad_concentracion_enviada=datos['unidad_concentracion_enviada'],
                                centro_destino=datos['centro_destino'],
                                lugar_destino=datos['lugar_destino'],
                                fecha_envio=timezone.now(),
                                usuario_envio=request.user
                            )
                            if datos['volumen_enviado'] >= muestra.volumen_actual:
                                muestra.volumen_actual = 0
                                muestra.concentracion_actual = 0
                                muestra.estado_actual = 'ENV'
                                muestra.save()
                                if Localizacion.objects.filter(muestra=muestra).exists():
                                    loc = Localizacion.objects.get(muestra=muestra)
                                    loc.muestra = None
                                    loc.save()
                                if Subposicion.objects.filter(muestra=muestra).exists():
                                    sub = Subposicion.objects.get(muestra=muestra)
                                    sub.muestra = None
                                    sub.vacia = True
                                    sub.save()
                            else:
                                muestra.volumen_actual -= float(datos['volumen_enviado'])
                                muestra.estado_actual = 'PENV'
                                muestra.save()
                            if _should_update(i, total):
                                yield _json_progress(i + 1, total, 'processing')
                    yield _json_progress(total, total, 'done', 'Envíos registrados correctamente', '/muestras/')
                except Exception as e:
                    yield _json_progress(0, total, 'error', str(e))

            return StreamingHttpResponse(gen_confirmar_envios(), content_type='application/x-ndjson')
        
        elif 'cancelar' in request.POST:
            # Si el usuario cancela, no se registra nada
            messages.error(request,'El envio no se ha registrado')
            return redirect('muestras_todas')
        
        elif 'descargar_excel_envio' in request.POST:
            # Si se solicita descargar el excel de envío, se genera y se rellena con los datos de las muestras a enviar, en caso de que se hayan seleccionado y no estén destruidas
            muestras = request.session.get('muestras_envio',[])
            response = HttpResponse(content_type='application/ms-excel')
            response['Content-Disposition'] = 'attachment; filename="listado_envio.xlsx"'
            wb = openpyxl.load_workbook(
                os.path.join(
                    settings.BASE_DIR,
                    'geslab_project',
                    'globalstaticfiles',
                    'files',
                    'plantilla_envios.xlsx',
                )
            )
            ws = wb.active
            row_num = 2
            for muestra in muestras:
                sample = Muestra.objects.get(id=muestra)
                if sample.estado_actual != 'DEST':
                    ws.cell(row_num,1).value=str(sample.nom_lab)
                    if sample.volumen_actual != None:
                        ws.cell(row_num,2).value=str(sample.volumen_actual) + ' ' + str(sample.unidad_volumen)
                    if sample.concentracion_actual != None:
                        ws.cell(row_num,3).value=str(sample.concentracion_actual) + ' ' + str(sample.unidad_concentracion)
                    ws.cell(row_num,5).value=str(sample.unidad_volumen)
                    ws.cell(row_num,7).value=str(sample.unidad_concentracion)
                    ws.cell(row_num,8).value=str(centro_envio.centro)
                    ws.cell(row_num,9).value=str(centro_envio.lugar)
                    row_num +=1 
            wb.save(response)
            return response
        elif 'excel_file' in request.FILES:
            # Limpiar sesión residual de uploads anteriores
            if 'columnas_adicionales' in request.session:
                del request.session['columnas_adicionales']
            # Si se sube un archivo excel, se procesa y valida
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
                    return render(request, 'upload_excel_envios.html', {'form': form, 'error': f'❌ Error al leer el archivo: {str(e)}'})
                rename_columns = {
                    'Muestra':'nom_lab',
                    'Volumen enviado':'volumen_enviado', 
                    'Volumen actual': 'volumen_actual',
                    'Concentración actual':'concentracion_actual',
                    'Concentración enviada':'concentracion_enviada',
                    'Unidad de volumen':'unidad_volumen_enviado',
                    'Unidad de concentración':'unidad_concentracion_enviada',
                    'Centro de destino':'centro_destino',
                    'Lugar de destino':'lugar_destino'
                }
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
           
                
                # Preparar datos para comprobaciones y variables previas

                cache = {
                    'muestras': Muestra.objects.values_list('nom_lab',flat=True),
                    'volumenes_actuales':{
                        sample.nom_lab : sample.volumen_actual
                        for sample in Muestra.objects.all() if sample.volumen_actual != None 
                    },
                    'estados_actuales': {
                        sample.nom_lab
                        for sample in Muestra.objects.all()
                        if sample.estado_actual in {'DISP', 'PENV'}
                    },
                    'centros_envio': agenda_envio.objects.values_list('centro','lugar')
                }

                total_filas = len(df)

                def gen_validar_envios():
                    filas_validas = []
                    errores = {}
                    nom_lab_excel = set()
                    numero_registros = 0

                    yield _json_progress(0, total_filas, 'start')

                    # Recorrer las filas del excel para realizar la validación previa a la carga de datos
                    for idx, row in df.iterrows():
                        numero_registros += 1
                        fila = idx + 2 
                        errores[fila]={"bloqueantes":[],"advertencias":[]}
                        # Registrar en el excel el centro y lugar de envio seleccionados de la agenda de envios
                        row['centro_destino'] = centro_envio.centro
                        row['lugar_destino'] = centro_envio.lugar

                        datos = {
                            "nom_lab":norm(row['nom_lab']),
                            "volumen_enviado":norm_code(row['volumen_enviado']),
                            "unidad_volumen_enviado":norm(row['unidad_volumen_enviado']),
                            "concentracion_enviada":norm_code(row['concentracion_enviada']),
                            "unidad_concentracion_enviada":norm(row['unidad_concentracion_enviada'])
                        }
              
                        # Comprobar si los campos obligatorios han sido rellenados
                        obligatorios = ["nom_lab", "volumen_enviado", "unidad_volumen_enviado", "concentracion_enviada", "unidad_concentracion_enviada"]

                        for campo in obligatorios:
                            if not datos.get(campo):
                                errores[fila]["bloqueantes"].append(f"campo_obligatorio_vacio:{campo}")
                    
                         # Comprobar si los campos estan en el formato correcto
                        for campo in ['volumen_enviado', 'concentracion_enviada']:
                            if datos[campo] != None:
                                try:
                                    datos[campo]=float(datos[campo])
                                except (TypeError, ValueError):
                                    errores[fila]["bloqueantes"].append(f"formato_incorrecto:{campo}")
                    
                        # Comprobar que la muestra exista en la base de datos y no esté duplicada en el excel
                        nom_lab = datos["nom_lab"]
                        if nom_lab not in cache["muestras"]:
                            errores[fila]["bloqueantes"].append("muestra_inexistente")
                        if nom_lab in nom_lab_excel:
                            errores[fila]["bloqueantes"].append("muestra_duplicada_excel")
                        else:
                            nom_lab_excel.add(nom_lab)

                        # Comprobar que el volumen a enviar no sea mayor al actual
                        volumen_envio = datos["volumen_enviado"]
                        if nom_lab in cache["volumenes_actuales"]:
                            if volumen_envio > cache["volumenes_actuales"][nom_lab]:
                                errores[fila]["bloqueantes"].append("volumen_alto")

                        # Comprobar que el estado de la muestra sea 'Disponible' o 'Parcialmente enviada'
                        if nom_lab not in cache['estados_actuales']:
                            errores[fila]["bloqueantes"].append("estado_no_disponible")

                        # Rellenar con el centro y lugar de destino
                        datos['centro_destino'] = centro_envio.centro
                        datos['lugar_destino'] = centro_envio.lugar
                    
                        # Registrar filas validas
                        if not errores[fila]["bloqueantes"]:
                            datos['fila'] = fila
                            filas_validas.append(datos)

                        if _should_update(idx, total_filas):
                            yield _json_progress(idx + 1, total_filas, 'processing')
                
                    request.session['filas_validas'] = filas_validas
                    request.session['errores'] = errores

                    mensajes = [{'level': 'info', 'text': f'El excel subido tiene {numero_registros} registros.'}]
                    numero_errores = sum(1 for f in errores if errores[f]['bloqueantes'])
                    if numero_errores > 0:
                        mensajes.append({'level': 'error', 'text': f'Pero contiene {numero_errores} filas con errores graves.'})
                    else:
                        mensajes.append({'level': 'success', 'text': 'Y no tiene errores en ningún campo.'})

                    request.session['confirmacion_pendiente'] = {
                        'template': 'confirmacion_upload_envio.html',
                        'context': {},
                        'mensajes': mensajes
                    }
                    request.session.save()
                    yield _json_progress(total_filas, total_filas, 'done', 'Validación completada', request.path + '?mostrar_confirmacion=1')

                return StreamingHttpResponse(gen_validar_envios(), content_type='application/x-ndjson')
        # Si se solicita un excel de errores, este se rellena en base a los errores detectados durante la validación 
        elif 'excel_errores' in request.POST:
                # Leer los errores y el excel de la sesión
                errores = request.session.get('errores',[])
                excel_bytes = base64.b64decode(request.session.get('excel_file_base64'))
                excel_file = io.BytesIO(excel_bytes)
                wb = openpyxl.load_workbook(excel_file)
                ws = wb.active
                # Definir los estilos para pintar el excel
                FILL_ERROR_CELL = PatternFill("solid", fgColor="F5C2C7")  # rojo fuerte
                # Diccionario de mensajes
                MENSAJES_ERROR = {
                    "campo_obligatorio_vacio": "Campo obligatorio vacío",
                    "formato_incorrecto": "Formato incorrecto de un campo",
                    "fecha_invalida": "Fecha inválida (Formato correcto: DD-MM-AAAA)",
                    "muestra_inexistente": "La muestra no existe en la base de datos",
                    "muestra_duplicada_excel": "Muestra duplicada dentro del Excel",
                    "volumen_alto": "La muestra no tiene suficiente volumen para el envio",
                    "estado_no_disponible": "La muestra está enviada o destruida, o no tiene un estado definido",
                    "caracter_invalido_semicolon": "El carácter ';' no está permitido en este campo"
                }
                # Diccionario de columnas del excel
                columnas_excel = {}
                rename_columns = {
                    'Muestra':'nom_lab',
                    'Volumen enviado':'volumen_enviado', 
                    'Volumen actual': 'volumen_actual',
                    'Concentración actual':'concentracion_actual',
                    'Concentración enviada':'concentracion_enviada',
                    'Unidad de volumen':'unidad_volumen_enviado',
                    'Unidad de concentración':'unidad_concentracion_enviada',
                    'Centro de destino':'centro_destino',
                    'Lugar de destino':'lugar_destino'
                }
                for cell in ws[1]:
                    columnas_excel[rename_columns[cell.value]] = cell.column
                # Añadir la columna de errores
                col_errores = ws.max_column + 1
                ws.cell(row=1, column=col_errores, value="Errores")
                # Mapeo de errores sin campo a sus columnas específicas (envíos)
                error_campo_map_envio = {
                    "muestra_inexistente": "nom_lab",
                    "muestra_duplicada_excel": "nom_lab",
                    "volumen_alto": "volumen_enviado",
                    "estado_no_disponible": "nom_lab"
                }
                # Recorrer filas con errores 
                for fila, info in errores.items():
                    has_error = bool(info["bloqueantes"])
                    if not has_error:
                        continue

                    # Colorear celdas específicas y construir mensajes
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
                            campo = error_campo_map_envio.get(err)
                            msg = f"(Error) {MENSAJES_ERROR[err]}"
                            if msg not in mensajes:
                                mensajes.append(msg)
                            if campo and campo in columnas_excel:
                                col_err = columnas_excel[campo]
                                ws.cell(row=int(fila), column=col_err).fill = FILL_ERROR_CELL
                    ws.cell(row=int(fila), column=col_errores, value="\n".join(mensajes))
                # Retornar el excel de errores 
                output = io.BytesIO()    
                wb.save(output)
                wb.close()
                response = HttpResponse(output.getvalue(),content_type='application/ms-excel')
                response['Content-Disposition'] = 'attachment; filename="listado_errores_envio.xlsx"'
                return response         
    else:
        form = UploadExcel(request)
    template = loader.get_template('upload_excel_envios.html')     
    return HttpResponse(template.render({'form': form},request))


@login_required
@permission_required('muestras.can_view_muestras_web')
def historial_envios(request,muestra_id):
    # Vista para ver el historial de envíos de una muestra específica
    sample = Muestra.objects.get(id=muestra_id)
    envios = Envio.objects.filter(muestra=sample).order_by('-fecha_envio')
    # Calcular el volumen original y el volumen restante
    volumen_original = sample.volumen_actual + sum(envio.volumen_enviado for envio in envios)
    volumen_restante = sample.volumen_actual
    template = loader.get_template('historial_envios.html')
    context = {
        'muestra':sample,
        'envios':envios,
        'volumen_original':volumen_original,
        'volumen_restante':volumen_restante
    }
    return HttpResponse(template.render(context,request))

@login_required
@permission_required('muestras.can_change_muestras_web')
def agenda(request):
    # Vista para ver la agenda de envíos de muestras
    agenda_envios = agenda_envio.objects.all()
    template = loader.get_template('agenda.html')
    return HttpResponse(template.render({'agenda':agenda_envios},request))

@login_required
@permission_required('muestras.can_change_muestras_web')
def nuevo_centro(request):
    # Vista para añadir un nuevo centro a la agenda de envíos
    if request.method == 'POST':
        form = CentroForm(request.POST)
        if form.is_valid():
            form.save()
            return redirect('agenda')
        else:
            messages.error(request, 'Hubo un error al añadir el centro.')
    else:
        form = CentroForm()
    template = loader.get_template('nuevo_centro.html')
    return HttpResponse(template.render({'form':form},request))

@login_required
@permission_required('muestras.can_change_muestras_web')
def editar_centro(request, id_centro):
    # Vista para editar un centro existente en la agenda de envíos
    centro = agenda_envio.objects.get(id=id_centro)
    if request.method == 'POST':
        form = CentroForm(request.POST, instance=centro)
        if form.is_valid():
            form.save()
            return redirect('agenda')
    else:
        form = CentroForm(instance=centro)
    return render(request, 'editar_centro.html', {'form': form, 'centro': centro})

@login_required
@permission_required('muestras.can_change_muestras_web')
def eliminar_centro(request):
    # Vista para eliminar centros de la agenda de envíos
    if request.method=="POST":
        ids = request.POST.getlist('ids_centro')
        for centro_id in ids:
            centro = agenda_envio.objects.get(id=centro_id)
            centro.delete()
    return redirect('agenda')
