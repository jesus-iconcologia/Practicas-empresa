from django.http import JsonResponse
from django.contrib.auth.decorators import login_required
from django.urls import reverse

from ..models import Bandeja, Estante, Rack, Caja, Subposicion


@login_required
def get_estantes_por_congelador(request):
    congelador_nombre = request.GET.get('congelador', '').strip()
    if not congelador_nombre:
        return JsonResponse({'error': 'Congelador no especificado'}, status=400)

    estantes = (
        Estante.objects
        .filter(congelador__congelador__iexact=congelador_nombre)
        .values('id', 'numero')
        .order_by('numero')
    )
    return JsonResponse({'estantes': list(estantes)})


@login_required
def get_racks_por_estante(request):
    estante_id = request.GET.get('estante_id')
    if not estante_id:
        return JsonResponse({'error': 'Estante no especificado'}, status=400)

    racks = (
        Rack.objects
        .filter(estante_id=estante_id)
        .values('id', 'numero')
        .order_by('numero')
    )
    return JsonResponse({'racks': list(racks)})


@login_required
def get_cajas_por_rack(request):
    rack_id = request.GET.get('rack_id')
    if not rack_id:
        return JsonResponse({'error': 'Rack no especificado'}, status=400)

    cajas = (
        Caja.objects
        .filter(rack_id=rack_id)
        .values('id', 'numero')
        .order_by('numero')
    )
    return JsonResponse({'cajas': list(cajas)})


@login_required
def get_bandejas_por_rack(request):
    rack_id = request.GET.get('rack_id')
    if not rack_id:
        return JsonResponse({'error': 'Rack no especificado'}, status=400)

    bandejas = (
        Bandeja.objects
        .filter(rack_id=rack_id)
        .values('id', 'numero')
        .order_by('numero')
    )
    return JsonResponse({'bandejas': list(bandejas)})


@login_required
def get_cajas_por_bandeja(request):
    bandeja_id = request.GET.get('bandeja_id')
    if not bandeja_id:
        return JsonResponse({'error': 'Bandeja no especificada'}, status=400)

    cajas = (
        Caja.objects
        .filter(bandeja_id=bandeja_id)
        .values('id', 'numero')
        .order_by('numero')
    )
    return JsonResponse({'cajas': list(cajas)})


@login_required
def get_subposiciones_por_caja(request):
    caja_id = request.GET.get('caja_id')
    if not caja_id:
        return JsonResponse({'error': 'Caja no especificada'}, status=400)

    subposiciones = (
        Subposicion.objects
        .filter(caja_id=caja_id, vacia=True)
        .values('id', 'fila', 'columna', 'numero')
        .order_by('fila', 'columna')
    )
    resultado = []
    for sub in subposiciones:
        resultado.append({
            'id': sub['id'],
            'fila': sub['fila'],
            'columna': sub['columna'],
            'numero': f"Fila {sub['fila']}, Columna {sub['columna']}",
        })
    return JsonResponse({'subposiciones': resultado})


@login_required
def get_subposiciones_por_caja_tree(request):
    caja_id = request.GET.get('caja_id')
    if not caja_id:
        return JsonResponse({'error': 'Caja no especificada'}, status=400)

    subposiciones = (
        Subposicion.objects
        .filter(caja_id=caja_id)
        .select_related('muestra')
        .order_by('fila', 'columna')
    )

    resultado = []
    for s in subposiciones:
        item = {
            'id': s.id,
            'numero': s.etiqueta,
            'fila': s.fila,
            'columna': s.columna,
            'vacia': s.vacia,
        }
        if not s.vacia and s.muestra:
            item['muestra_nom_lab'] = s.muestra.nom_lab
            item['muestra_estado'] = s.muestra.get_estado_actual_display() or ''
            item['muestra_detalle_url'] = reverse('detalles_muestra', args=[s.muestra.nom_lab])
        resultado.append(item)

    return JsonResponse({'subposiciones': resultado})
