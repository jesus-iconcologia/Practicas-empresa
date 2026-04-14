from ..models import Localizacion, Subposicion, historial_localizaciones


def liberar_subposicion(subposicion):
    if not subposicion:
        return

    subposicion.vacia = True
    subposicion.muestra = None
    subposicion.save()


def ocupar_subposicion(subposicion, muestra):
    subposicion.vacia = False
    subposicion.muestra = muestra
    subposicion.save()


def crear_o_actualizar_localizacion(muestra, subposicion):
    caja = subposicion.caja
    rack = caja.rack
    estante = rack.estante
    congelador = estante.congelador
    bandeja = getattr(caja, "bandeja", None)

    localizacion, _ = Localizacion.objects.update_or_create(
        muestra=muestra,
        defaults={
            "congelador": congelador.congelador,
            "estante": estante.numero,
            "rack": rack.numero,
            "bandeja": bandeja.numero if bandeja else None,
            "caja": caja.numero,
            "fila": subposicion.fila,
            "columna": subposicion.columna,
            "subposicion": subposicion.numero,
        },
    )
    return localizacion


def registrar_historial_localizacion(muestra, usuario, localizacion):
    return historial_localizaciones.objects.create(
        muestra=muestra,
        localizacion=localizacion,
        usuario_asignacion=usuario,
    )


def mover_muestra_a_subposicion(muestra, nueva_subposicion, usuario=None):
    subposicion_actual = getattr(muestra, "subposicion", None)
    if subposicion_actual:
        liberar_subposicion(subposicion_actual)

    ocupar_subposicion(nueva_subposicion, muestra)
    localizacion = crear_o_actualizar_localizacion(muestra, nueva_subposicion)

    if usuario is not None:
        registrar_historial_localizacion(muestra, usuario, localizacion)

    return localizacion


def eliminar_localizacion_de_muestra(muestra):
    Localizacion.objects.filter(muestra=muestra).delete()
    subposicion_actual = getattr(muestra, "subposicion", None)
    if subposicion_actual:
        liberar_subposicion(subposicion_actual)
