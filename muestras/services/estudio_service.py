from django.db import connection, transaction
from django.utils import timezone

from ..models import Estudio, Muestra, historial_estudios


def asociar_muestra_a_estudio(muestra: Muestra, estudio: Estudio, usuario=None):
    muestra.estudio = estudio
    muestra.save()

    return historial_estudios.objects.create(
        muestra=muestra,
        estudio=estudio,
        fecha_asignacion=timezone.now(),
        usuario_asignacion=usuario,
    )


def asociar_muestras_a_estudio(muestras, estudio, usuario=None):
    historial = []
    for muestra in muestras:
        historial.append(asociar_muestra_a_estudio(muestra, estudio, usuario))
    return historial


def renombrar_estudio_y_migrar_muestras(estudio: Estudio, nuevo_nombre: str):
    """
    Actualiza las referencias de `Muestra.estudio` antes de renombrar el estudio.

    `Muestra.estudio` apunta a `Estudio.nombre_estudio` mediante `to_field`, por lo
    que cambiar el nombre requiere sincronizar manualmente la columna subyacente.
    """
    if nuevo_nombre == estudio.nombre_estudio:
        return

    with connection.cursor() as cursor:
        cursor.execute("SET FOREIGN_KEY_CHECKS=0")

    try:
        with transaction.atomic():
            with connection.cursor() as cursor:
                cursor.execute(
                    "UPDATE muestras_muestra SET estudio_id = %s WHERE estudio_id = %s",
                    [nuevo_nombre, estudio.nombre_estudio],
                )
    finally:
        with connection.cursor() as cursor:
            cursor.execute("SET FOREIGN_KEY_CHECKS=1")
