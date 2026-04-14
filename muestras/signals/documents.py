import os
import shutil

from django.conf import settings
from django.db.models.signals import post_delete
from django.dispatch import receiver

from ..models.estudio import Documento, Estudio


@receiver(post_delete, sender=Documento)
def eliminar_archivo_documento(sender, instance, **kwargs):
    try:
        if instance.archivo:
            instance.archivo.delete(save=False)
    except Exception:
        pass


@receiver(post_delete, sender=Estudio)
def eliminar_carpeta_estudio(sender, instance, **kwargs):
    try:
        carpeta = os.path.join(settings.MEDIA_ROOT, 'estudios', str(instance.id))
        if os.path.exists(carpeta) and os.path.isdir(carpeta):
            shutil.rmtree(carpeta)
    except Exception:
        pass