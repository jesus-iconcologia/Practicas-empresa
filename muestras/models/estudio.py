from django.db import models
from django.utils import timezone
from django.contrib.auth.models import User
from django.conf import settings


class Estudio(models.Model):
    """Entidad principal para agrupar muestras y documentación asociada."""

    referencia_estudio = models.CharField(max_length=100, blank=True, null=True)
    nombre_estudio = models.CharField(max_length=100, unique=True)
    descripcion_estudio = models.TextField(blank=True, null=True)
    fecha_inicio_estudio = models.DateField(blank=True, null=True)
    fecha_fin_estudio = models.DateField(blank=True, null=True)
    investigador_principal = models.CharField(max_length=100, blank=True, null=True)
    investigadores_asociados = models.ManyToManyField(
        settings.AUTH_USER_MODEL,
        limit_choices_to={'groups__name': 'Investigadores'},
        related_name="estudios_asignados",
        blank=True
    )

    class Meta:
        permissions = [
            ("can_view_estudios_web", "Puede ver estudios en la web"),
            ("can_add_estudios_web", "Puede añadir estudios en la web"),
            ("can_change_estudios_web", "Puede cambiar estudios en la web"),
            ("can_delete_estudios_web", "Puede eliminar estudios en la web"),
        ]

    def __str__(self):
        return f"Estudio {self.nombre_estudio}"


class historial_estudios(models.Model):
    """Trazabilidad de cambios de estudio sobre cada muestra."""

    muestra = models.ForeignKey('Muestra', related_name="historial_estudios", on_delete=models.CASCADE)
    estudio = models.ForeignKey('Estudio', related_name="historial_estudios", on_delete=models.SET_NULL, blank=True, null=True)
    fecha_asignacion = models.DateField(default=timezone.now)
    usuario_asignacion = models.ForeignKey(User, on_delete=models.PROTECT, blank=True, null=True)


def ruta_documentos(instance, filename):
    """Organiza los ficheros por estudio para simplificar su gestión en disco."""
    return f"estudios/{instance.estudio.id}/{filename}"


class Documento(models.Model):
    """Archivo subido al repositorio documental de un estudio."""

    estudio = models.ForeignKey('Estudio', related_name="estudio", on_delete=models.CASCADE)
    archivo = models.FileField(upload_to=ruta_documentos)
    fecha_subida = models.DateTimeField(auto_now_add=True)
    categoria = models.CharField(blank=True, null=True, max_length=50)
    usuario_subida = models.ForeignKey(User, on_delete=models.PROTECT)
    descripcion = models.TextField(blank=True, null=True)
    eliminado = models.BooleanField(default=False)
    fecha_eliminacion = models.DateField(blank=True, null=True)

    def delete(self, *args, **kwargs):
        """Elimina tambien el fichero fisico cuando se borra el registro."""
        archivo = self.archivo
        storage = archivo.storage if archivo else None
        archivo_name = archivo.name if archivo else None
        super().delete(*args, **kwargs)
        if (
            archivo_name
            and storage
            and not Documento.objects.filter(archivo=archivo_name).exists()
            and storage.exists(archivo_name)
        ):
            storage.delete(archivo_name)
