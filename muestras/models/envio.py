from django.db import models
from django.utils import timezone
from django.contrib.auth.models import User


class Envio(models.Model):
    muestra = models.ForeignKey('Muestra', related_name="envio", on_delete=models.CASCADE)
    volumen_enviado = models.FloatField()
    unidad_volumen_enviado = models.CharField(max_length=15)
    concentracion_enviada = models.FloatField()
    unidad_concentracion_enviada = models.CharField(max_length=15)
    centro_destino = models.CharField(max_length=100)
    lugar_destino = models.CharField(max_length=100)
    fecha_envio = models.DateField(default=timezone.now)
    usuario_envio = models.ForeignKey(User, on_delete=models.PROTECT, blank=True, null=True)

    def __str__(self):
        return f"Envio de Muestra {self.muestra.id_individuo} - {self.muestra.nom_lab} el {self.fecha_envio}"


class agenda_envio(models.Model):
    centro = models.CharField(max_length=200, unique=True, default=None)
    lugar = models.CharField(max_length=200)
    direccion = models.TextField()
    persona_contacto = models.CharField(max_length=200, blank=True, null=True)
    telefono_contacto = models.IntegerField(blank=True, null=True)