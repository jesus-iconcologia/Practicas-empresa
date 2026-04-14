from django.contrib.auth.models import User
from django.db import models
from django.utils import timezone


class Material(models.Model):
    nombre = models.CharField(max_length=100, unique=True)
    activo = models.BooleanField(default=True)
    usa_volumen = models.BooleanField(default=False)
    unidades_volumen = models.CharField(max_length=250, blank=True, null=True)
    usa_concentracion = models.BooleanField(default=False)
    unidades_concentracion = models.CharField(max_length=250, blank=True, null=True)
    usa_unidades_bloque = models.BooleanField(default=False)
    unidades_recuento = models.CharField(max_length=250, blank=True, null=True)
    notas = models.TextField(blank=True, null=True)

    class Meta:
        ordering = ["nombre"]
        verbose_name = "Material"
        verbose_name_plural = "Materiales"

    def __str__(self):
        return self.nombre

    def unidades_volumen_list(self):
        return [item.strip() for item in (self.unidades_volumen or "").split(",") if item.strip()]

    def unidades_concentracion_list(self):
        return [item.strip() for item in (self.unidades_concentracion or "").split(",") if item.strip()]

    def unidades_recuento_list(self):
        return [item.strip() for item in (self.unidades_recuento or "").split(",") if item.strip()]


class Muestra(models.Model):
    id_individuo = models.CharField(max_length=20, blank=True, null=True)
    nom_lab = models.CharField(
        max_length=100,
        unique=True,
        help_text="Nombre unico de la muestra asignado por el laboratorio",
    )
    material = models.ForeignKey(
        "Material",
        blank=True,
        null=True,
        on_delete=models.SET_NULL,
        related_name="muestras",
    )
    id_material = models.CharField(max_length=20, blank=True, null=True, help_text="Material de la muestra")
    volumen_actual = models.FloatField(blank=True, null=True)
    unidad_volumen = models.CharField(max_length=15, blank=True, null=True)
    concentracion_actual = models.FloatField(blank=True, null=True)
    unidad_concentracion = models.CharField(max_length=15, blank=True, null=True)
    masa_actual = models.FloatField(blank=True, null=True)
    unidad_masa = models.CharField(max_length=15, blank=True, null=True)
    fecha_extraccion = models.DateField(blank=True, null=True)
    fecha_llegada = models.DateField(blank=True, null=True)
    observaciones = models.TextField(blank=True, null=True)
    estado_inicial = models.CharField(max_length=50, blank=True, null=True)
    centro_procedencia = models.CharField(max_length=100, blank=True, null=True)
    lugar_procedencia = models.CharField(max_length=100, blank=True, null=True)
    estado_actual = models.CharField(
        max_length=50,
        default="DISP",
        choices=[
            ("DISP", "Disponible"),
            ("ENV", "Enviada"),
            ("PENV", "Parcialmente enviada"),
            ("DEST", "Destruida"),
        ],
        blank=True,
        null=True,
        help_text="Disponible por defecto",
    )
    estudio = models.ForeignKey(
        "Estudio",
        blank=True,
        to_field="nombre_estudio",
        on_delete=models.SET_NULL,
        null=True,
    )

    class Meta:
        permissions = [
            ("can_view_muestras_web", "Puede ver muestras en la web"),
            ("can_add_muestras_web", "Puede añadir muestras en la web"),
            ("can_change_muestras_web", "Puede cambiar muestras en la web"),
            ("can_delete_muestras_web", "Puede eliminar muestras en la web"),
        ]

    def posicion_completa(self):
        sub = getattr(self, "subposicion", None)
        if not sub:
            return None

        caja = sub.caja
        bandeja = caja.bandeja
        rack = caja.rack
        estante = rack.estante
        congelador = estante.congelador

        return (
            f"{congelador.congelador}-{estante.numero}-"
            f"{rack.posicion_rack_estante}-{rack.numero}-"
            f"{(bandeja.posicion_bandeja_rack if bandeja else '')}-"
            f"{(bandeja.numero if bandeja else '')}-"
            f"{caja.posicion_caja_rack}-{caja.numero}-"
            f"{sub.fila}-{sub.columna}"
        )

    def save(self, *args, **kwargs):
        if self.material is not None:
            self.id_material = self.material.nombre
        else:
            self.id_material = None
        super().save(*args, **kwargs)


class Localizacion(models.Model):
    muestra = models.ForeignKey(
        "Muestra",
        to_field="nom_lab",
        related_name="localizacion",
        blank=True,
        null=True,
        on_delete=models.SET_NULL,
    )
    congelador = models.CharField(max_length=50, blank=True, null=True)
    estante = models.CharField(max_length=50, blank=True, null=True)
    rack = models.CharField(max_length=50, blank=True, null=True)
    bandeja = models.CharField(max_length=50, blank=True, null=True)
    caja = models.CharField(max_length=50, blank=True, null=True)
    fila = models.CharField(max_length=50, blank=True, null=True)
    columna = models.CharField(max_length=50, blank=True, null=True)
    subposicion = models.CharField(max_length=50, blank=True, null=True)

    class Meta:
        permissions = [
            ("can_view_localizaciones_web", "Puede ver localizaciones en la web"),
            ("can_add_localizaciones_web", "Puede añadir localizaciones en la web"),
            ("can_change_localizaciones_web", "Puede cambiar localizaciones en la web"),
            ("can_delete_localizaciones_web", "Puede eliminar localizaciones en la web"),
        ]

    def __str__(self):
        partes = [self.congelador, self.estante, self.rack]
        if self.bandeja:
            partes.append(self.bandeja)
        partes.append(self.caja)
        if self.fila or self.columna:
            partes.append(f"Fila {self.fila}, Columna {self.columna}")
        elif self.subposicion:
            partes.append(self.subposicion)
        return " - ".join([str(parte) for parte in partes if parte not in (None, "")])


class historial_localizaciones(models.Model):
    muestra = models.ForeignKey("Muestra", related_name="historial_localizaciones", on_delete=models.CASCADE)
    localizacion = models.ForeignKey(
        "Localizacion",
        related_name="historial_localizaciones",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
    )
    fecha_asignacion = models.DateField(default=timezone.now)
    usuario_asignacion = models.ForeignKey(User, on_delete=models.PROTECT, blank=True, null=True)


class registro_destruido(models.Model):
    muestra = models.ForeignKey("Muestra", related_name="estado_destruido", on_delete=models.CASCADE)
    fecha = models.DateField(default=timezone.now)
    usuario = models.ForeignKey(User, on_delete=models.PROTECT, blank=True, null=True)
    motivo = models.CharField(max_length=250, blank=True, null=True)
    metodo = models.CharField(max_length=250, blank=True, null=True)
    lugar = models.CharField(max_length=250, blank=True, null=True)
    responsable = models.CharField(max_length=250, blank=True, null=True)
    tecnico = models.CharField(max_length=250, blank=True, null=True)
    observaciones = models.TextField(blank=True, null=True)
