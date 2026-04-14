from django.core.exceptions import ValidationError
from django.db import models
from .muestra import Muestra, Localizacion


class Congelador(models.Model):
    """Recurso físico de almacenaje con una estructura jerárquica fija."""

    ESTRUCTURA_CONGELADOR_80 = "congelador_80"
    ESTRUCTURA_CONGELADOR_20 = "congelador_20"
    ESTRUCTURA_NEVERA = "nevera"
    ESTRUCTURA_CHOICES = [
        (ESTRUCTURA_CONGELADOR_80, "Congelador -80"),
        (ESTRUCTURA_CONGELADOR_20, "Congelador -20"),
        (ESTRUCTURA_NEVERA, "Nevera"),
    ]

    congelador = models.CharField(max_length=50, unique=True)
    tipo_estructura = models.CharField(
        max_length=30,
        choices=ESTRUCTURA_CHOICES,
        blank=True,
        null=True,
    )
    modelo = models.CharField(max_length=50, blank=True, null=True)
    ano_compra = models.PositiveIntegerField(blank=True, null=True)
    temperatura_objetivo = models.CharField(max_length=50, blank=True, null=True)
    temperatura_minima = models.CharField(max_length=50, blank=True, null=True)
    temperatura_maxima = models.CharField(max_length=50, blank=True, null=True)
    localizacion_edificio = models.CharField(max_length=50, blank=True, null=True)
    fotografia = models.ImageField(upload_to='congeladores/', blank=True, null=True)

    def save(self, *args, **kwargs):
        """Mantiene sincronizado el historial textual y limpia fotos sustituidas."""
        old_fotografia_name = None
        if self.pk:
            old = Congelador.objects.get(pk=self.pk)
            if old.congelador != self.congelador:
                Localizacion.objects.filter(congelador=old.congelador).update(congelador=self.congelador)
            if old.fotografia and old.fotografia.name != getattr(self.fotografia, "name", None):
                old_fotografia_name = old.fotografia.name
        super().save(*args, **kwargs)
        if (
            old_fotografia_name
            and not Congelador.objects.filter(fotografia=old_fotografia_name).exclude(pk=self.pk).exists()
            and self.fotografia.storage.exists(old_fotografia_name)
        ):
            self.fotografia.storage.delete(old_fotografia_name)

    def delete(self, *args, **kwargs):
        """Elimina la fotografia del disco cuando se borra el recurso."""
        fotografia = self.fotografia
        storage = fotografia.storage if fotografia else None
        fotografia_name = fotografia.name if fotografia else None
        super().delete(*args, **kwargs)
        if (
            fotografia_name
            and storage
            and not Congelador.objects.filter(fotografia=fotografia_name).exists()
            and storage.exists(fotografia_name)
        ):
            storage.delete(fotografia_name)

    def __str__(self):
        return self.congelador

    @property
    def tipo_estructura_display_safe(self):
        """Devuelve una etiqueta legible incluso si la estructura aún no está definida."""
        return self.get_tipo_estructura_display() if self.tipo_estructura else "No definida"


class ReparacionRecurso(models.Model):
    """Incidencia o actuación de mantenimiento asociada a un recurso."""

    recurso = models.ForeignKey(
        Congelador,
        on_delete=models.CASCADE,
        related_name='reparaciones',
    )
    fecha_reparacion = models.DateField()
    descripcion = models.TextField()
    coste = models.DecimalField(max_digits=10, decimal_places=2, blank=True, null=True)

    class Meta:
        ordering = ['-fecha_reparacion', '-id']

    def clean(self):
        super().clean()
        if (
            self.recurso_id
            and self.recurso.ano_compra
            and self.fecha_reparacion
            and self.fecha_reparacion.year < self.recurso.ano_compra
        ):
            raise ValidationError({
                'fecha_reparacion': (
                    'La fecha de reparación no puede ser anterior al año de compra del recurso.'
                )
            })

    def __str__(self):
        return f"{self.recurso.congelador} - {self.fecha_reparacion}"


class Estante(models.Model):
    """Primer nivel de organización dentro de un recurso."""

    congelador = models.ForeignKey(
        Congelador,
        on_delete=models.CASCADE,
        related_name='estantes',
        to_field='congelador'
    )
    numero = models.CharField(max_length=50)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['congelador', 'numero'],
                name='unique_estante_por_congelador'
            )
        ]

    def save(self, *args, **kwargs):
        """Propaga cambios de numeración al histórico de localizaciones."""
        if self.pk:
            old = Estante.objects.get(pk=self.pk)
            if old.numero != self.numero:
                Localizacion.objects.filter(
                    congelador=self.congelador.congelador,
                    estante=old.numero
                ).update(estante=self.numero)
        super().save(*args, **kwargs)


class Rack(models.Model):
    """Segundo nivel de organización dentro de un estante."""

    estante = models.ForeignKey(Estante, on_delete=models.CASCADE, related_name='racks')
    numero = models.CharField(max_length=50)
    posicion_rack_estante = models.CharField(max_length=50)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['estante', 'numero'],
                name='unique_rack_por_estante'
            ),
            models.UniqueConstraint(
                fields=['estante', 'posicion_rack_estante'],
                name='unique_posicion_rack_por_estante'
            )
        ]

    def save(self, *args, **kwargs):
        """Propaga cambios de numeración al histórico de localizaciones."""
        if self.pk:
            old = Rack.objects.get(pk=self.pk)
            if old.numero != self.numero:
                Localizacion.objects.filter(
                    congelador=self.estante.congelador.congelador,
                    estante=self.estante.numero,
                    rack=old.numero
                ).update(rack=self.numero)
        super().save(*args, **kwargs)


class Bandeja(models.Model):
    """Nivel intermedio usado por recursos de estructura completa."""

    rack = models.ForeignKey(Rack, on_delete=models.CASCADE, related_name='bandejas')
    numero = models.CharField(max_length=50)
    posicion_bandeja_rack = models.CharField(max_length=50)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['rack', 'numero'],
                name='unique_bandeja_por_rack'
            ),
            models.UniqueConstraint(
                fields=['rack', 'posicion_bandeja_rack'],
                name='unique_posicion_bandeja_por_rack'
            )
        ]

    def save(self, *args, **kwargs):
        """Propaga cambios de numeración al histórico de localizaciones."""
        if self.pk:
            old = Bandeja.objects.get(pk=self.pk)
            if old.numero != self.numero:
                Localizacion.objects.filter(
                    congelador=self.rack.estante.congelador.congelador,
                    estante=self.rack.estante.numero,
                    rack=self.rack.numero,
                    bandeja=old.numero,
                ).update(bandeja=self.numero)
        super().save(*args, **kwargs)


class Caja(models.Model):
    """Contenedor final antes de llegar a la subposición de muestra."""

    rack = models.ForeignKey(Rack, on_delete=models.CASCADE, related_name='cajas')
    bandeja = models.ForeignKey(
        Bandeja,
        on_delete=models.CASCADE,
        related_name='cajas',
        blank=True,
        null=True,
    )
    numero = models.CharField(max_length=50)
    posicion_caja_rack = models.CharField(max_length=50)

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['bandeja', 'numero'],
                name='unique_caja_por_bandeja'
            ),
            models.UniqueConstraint(
                fields=['bandeja', 'posicion_caja_rack'],
                name='unique_posicion_caja_por_bandeja'
            )
        ]

    def save(self, *args, **kwargs):
        """Sincroniza el rack desde la bandeja y mantiene el histórico al renombrar."""
        if self.bandeja_id:
            self.rack = self.bandeja.rack
        if self.pk:
            old = Caja.objects.get(pk=self.pk)
            if old.numero != self.numero:
                Localizacion.objects.filter(
                    congelador=self.rack.estante.congelador.congelador,
                    estante=self.rack.estante.numero,
                    rack=self.rack.numero,
                    caja=old.numero
                ).update(caja=self.numero)
        super().save(*args, **kwargs)


class Subposicion(models.Model):
    """Hueco individual donde puede almacenarse una muestra concreta."""

    caja = models.ForeignKey(Caja, on_delete=models.CASCADE, related_name='subposiciones')
    numero = models.CharField(max_length=50)
    fila = models.CharField(max_length=50, blank=True, null=True)
    columna = models.CharField(max_length=50, blank=True, null=True)
    vacia = models.BooleanField(default=True)
    muestra = models.OneToOneField(
        Muestra,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='subposicion'
    )

    class Meta:
        constraints = [
            models.UniqueConstraint(
                fields=['caja', 'fila', 'columna'],
                name='unique_subposicion_por_caja'
            )
        ]

    @property
    def etiqueta(self):
        """Etiqueta legible para mostrar coordenadas o numeración simple."""
        if self.fila and self.columna:
            return f"Fila {self.fila}, Columna {self.columna}"
        return self.numero

    def save(self, *args, **kwargs):
        """Genera el identificador compuesto y propaga renombrados al histórico."""
        if self.fila and self.columna:
            self.numero = f"{self.fila}-{self.columna}"
        if self.pk:
            old = Subposicion.objects.get(pk=self.pk)
            if old.numero != self.numero:
                Localizacion.objects.filter(
                    congelador=self.caja.rack.estante.congelador.congelador,
                    estante=self.caja.rack.estante.numero,
                    rack=self.caja.rack.numero,
                    caja=self.caja.numero,
                    subposicion=old.numero
                ).update(subposicion=self.numero)
        super().save(*args, **kwargs)
