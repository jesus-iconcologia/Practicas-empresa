from django import forms
from django.core.exceptions import ValidationError

from ..models import Estudio, Documento
from .common import no_semicolon


class EstudioForm(forms.ModelForm):
    """Formulario principal del módulo de estudios con validación centralizada."""

    class Meta:
        model = Estudio
        fields = [
            "referencia_estudio",
            "nombre_estudio",
            "descripcion_estudio",
            "fecha_inicio_estudio",
            "fecha_fin_estudio",
            "investigador_principal",
        ]
        widgets = {
            "fecha_inicio_estudio": forms.DateInput(format="%Y-%m-%d", attrs={"type": "date"}),
            "fecha_fin_estudio": forms.DateInput(format="%Y-%m-%d", attrs={"type": "date"}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        text_fields = [
            "referencia_estudio",
            "nombre_estudio",
            "descripcion_estudio",
            "investigador_principal",
        ]
        for field_name in text_fields:
            if field_name in self.fields:
                if not isinstance(self.fields[field_name].validators, list):
                    self.fields[field_name].validators = list(self.fields[field_name].validators)
                self.fields[field_name].validators.append(no_semicolon)

    def _clean_text_field(self, field_name, *, required=False):
        value = self.cleaned_data.get(field_name)
        if value is None:
            return None

        normalized = value.strip()
        if required and not normalized:
            raise ValidationError("Este campo es obligatorio.")
        return normalized or None

    def clean_nombre_estudio(self):
        nombre = self._clean_text_field("nombre_estudio", required=True)
        duplicate_qs = Estudio.objects.filter(nombre_estudio__iexact=nombre)
        if self.instance.pk:
            duplicate_qs = duplicate_qs.exclude(pk=self.instance.pk)
        if duplicate_qs.exists():
            raise ValidationError("Ya existe otro estudio con ese nombre.")
        return nombre

    def clean_referencia_estudio(self):
        referencia = self._clean_text_field("referencia_estudio")
        if not referencia:
            return None

        duplicate_qs = Estudio.objects.filter(referencia_estudio__iexact=referencia)
        if self.instance.pk:
            duplicate_qs = duplicate_qs.exclude(pk=self.instance.pk)
        if duplicate_qs.exists():
            raise ValidationError("Ya existe otro estudio con esa referencia.")
        return referencia

    def clean_descripcion_estudio(self):
        return self._clean_text_field("descripcion_estudio")

    def clean_investigador_principal(self):
        return self._clean_text_field("investigador_principal")

    def clean(self):
        cleaned_data = super().clean()
        fecha_inicio = cleaned_data.get("fecha_inicio_estudio")
        fecha_fin = cleaned_data.get("fecha_fin_estudio")

        if fecha_inicio and fecha_fin and fecha_fin < fecha_inicio:
            self.add_error(
                "fecha_fin_estudio",
                "La fecha de fin debe ser igual o posterior a la fecha de inicio.",
            )

        return cleaned_data


class DocumentoForm(forms.ModelForm):
    class Meta:
        model = Documento
        fields = ["archivo", "categoria", "descripcion"]

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        text_fields = ["categoria", "descripcion"]
        for field_name in text_fields:
            if field_name in self.fields:
                if not isinstance(self.fields[field_name].validators, list):
                    self.fields[field_name].validators = list(self.fields[field_name].validators)
                self.fields[field_name].validators.append(no_semicolon)
