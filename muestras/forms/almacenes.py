from django import forms
from django.forms import inlineformset_factory

from ..models import Congelador, Localizacion, ReparacionRecurso
from .common import no_semicolon


class RecursoClearableFileInput(forms.ClearableFileInput):
    """Textos del widget de imagen adaptados al lenguaje de recursos."""

    initial_text = "Fotografía actual"
    input_text = "Nueva fotografía"
    clear_checkbox_label = "Sin fotografía"


class ArchivarMuestraForm(forms.ModelForm):
    class Meta:
        model = Localizacion
        exclude = ("muestra",)


class CongeladorForm(forms.ModelForm):
    """Formulario base para crear y editar los datos descriptivos de un recurso."""

    class Meta:
        model = Congelador
        fields = (
            "congelador",
            "modelo",
            "ano_compra",
            "temperatura_objetivo",
            "localizacion_edificio",
            "fotografia",
            "tipo_estructura",
        )
        labels = {
            "congelador": "Nombre del recurso",
            "modelo": "Modelo",
            "ano_compra": "Año de compra",
            "temperatura_objetivo": "Temperatura en ºC",
            "localizacion_edificio": "Localización",
            "fotografia": "Fotografía del recurso",
            "tipo_estructura": "Tipo de recurso",
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        fotografia = self.fields.get("fotografia")
        if fotografia:
            fotografia.widget = RecursoClearableFileInput()

        for field in self.fields.values():
            if isinstance(field, forms.CharField):
                if not isinstance(field.validators, list):
                    field.validators = list(field.validators)
                field.validators.append(no_semicolon)

        if self.instance and self.instance.pk and "tipo_estructura" in self.fields:
            self.fields["tipo_estructura"].disabled = True
            self.fields["tipo_estructura"].help_text = (
                "La estructura del recurso no se puede modificar una vez creado."
            )


class CrearRecursoCongeladorForm(CongeladorForm):
    """Formulario de alta de recursos con estructura obligatoria."""

    class Meta(CongeladorForm.Meta):
        fields = (
            "congelador",
            "modelo",
            "ano_compra",
            "temperatura_objetivo",
            "localizacion_edificio",
            "fotografia",
            "tipo_estructura",
        )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if "tipo_estructura" in self.fields:
            self.fields["tipo_estructura"].required = True


class CrearPosicionManualForm(forms.Form):
    """Formulario manual para crear posiciones según la estructura del recurso."""

    estante = forms.IntegerField(label="Estante", min_value=1)
    posicion_rack_estante = forms.IntegerField(label="Posición del cajón en el estante", min_value=1)
    rack = forms.CharField(label="Cajón", max_length=50)
    posicion_bandeja_rack = forms.IntegerField(label="Posición de la bandeja en el rack", min_value=1, required=False)
    bandeja = forms.CharField(label="Bandeja", max_length=50, required=False)
    posicion_caja_rack = forms.IntegerField(label="Posición de la caja en la bandeja", min_value=1, required=False)
    caja = forms.CharField(label="Caja", max_length=50)
    fila = forms.CharField(label="Fila", max_length=50)
    columna = forms.CharField(label="Columna", max_length=50)

    def __init__(self, *args, **kwargs):
        self.recurso = kwargs.pop("recurso")
        super().__init__(*args, **kwargs)

        if self.recurso.tipo_estructura == Congelador.ESTRUCTURA_CONGELADOR_80:
            self.fields["rack"].label = "Rack"
        else:
            self.fields.pop("posicion_bandeja_rack")
            self.fields.pop("bandeja")
            self.fields.pop("posicion_caja_rack")

        for field in self.fields.values():
            if isinstance(field, forms.CharField):
                if not isinstance(field.validators, list):
                    field.validators = list(field.validators)
                field.validators.append(no_semicolon)


class ReparacionRecursoForm(forms.ModelForm):
    """Formulario compacto para registrar el historial de reparaciones."""

    class Meta:
        model = ReparacionRecurso
        fields = ("fecha_reparacion", "descripcion", "coste")
        labels = {
            "fecha_reparacion": "Fecha de reparación",
            "descripcion": "Descripción de la reparación",
            "coste": "Coste",
        }
        widgets = {
            "fecha_reparacion": forms.DateInput(attrs={"type": "date"}),
            "descripcion": forms.Textarea(attrs={"rows": 3}),
        }

    def __init__(self, *args, **kwargs):
        self.recurso = kwargs.pop("recurso", None)
        super().__init__(*args, **kwargs)
        descripcion = self.fields.get("descripcion")
        if descripcion:
            if not isinstance(descripcion.validators, list):
                descripcion.validators = list(descripcion.validators)
            descripcion.validators.append(no_semicolon)

    def clean_fecha_reparacion(self):
        fecha_reparacion = self.cleaned_data.get("fecha_reparacion")
        recurso = self.recurso or getattr(self.instance, "recurso", None)
        if recurso and recurso.ano_compra and fecha_reparacion and fecha_reparacion.year < recurso.ano_compra:
            raise forms.ValidationError(
                "La fecha de reparación no puede ser anterior al año de compra del recurso."
            )
        return fecha_reparacion


ReparacionRecursoFormSet = inlineformset_factory(
    Congelador,
    ReparacionRecurso,
    form=ReparacionRecursoForm,
    extra=1,
    can_delete=True,
)
