from django import forms

from ..models import agenda_envio
from .common import no_semicolon


class CentroForm(forms.ModelForm):
    class Meta:
        model = agenda_envio
        fields = "__all__"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        for field_name, field in self.fields.items():
            if isinstance(field, forms.CharField):
                if not isinstance(field.validators, list):
                    field.validators = list(field.validators)
                field.validators.append(no_semicolon)