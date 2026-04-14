from django import forms

from ..models import Estudio, Material, Muestra
from .common import no_semicolon


VOLUME_UNIT_CHOICES = [
    ("", "-- Seleccionar unidad --"),
    ("uL", "uL"),
    ("mL", "mL"),
    ("L", "L"),
]

CONCENTRATION_UNIT_CHOICES = [
    ("", "-- Seleccionar unidad --"),
    ("ng/uL", "ng/uL"),
    ("ng/mL", "ng/mL"),
    ("ug/mL", "ug/mL"),
    ("mg/mL", "mg/mL"),
    ("g/L", "g/L"),
]


class MuestraForm(forms.ModelForm):
    class Meta:
        model = Muestra
        fields = [
            "id_individuo",
            "nom_lab",
            "material",
            "volumen_actual",
            "unidad_volumen",
            "concentracion_actual",
            "unidad_concentracion",
            "masa_actual",
            "unidad_masa",
            "fecha_extraccion",
            "fecha_llegada",
            "observaciones",
            "estado_inicial",
            "centro_procedencia",
            "lugar_procedencia",
            "estado_actual",
            "estudio",
        ]
        widgets = {
            "fecha_extraccion": forms.DateInput(format="%Y-%m-%d", attrs={"type": "date"}),
            "fecha_llegada": forms.DateInput(format="%Y-%m-%d", attrs={"type": "date"}),
        }
        labels = {
            "nom_lab": "ID muestra",
            "material": "Material",
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        if "estado_actual" in self.fields:
            self.fields["estado_actual"].required = True

        if "material" in self.fields:
            queryset = Material.objects.filter(activo=True)
            if self.instance and self.instance.material_id:
                queryset = Material.objects.filter(pk=self.instance.material_id) | queryset
            self.fields["material"].queryset = queryset.distinct().order_by("nombre")
            self.fields["material"].required = True

        text_fields = [
            "id_individuo",
            "nom_lab",
            "centro_procedencia",
            "lugar_procedencia",
            "estado_actual",
            "observaciones",
            "estado_inicial",
        ]
        for field_name in text_fields:
            if field_name in self.fields:
                if not isinstance(self.fields[field_name].validators, list):
                    self.fields[field_name].validators = list(self.fields[field_name].validators)
                self.fields[field_name].validators.append(no_semicolon)


class IndividuoEstudioForm(forms.Form):
    id_individuo = forms.CharField(max_length=20, label="ID individuo")
    estudio = forms.ModelChoiceField(
        queryset=Estudio.objects.all().order_by("nombre_estudio"),
        required=False,
        empty_label="-- Sin estudio asociado --",
        label="Estudio asociado",
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["id_individuo"].validators.append(no_semicolon)


class DatosMuestrasComunesForm(forms.Form):
    fecha_extraccion = forms.DateField(
        required=False,
        label="Fecha de extracción",
        widget=forms.DateInput(attrs={"type": "date"}),
    )
    fecha_llegada = forms.DateField(
        required=False,
        label="Fecha de llegada",
        widget=forms.DateInput(attrs={"type": "date"}),
    )
    centro_procedencia = forms.CharField(max_length=100, required=False, label="Centro de procedencia")
    observaciones = forms.CharField(
        required=False,
        widget=forms.Textarea(attrs={"rows": 2}),
        label="Observaciones generales",
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for field_name in ("centro_procedencia", "observaciones"):
            self.fields[field_name].validators.append(no_semicolon)


class MaterialForm(forms.ModelForm):
    class Meta:
        model = Material
        fields = [
            "nombre",
            "usa_volumen",
            "unidades_volumen",
            "usa_concentracion",
            "unidades_concentracion",
            "usa_unidades_bloque",
            "unidades_recuento",
        ]
        labels = {
            "usa_volumen": "Usa volumen actual",
            "unidades_volumen": "Unidades de volumen permitidas",
            "usa_concentracion": "Usa concentración actual",
            "unidades_concentracion": "Unidades de concentración permitidas",
            "usa_unidades_bloque": "Usa unidades",
            "unidades_recuento": "Unidades permitidas",
        }
        widgets = {
            "unidades_volumen": forms.TextInput(attrs={"placeholder": "Ej: uL, mL"}),
            "unidades_concentracion": forms.TextInput(attrs={"placeholder": "Ej: ng/uL, mg/mL"}),
            "unidades_recuento": forms.TextInput(attrs={"placeholder": "Ej: bloque, lámina"}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields["nombre"].validators.append(no_semicolon)
        self.fields["unidades_volumen"].help_text = "Separe las opciones por comas."
        self.fields["unidades_concentracion"].help_text = "Separe las opciones por comas."
        self.fields["unidades_recuento"].help_text = "Separe las opciones por comas."

    def clean(self):
        cleaned_data = super().clean()
        usa_volumen = cleaned_data.get("usa_volumen")
        usa_concentracion = cleaned_data.get("usa_concentracion")
        usa_unidades = cleaned_data.get("usa_unidades_bloque")

        unidades_volumen = (cleaned_data.get("unidades_volumen") or "").strip()
        unidades_concentracion = (cleaned_data.get("unidades_concentracion") or "").strip()
        unidades_recuento = (cleaned_data.get("unidades_recuento") or "").strip()

        if usa_volumen and not unidades_volumen:
            self.add_error("unidades_volumen", "Indica qué unidades de volumen puede utilizar este material.")
        if not usa_volumen:
            cleaned_data["unidades_volumen"] = ""

        if usa_concentracion and not unidades_concentracion:
            self.add_error("unidades_concentracion", "Indica qué unidades de concentración puede utilizar este material.")
        if not usa_concentracion:
            cleaned_data["unidades_concentracion"] = ""

        if usa_unidades and not unidades_recuento:
            self.add_error("unidades_recuento", "Indica qué unidades puede utilizar este material.")
        if not usa_unidades:
            cleaned_data["unidades_recuento"] = ""

        return cleaned_data


class AliquotaForm(forms.Form):
    nom_lab = forms.CharField(max_length=100, label="ID muestra")
    material = forms.ModelChoiceField(
        queryset=Material.objects.filter(activo=True).order_by("nombre"),
        label="Material",
        empty_label="-- Seleccionar material --",
    )
    volumen_actual = forms.FloatField(required=False, label="Volumen actual")
    unidad_volumen = forms.ChoiceField(required=False, choices=VOLUME_UNIT_CHOICES, label="Unidad de volumen")
    concentracion_actual = forms.FloatField(required=False, label="Concentración actual")
    unidad_concentracion = forms.ChoiceField(
        required=False,
        choices=CONCENTRATION_UNIT_CHOICES,
        label="Unidad de concentración",
    )
    masa_actual = forms.FloatField(required=False, label="Unidades de bloque")
    unidad_masa = forms.ChoiceField(
        required=False,
        choices=[("", "-- Seleccionar unidad --")],
        label="Unidad de cantidad",
    )
    fecha_extraccion = forms.DateField(
        required=False,
        label="Fecha de extracción",
        widget=forms.DateInput(attrs={"type": "date"}),
    )
    fecha_llegada = forms.DateField(
        required=False,
        label="Fecha de llegada",
        widget=forms.DateInput(attrs={"type": "date"}),
    )
    observaciones = forms.CharField(required=False, widget=forms.Textarea, label="Observaciones")
    centro_procedencia = forms.CharField(max_length=100, required=False, label="Centro de procedencia")
    estado_actual = forms.ChoiceField(
        choices=Muestra._meta.get_field("estado_actual").choices,
        initial="DISP",
        required=True,
        label="Estado actual",
    )
    subposicion_id = forms.CharField(required=False, widget=forms.HiddenInput())

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        for field_name in ("nom_lab", "centro_procedencia", "observaciones", "estado_actual"):
            self.fields[field_name].validators.append(no_semicolon)

        material = None
        prefix = self.prefix or ""
        material_key = f"{prefix}-material" if prefix else "material"

        if self.is_bound:
            material_id = self.data.get(material_key)
            if material_id:
                material = Material.objects.filter(pk=material_id).first()
        else:
            initial_material = self.initial.get("material") if hasattr(self, "initial") else None
            if isinstance(initial_material, Material):
                material = initial_material
            elif initial_material:
                material = Material.objects.filter(pk=initial_material).first()

        if material:
            self.fields["unidad_volumen"].choices = [
                ("", "-- Seleccionar unidad --"),
                *[(value, value) for value in material.unidades_volumen_list()],
            ]
            self.fields["unidad_concentracion"].choices = [
                ("", "-- Seleccionar unidad --"),
                *[(value, value) for value in material.unidades_concentracion_list()],
            ]
            self.fields["unidad_masa"].choices = [
                ("", "-- Seleccionar unidad --"),
                *[(value, value) for value in material.unidades_recuento_list()],
            ]

    def clean_nom_lab(self):
        nom_lab = (self.cleaned_data.get("nom_lab") or "").strip()
        if not nom_lab:
            raise forms.ValidationError("Este campo es obligatorio.")
        if Muestra.objects.filter(nom_lab__iexact=nom_lab).exists():
            raise forms.ValidationError("Ya existe una muestra con ese ID.")
        return nom_lab

    def clean(self):
        cleaned_data = super().clean()
        material = cleaned_data.get("material")
        if not material:
            return cleaned_data

        unidad_volumen = cleaned_data.get("unidad_volumen")
        unidad_concentracion = cleaned_data.get("unidad_concentracion")
        unidad_masa = cleaned_data.get("unidad_masa")

        if material.usa_volumen:
            unidades_validas = material.unidades_volumen_list()
            if unidad_volumen and unidad_volumen not in unidades_validas:
                self.add_error("unidad_volumen", "Selecciona una unidad de volumen válida.")
        else:
            cleaned_data["volumen_actual"] = None
            cleaned_data["unidad_volumen"] = ""

        if material.usa_concentracion:
            if concentracion_actual in (None, ""):
                self.add_error("concentracion_actual", "Indica la concentración actual.")
            if not unidad_concentracion:
                self.add_error("unidad_concentracion", "Selecciona la unidad de concentración.")
        else:
            cleaned_data["concentracion_actual"] = None
            cleaned_data["unidad_concentracion"] = ""

        if material.usa_unidades_bloque:
            unidades_validas = material.unidades_recuento_list()
            if unidad_masa and unidad_masa not in unidades_validas:
                self.add_error("unidad_masa", "Selecciona una unidad de cantidad válida.")
        else:
            cleaned_data["masa_actual"] = None
            cleaned_data["unidad_masa"] = ""

        return cleaned_data

    def clean(self):
        cleaned_data = super().clean()
        material = cleaned_data.get("material")
        if not material:
            return cleaned_data

        unidad_volumen = cleaned_data.get("unidad_volumen")
        unidad_concentracion = cleaned_data.get("unidad_concentracion")
        unidad_masa = cleaned_data.get("unidad_masa")

        if material.usa_volumen:
            unidades_validas = material.unidades_volumen_list()
            if unidad_volumen and unidad_volumen not in unidades_validas:
                self.add_error("unidad_volumen", "Selecciona una unidad de volumen válida.")
        else:
            cleaned_data["volumen_actual"] = None
            cleaned_data["unidad_volumen"] = ""

        if material.usa_concentracion:
            unidades_validas = material.unidades_concentracion_list()
            if unidad_concentracion and unidad_concentracion not in unidades_validas:
                self.add_error("unidad_concentracion", "Selecciona una unidad de concentración válida.")
        else:
            cleaned_data["concentracion_actual"] = None
            cleaned_data["unidad_concentracion"] = ""

        if material.usa_unidades_bloque:
            unidades_validas = material.unidades_recuento_list()
            if unidad_masa and unidad_masa not in unidades_validas:
                self.add_error("unidad_masa", "Selecciona una unidad de cantidad válida.")
        else:
            cleaned_data["masa_actual"] = None
            cleaned_data["unidad_masa"] = ""

        return cleaned_data


class UploadExcel(forms.Form):
    excel_file = forms.FileField(required=True)

    def clean_excel_file(self):
        excel_file = self.cleaned_data["excel_file"]
        file_name = (excel_file.name or "").lower()
        if not file_name.endswith(".xlsx"):
            raise forms.ValidationError("Solo se permiten archivos Excel en formato .xlsx.")
        return excel_file


class DestruirMuestrasForm(forms.Form):
    motivo_destruccion = forms.CharField(max_length=250, required=True, label="Motivo de destrucción")
    metodo_destruccion = forms.CharField(max_length=250, required=True, label="Método de destrucción")
    lugar_destruccion = forms.CharField(max_length=250, required=True, label="Lugar de destrucción")
    responsable_autoriza = forms.CharField(max_length=250, required=True, label="Responsable que autoriza")
    tecnico_realiza = forms.CharField(max_length=250, required=True, label="Técnico que realiza")
    fecha_destruccion = forms.DateField(
        required=True,
        label="Fecha de destrucción",
        widget=forms.DateInput(attrs={"type": "date"})
    )
    observaciones_destruccion = forms.CharField(
        widget=forms.Textarea,
        required=False,
        label="Otras observaciones"
    )
    confirmar_destruccion = forms.BooleanField(
        required=True,
        label="Confirmo que las muestras seleccionadas van a ser destruidas"
    )
