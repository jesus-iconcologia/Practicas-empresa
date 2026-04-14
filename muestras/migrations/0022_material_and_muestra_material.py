from django.db import migrations, models
import django.db.models.deletion


def crear_materiales_base(apps, schema_editor):
    Material = apps.get_model("muestras", "Material")
    Muestra = apps.get_model("muestras", "Muestra")

    materiales = {
        "Sangre": {
            "usa_volumen": True,
            "usa_concentracion": False,
            "usa_unidades_bloque": False,
        },
        "Orina": {
            "usa_volumen": True,
            "usa_concentracion": True,
            "usa_unidades_bloque": False,
        },
        "Parafina": {
            "usa_volumen": False,
            "usa_concentracion": False,
            "usa_unidades_bloque": True,
        },
    }

    creados = {}
    for nombre, config in materiales.items():
        material, _ = Material.objects.get_or_create(nombre=nombre, defaults=config)
        creados[nombre.lower()] = material

    for muestra in Muestra.objects.exclude(id_material__isnull=True).exclude(id_material=""):
        material = creados.get(str(muestra.id_material).strip().lower())
        if material:
            muestra.material = material
            muestra.save(update_fields=["material"])


class Migration(migrations.Migration):

    dependencies = [
        ("muestras", "0021_remove_congelador_es_borrador"),
    ]

    operations = [
        migrations.CreateModel(
            name="Material",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("nombre", models.CharField(max_length=100, unique=True)),
                ("activo", models.BooleanField(default=True)),
                ("usa_volumen", models.BooleanField(default=False)),
                ("usa_concentracion", models.BooleanField(default=False)),
                ("usa_unidades_bloque", models.BooleanField(default=False)),
                ("notas", models.TextField(blank=True, null=True)),
            ],
            options={
                "ordering": ["nombre"],
                "verbose_name": "Material",
                "verbose_name_plural": "Materiales",
            },
        ),
        migrations.AddField(
            model_name="muestra",
            name="material",
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name="muestras", to="muestras.material"),
        ),
        migrations.RunPython(crear_materiales_base, migrations.RunPython.noop),
    ]
