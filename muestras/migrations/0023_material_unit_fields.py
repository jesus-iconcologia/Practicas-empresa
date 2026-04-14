from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("muestras", "0022_material_and_muestra_material"),
    ]

    operations = [
        migrations.AddField(
            model_name="material",
            name="unidades_concentracion",
            field=models.CharField(blank=True, max_length=250, null=True),
        ),
        migrations.AddField(
            model_name="material",
            name="unidades_recuento",
            field=models.CharField(blank=True, max_length=250, null=True),
        ),
        migrations.AddField(
            model_name="material",
            name="unidades_volumen",
            field=models.CharField(blank=True, max_length=250, null=True),
        ),
    ]
