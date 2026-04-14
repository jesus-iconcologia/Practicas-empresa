from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("muestras", "0017_bandejas_filas_columnas"),
    ]

    operations = [
        migrations.AddField(
            model_name="congelador",
            name="tipo_estructura",
            field=models.CharField(
                blank=True,
                choices=[
                    ("congelador_80", "Congelador -80"),
                    ("congelador_20", "Congelador -20"),
                    ("nevera", "Nevera"),
                ],
                max_length=30,
                null=True,
            ),
        ),
    ]
