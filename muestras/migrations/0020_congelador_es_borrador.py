from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("muestras", "0019_recurso_campos_y_reparaciones"),
    ]

    operations = [
        migrations.AddField(
            model_name="congelador",
            name="es_borrador",
            field=models.BooleanField(default=False),
        ),
    ]
