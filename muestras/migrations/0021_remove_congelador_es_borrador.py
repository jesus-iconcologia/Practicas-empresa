from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ("muestras", "0020_congelador_es_borrador"),
    ]

    operations = [
        migrations.RemoveField(
            model_name="congelador",
            name="es_borrador",
        ),
    ]
