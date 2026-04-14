from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("muestras", "0015_registro_destruido_fields"),
    ]

    operations = [
        migrations.AlterField(
            model_name="muestra",
            name="nom_lab",
            field=models.CharField(
                help_text="Nombre unico de la muestra asignado por el laboratorio",
                max_length=100,
                unique=True,
            ),
        ),
    ]
