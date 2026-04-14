from django.db import migrations, models

# Añade campos de detalle al registro de destrucción de muestras


class Migration(migrations.Migration):

    dependencies = [
        ('muestras', '0014_alter_muestra_fecha_extraccion_and_more'),
    ]

    operations = [
        migrations.AddField(
            model_name='registro_destruido',
            name='motivo',
            field=models.CharField(blank=True, max_length=250, null=True),
        ),
        migrations.AddField(
            model_name='registro_destruido',
            name='metodo',
            field=models.CharField(blank=True, max_length=250, null=True),
        ),
        migrations.AddField(
            model_name='registro_destruido',
            name='lugar',
            field=models.CharField(blank=True, max_length=250, null=True),
        ),
        migrations.AddField(
            model_name='registro_destruido',
            name='responsable',
            field=models.CharField(blank=True, max_length=250, null=True),
        ),
        migrations.AddField(
            model_name='registro_destruido',
            name='tecnico',
            field=models.CharField(blank=True, max_length=250, null=True),
        ),
        migrations.AddField(
            model_name='registro_destruido',
            name='observaciones',
            field=models.TextField(blank=True, null=True),
        ),
    ]
