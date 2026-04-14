from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ("muestras", "0018_congelador_tipo_estructura"),
    ]

    operations = [
        migrations.AddField(
            model_name="congelador",
            name="ano_compra",
            field=models.PositiveIntegerField(blank=True, null=True),
        ),
        migrations.AddField(
            model_name="congelador",
            name="temperatura_objetivo",
            field=models.CharField(blank=True, max_length=50, null=True),
        ),
        migrations.CreateModel(
            name="ReparacionRecurso",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("fecha_reparacion", models.DateField()),
                ("descripcion", models.TextField()),
                ("coste", models.DecimalField(blank=True, decimal_places=2, max_digits=10, null=True)),
                ("recurso", models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name="reparaciones", to="muestras.congelador")),
            ],
            options={
                "ordering": ["-fecha_reparacion", "-id"],
            },
        ),
    ]
