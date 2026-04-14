from django.db import migrations, models
import django.db.models.deletion


def _split_subposicion(valor):
    if valor is None:
        return "1", "1"
    texto = str(valor).strip()
    if not texto:
        return "1", "1"
    for separador in ("-", ",", "/", "x", "X"):
        if separador in texto:
            partes = [p.strip() for p in texto.split(separador) if p.strip()]
            if len(partes) >= 2:
                return partes[0], partes[1]
    return "1", texto


def poblar_bandejas_y_subposiciones(apps, schema_editor):
    Rack = apps.get_model("muestras", "Rack")
    Bandeja = apps.get_model("muestras", "Bandeja")
    Caja = apps.get_model("muestras", "Caja")
    Subposicion = apps.get_model("muestras", "Subposicion")
    Localizacion = apps.get_model("muestras", "Localizacion")

    bandejas_por_rack = {}
    for rack in Rack.objects.all():
        bandeja, _ = Bandeja.objects.get_or_create(
            rack_id=rack.id,
            numero="1",
            defaults={"posicion_bandeja_rack": "1"},
        )
        bandejas_por_rack[rack.id] = bandeja.id

    for caja in Caja.objects.all():
        bandeja_id = bandejas_por_rack.get(caja.rack_id)
        if bandeja_id:
            Caja.objects.filter(pk=caja.pk).update(bandeja_id=bandeja_id)

    for sub in Subposicion.objects.all():
        fila, columna = _split_subposicion(sub.numero)
        Subposicion.objects.filter(pk=sub.pk).update(
            fila=fila,
            columna=columna,
            numero=f"{fila}-{columna}",
        )

    for loc in Localizacion.objects.all():
        fila, columna = _split_subposicion(loc.subposicion)
        update_kwargs = {
            "fila": fila,
            "columna": columna,
            "subposicion": f"{fila}-{columna}",
        }
        if loc.rack:
            update_kwargs["bandeja"] = "1"
        Localizacion.objects.filter(pk=loc.pk).update(**update_kwargs)


class Migration(migrations.Migration):

    dependencies = [
        ("muestras", "0016_alter_muestra_nom_lab"),
    ]

    operations = [
        migrations.CreateModel(
            name="Bandeja",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("numero", models.CharField(max_length=50)),
                ("posicion_bandeja_rack", models.CharField(max_length=50)),
                ("rack", models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name="bandejas", to="muestras.rack")),
            ],
        ),
        migrations.AddField(
            model_name="caja",
            name="bandeja",
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name="cajas", to="muestras.bandeja"),
        ),
        migrations.AddField(
            model_name="localizacion",
            name="bandeja",
            field=models.CharField(blank=True, max_length=50, null=True),
        ),
        migrations.AddField(
            model_name="localizacion",
            name="columna",
            field=models.CharField(blank=True, max_length=50, null=True),
        ),
        migrations.AddField(
            model_name="localizacion",
            name="fila",
            field=models.CharField(blank=True, max_length=50, null=True),
        ),
        migrations.AddField(
            model_name="subposicion",
            name="columna",
            field=models.CharField(blank=True, max_length=50, null=True),
        ),
        migrations.AddField(
            model_name="subposicion",
            name="fila",
            field=models.CharField(blank=True, max_length=50, null=True),
        ),
        migrations.RunPython(poblar_bandejas_y_subposiciones, migrations.RunPython.noop),
        migrations.RemoveConstraint(
            model_name="caja",
            name="unique_caja_por_rack",
        ),
        migrations.RemoveConstraint(
            model_name="caja",
            name="unique_posicion_caja_por_rack",
        ),
        migrations.RemoveConstraint(
            model_name="subposicion",
            name="unique_subposicion_por_caja",
        ),
        migrations.AddConstraint(
            model_name="bandeja",
            constraint=models.UniqueConstraint(fields=("rack", "numero"), name="unique_bandeja_por_rack"),
        ),
        migrations.AddConstraint(
            model_name="bandeja",
            constraint=models.UniqueConstraint(fields=("rack", "posicion_bandeja_rack"), name="unique_posicion_bandeja_por_rack"),
        ),
        migrations.AddConstraint(
            model_name="caja",
            constraint=models.UniqueConstraint(fields=("bandeja", "numero"), name="unique_caja_por_bandeja"),
        ),
        migrations.AddConstraint(
            model_name="caja",
            constraint=models.UniqueConstraint(fields=("bandeja", "posicion_caja_rack"), name="unique_posicion_caja_por_bandeja"),
        ),
        migrations.AddConstraint(
            model_name="subposicion",
            constraint=models.UniqueConstraint(fields=("caja", "fila", "columna"), name="unique_subposicion_por_caja"),
        ),
    ]
