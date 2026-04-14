# GesLab

Aplicacion web de gestion de muestras desarrollada con Django y MySQL.

## Documentacion incluida

- `DOCUMENTACION_FUNCIONAL_GESLAB.md`: explicacion del sistema en lenguaje no tecnico.
- `REVISION_TECNICA_GESLAB.md`: revision del codigo, riesgos detectados y recomendaciones.

## Ejecucion en local con Docker

```bash
docker compose up --build
```

## Configuracion local recomendada

- Usa `.env` solo para tu entorno local y toma `.env.example` como plantilla para nuevas instalaciones.
- La configuracion de desarrollo queda limitada a `127.0.0.1`, de modo que Django y MySQL no se publiquen hacia la red local por defecto.
- `ALLOWED_HOSTS` ya no usa `*`; en local se limita a `127.0.0.1` y `localhost`.
- La variable recomendada para desarrollo es `DJANGO_DEBUG`, para evitar colisiones con variables globales del sistema anfitrion.
- `docker-compose.yml` solo inyecta en cada contenedor las variables que realmente necesita.
