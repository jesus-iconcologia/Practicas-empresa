#!/bin/sh

DB_HOST=$MYSQL_HOST
DB_PORT=$MYSQL_PORT

echo "Waiting for database host ($DB_HOST:$DB_PORT) to be available..."

while ! nc -z "$DB_HOST" "$DB_PORT"; do
  sleep 0.5
done

echo "Database port is open. Proceeding with Django setup..."

echo "Collecting static files..."
python manage.py collectstatic --noinput

echo "Applying database migrations..."
python manage.py migrate

echo "Bootstrapping initial groups and permissions..."
python manage.py bootstrap_initial_data

if [ -n "$DJANGO_SUPERUSER_USERNAME" ] && [ -n "$DJANGO_SUPERUSER_PASSWORD" ] && [ -n "$DJANGO_SUPERUSER_EMAIL" ]; then
  echo "Ensuring superuser exists..."
  python manage.py ensure_superuser
fi

echo "Starting Django server..."
exec "$@"
