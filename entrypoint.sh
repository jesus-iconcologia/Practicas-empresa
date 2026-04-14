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
  python manage.py shell -c "
from django.contrib.auth import get_user_model
import os

User = get_user_model()
username = os.environ['DJANGO_SUPERUSER_USERNAME']
email = os.environ['DJANGO_SUPERUSER_EMAIL']
password = os.environ['DJANGO_SUPERUSER_PASSWORD']

user, created = User.objects.get_or_create(
    username=username,
    defaults={"email": email},
)
user.email = email
user.set_password(password)
if not user.is_staff or not user.is_superuser:
    user.is_staff = True
    user.is_superuser = True
user.save()
if created:
    print('Superuser created.')
else:
    print('Superuser updated.')
"
fi

echo "Starting Django server..."
exec "$@"
