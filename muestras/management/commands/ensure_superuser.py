from django.core.management.base import BaseCommand

from muestras.services.auth_service import ensure_superuser_from_env


class Command(BaseCommand):
    help = "Crea o actualiza el superusuario configurado por variables de entorno."

    def handle(self, *args, **options):
        user, created = ensure_superuser_from_env()

        if user is None:
            self.stdout.write(
                self.style.WARNING(
                    "DJANGO_SUPERUSER_USERNAME, DJANGO_SUPERUSER_PASSWORD o DJANGO_SUPERUSER_EMAIL no estan definidos."
                )
            )
            return

        if created:
            self.stdout.write(self.style.SUCCESS("Superuser created."))
        else:
            self.stdout.write(self.style.SUCCESS("Superuser updated."))
