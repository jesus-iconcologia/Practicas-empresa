from django.contrib.auth.models import Group, Permission
from django.core.management.base import BaseCommand


class Command(BaseCommand):
    help = "Crea los grupos y permisos base de GesLab de forma idempotente."

    GROUPS = {
        "Tecnicos de laboratorio": {
            "muestras.can_view_muestras_web",
            "muestras.can_add_muestras_web",
            "muestras.can_change_muestras_web",
            "muestras.can_delete_muestras_web",
            "muestras.can_view_localizaciones_web",
            "muestras.can_add_localizaciones_web",
            "muestras.can_change_localizaciones_web",
            "muestras.can_delete_localizaciones_web",
            "muestras.can_view_estudios_web",
            "muestras.can_add_estudios_web",
            "muestras.can_change_estudios_web",
            "muestras.can_delete_estudios_web",
        },
        "Investigadores": {
            "muestras.can_view_estudios_web",
            "muestras.can_change_estudios_web",
        },
    }

    def handle(self, *args, **options):
        created_groups = 0

        for group_name, permission_names in self.GROUPS.items():
            group, created = Group.objects.get_or_create(name=group_name)
            if created:
                created_groups += 1

            permissions = []
            missing = []
            for full_name in permission_names:
                app_label, codename = full_name.split(".", 1)
                permission = Permission.objects.filter(
                    content_type__app_label=app_label,
                    codename=codename,
                ).first()
                if permission is None:
                    missing.append(full_name)
                    continue
                permissions.append(permission)

            if missing:
                self.stdout.write(
                    self.style.WARNING(
                        f"Permisos no encontrados para '{group_name}': {', '.join(sorted(missing))}"
                    )
                )

            group.permissions.set(permissions)

        self.stdout.write(
            self.style.SUCCESS(
                f"Bootstrap completado. Grupos creados en esta ejecucion: {created_groups}."
            )
        )
