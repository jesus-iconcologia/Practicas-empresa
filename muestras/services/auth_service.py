import os

from django.contrib.auth import get_user_model


def ensure_superuser_from_env():
    username = os.environ.get("DJANGO_SUPERUSER_USERNAME")
    password = os.environ.get("DJANGO_SUPERUSER_PASSWORD")
    email = os.environ.get("DJANGO_SUPERUSER_EMAIL")

    if not username or not password or not email:
        return None, False

    User = get_user_model()
    user, created = User.objects.get_or_create(
        username=username,
        defaults={"email": email},
    )

    user.email = email
    user.set_password(password)
    user.is_staff = True
    user.is_superuser = True
    user.is_active = True
    user.save()

    return user, created
