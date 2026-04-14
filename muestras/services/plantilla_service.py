import os
from django.conf import settings


def get_global_files_path():
    return os.path.join(
        settings.BASE_DIR,
        "geslab_project",
        "globalstaticfiles",
        "files",
    )


def get_plantilla_path(nombre_fichero: str) -> str:
    return os.path.join(get_global_files_path(), nombre_fichero)


def get_listado_muestras_path() -> str:
    return get_plantilla_path("listado_muestras.xlsx")


def get_plantilla_muestras_path() -> str:
    return get_plantilla_path("plantilla_muestras.xlsx")


def get_plantilla_envios_path() -> str:
    return get_plantilla_path("plantilla_envios.xlsx")


def get_plantilla_estudios_path() -> str:
    return get_plantilla_path("plantilla_estudios.xlsx")


def get_plantilla_localizaciones_path() -> str:
    return get_plantilla_path("plantilla_localizaciones.xlsx")


def get_plantilla_localizaciones_macros_path() -> str:
    return get_plantilla_path("plantilla_localizaciones_macros.xlsm")


def get_plantilla_cambio_posicion_path() -> str:
    return get_plantilla_path("plantilla_cambio_posicion.xlsx")