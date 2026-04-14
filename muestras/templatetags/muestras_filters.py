from django import template
import os

register = template.Library()

ESTADO_LABELS = {
    "DISP": "Disponible",
    "PENV": "Parcialmente enviada",
    "ENV": "Enviada",
    "DEST": "Destruida",
}

ESTADO_BADGE_CLASSES = {
    "DISP": "disponible",
    "PENV": "parcial",
    "ENV": "enviada",
    "DEST": "destruida",
}


@register.filter
def startswith(text, starts):
    if isinstance(text, str):
        return text.startswith(starts)
    return False


@register.filter
def get_item(dictionary, key):
    try:
        return dictionary.get(key)
    except Exception:
        return ""


@register.filter
def estado_label(value):
    return ESTADO_LABELS.get(value, value or "")


@register.filter
def estado_badge_class(value):
    return ESTADO_BADGE_CLASSES.get(value, "")


@register.filter
def basename(value):
    if not value:
        return ""
    return os.path.basename(str(value))
