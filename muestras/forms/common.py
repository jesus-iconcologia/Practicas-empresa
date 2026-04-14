from django.core.exceptions import ValidationError


def no_semicolon(value):
    if value and ";" in value:
        raise ValidationError("Aquest camp no pot contenir el caràcter ';'")