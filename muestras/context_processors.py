from .ui_symbols import UI_SYMBOLS


def ui_symbols(request):
    context = dict(UI_SYMBOLS)
    context["UI_SYMBOLS"] = UI_SYMBOLS
    context["ICONO_DESPLEGABLE_CSS"] = f'"{UI_SYMBOLS["ICONO_DESPLEGABLE"]}"'
    return context
