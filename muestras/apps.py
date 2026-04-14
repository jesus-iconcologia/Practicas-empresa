from django.apps import AppConfig

class MuestrasConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'muestras'

    def ready(self):
        import muestras.signals.documents
