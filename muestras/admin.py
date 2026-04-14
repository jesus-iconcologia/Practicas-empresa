from django.contrib import admin
from .models import Muestra, Localizacion, Estudio, Envio, Documento,historial_estudios,historial_localizaciones,registro_destruido, Congelador, Subposicion, Caja, Rack, Estante, Material
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin
from django.contrib.auth.models import User

# Se posibilita que el usuario admin añada, elimine y cambie los datos de todos los modelos de la aplicación 

admin.site.register(Muestra)
admin.site.register(Localizacion)
admin.site.register(Estudio)
admin.site.register(Envio)
admin.site.register(Documento)
admin.site.register(historial_estudios)
admin.site.register(historial_localizaciones)
admin.site.register(registro_destruido)
admin.site.register(Congelador)
admin.site.register(Estante)
admin.site.register(Caja)
admin.site.register(Rack)
admin.site.register(Subposicion)


@admin.register(Material)
class MaterialAdmin(admin.ModelAdmin):
    list_display = (
        'nombre',
        'activo',
        'usa_volumen',
        'usa_concentracion',
        'usa_unidades_bloque',
    )
    list_filter = (
        'activo',
        'usa_volumen',
        'usa_concentracion',
        'usa_unidades_bloque',
    )
    search_fields = ('nombre',)

# Clases y funcionalidades necesarias para que el admin asocie los estudios a los investigadores
class EstudioInline(admin.TabularInline):
    model = Estudio.investigadores_asociados.through  
    extra = 1  
    verbose_name = "Estudio asociado"
    verbose_name_plural = "Estudios asociados"
class CustomUserAdmin(BaseUserAdmin):
    def get_inlines(self, request, obj=None):
        # Si no hay objeto (estamos creando un usuario nuevo) no mostramos estudios
        if obj is None:
            return []
        
        # Verificamos si el usuario pertenece al grupo 'Investigadores'
        if obj.groups.filter(name='Investigadores').exists():
            return [EstudioInline]
        
        # Si no es investigador, no devolvemos ningún inline
        return []

admin.site.unregister(User)
admin.site.register(User, CustomUserAdmin)
