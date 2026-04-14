"""
Archivo de parametrización para uploads de Excel.
Define mensajes, colores y estilos para muestras, estudios y localizaciones.
"""

# ============================================================================
# CONFIGURACIÓN DE MENSAJES POR TIPO DE UPLOAD
# ============================================================================

UPLOAD_MESSAGES = {
    'muestras': {
        'titulo_inicial': 'El excel contiene',
        'sin_errores': 'No tiene errores en ningún campo.',
        'con_advertencias': 'Contiene {count} filas con advertencias',
        'con_bloqueantes': 'Contiene {count} filas con errores graves',
        'columnas_extras': 'Contiene {count} columnas extras: {detalles}',
    },
    'estudios': {
        'titulo_inicial': 'El excel contiene',
        'sin_errores': 'No tiene errores en ningún campo.',
        'con_advertencias': 'Contiene {count} filas con advertencias',
        'con_bloqueantes': 'Contiene {count} filas con errores graves',
        'columnas_extras': 'Contiene {count} columnas extras: {detalles}',
    },
    'localizaciones': {
        'titulo_inicial': 'El excel contiene',
        'sin_errores': 'No tiene errores en ningún campo.',
        'con_advertencias': None,  # No se usa para localizaciones
        'con_bloqueantes': 'Contiene {count} filas con errores graves',
        'columnas_extras': 'Contiene {count} columnas extras: {detalles}',
    },
    'cambio_posicion': {
        'titulo_inicial': 'El excel contiene',
        'sin_errores': 'No tiene errores en ningún campo.',
        'con_advertencias': None,  # No se usa para cambio de posición
        'con_bloqueantes': 'Contiene {count} filas con errores graves',
        'columnas_extras': 'Contiene {count} columnas extras: {detalles}',
    },
}

# ============================================================================
# CONFIGURACIÓN DE COLORES PARA EXCEL
# ============================================================================

EXCEL_COLORS = {
    # Colores para filas con errores
    'error_row': 'F8D7DA',        # Rojo claro - fondo de fila con error
    'error_cell': 'F5C2C7',       # Rojo fuerte - celdas específicas con error
    
    # Colores para filas con advertencias
    'warning_row': 'FFF3CD',      # Amarillo claro - fondo de fila con advertencia
    'warning_cell': 'FFECB5',     # Amarillo fuerte - celdas específicas con advertencia
    
    # Colores para columnas extras
    'extra_column': 'F5C2C7',     # Rojo fuerte - igual que error_cell
}

# ============================================================================
# FUNCIÓN AUXILIAR PARA OBTENER COLORES (facilita refactorización)
# ============================================================================

def get_excel_colors():
    """
    Retorna un diccionario con los colores para usar en openpyxl.
    Uso: colors = get_excel_colors()
         fill = PatternFill("solid", fgColor=colors['error_row'])
    """
    return EXCEL_COLORS.copy()

def get_upload_messages(upload_type):
    """
    Retorna los mensajes para un tipo de upload específico.
    Uso: messages = get_upload_messages('muestras')
         titulo = messages['titulo_inicial']
    """
    if upload_type not in UPLOAD_MESSAGES:
        raise ValueError(f"Tipo de upload no soportado: {upload_type}")
    return UPLOAD_MESSAGES[upload_type].copy()
