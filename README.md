# Sistema de Gestión de Responsivas - VUBA LOGISTICS

## 📋 Descripción
Sistema profesional para generar responsivas de equipos de cómputo con base de datos SQLite local y exportación a Excel y PDF.

## ✨ Características

- ✅ **Base de datos SQLite local** (no requiere SQL Server)
- ✅ **Exportación a Excel** con formato profesional
- ✅ **Exportación a PDF** con diseño limpio
- ✅ **Interfaz gráfica moderna** con Tkinter
- ✅ **Gestión de empleados y equipos**
- ✅ **Registro de accesorios asignados**
- ✅ **Historial de responsivas generadas**

## 🔧 Requisitos

### Python
- Python 3.8 o superior

### Librerías necesarias
```bash
pip install openpyxl reportlab
```

## 📦 Instalación

1. **Instala las dependencias:**
```bash
pip install openpyxl reportlab
```

2. **Ejecuta el programa:**
```bash
python sistema_responsivas.py
```

## 🚀 Uso del Sistema

### Primera Ejecución
Al ejecutar el programa por primera vez, se creará automáticamente:
- Base de datos `responsivas.db` con datos de ejemplo
- Carpeta `Responsivas_Generadas/` para los archivos generados

### Búsqueda de Empleados
1. Escribe el nombre del empleado en el campo de búsqueda
2. Presiona Enter o clic en "🔍 Buscar"
3. Se mostrarán todos los datos del empleado

### Generar Responsivas
1. Busca un empleado
2. Clic en "📊 Generar Excel" o "📄 Generar PDF"
3. El archivo se guardará en `Responsivas_Generadas/`

### Ver Todos los Empleados
- Clic en "📋 Ver Todos" para ver la lista completa
- Doble clic en un empleado para seleccionarlo

## 📊 Estructura de la Base de Datos

### Tablas creadas:
- **empresas**: Empresas registradas
- **areas**: Departamentos/áreas
- **usuarios**: Empleados y equipos asignados
- **accesorios**: Accesorios por empleado
- **responsivas**: Historial de documentos generados

### Datos de ejemplo incluidos:
- 1 empresa (VUBA LOGISTICS)
- 5 áreas (Tecnología, Administración, RRHH, Operaciones, Logística)
- 3 empleados con equipos y accesorios

## 📄 Formatos de Exportación

### Excel (.xlsx)
- Diseño profesional con colores corporativos
- Tablas organizadas por secciones
- Formato listo para imprimir
- Espacios para firmas

### PDF (.pdf)
- Diseño limpio y profesional
- Optimizado para impresión
- Incluye todas las secciones requeridas
- Firmas de responsabilidad

## 🗂️ Estructura de Archivos

```
proyecto/
│
├── sistema_responsivas.py      # Programa principal
├── responsivas.db              # Base de datos SQLite (se crea automáticamente)
├── README.md                   # Este archivo
└── Responsivas_Generadas/      # Carpeta de salida (se crea automáticamente)
    ├── Responsiva_Juan_Perez_20240122_143022.xlsx
    └── Responsiva_Juan_Perez_20240122_143025.pdf
```

## 🔄 Diferencias con la versión anterior

### ✅ Mejoras implementadas:

1. **Base de datos local SQLite** (antes: SQL Server remoto)
   - No requiere servidor
   - Más fácil de usar
   - Portable

2. **Código más limpio y organizado**
   - Separación de responsabilidades
   - Clases bien definidas
   - Mejor manejo de errores

3. **Exportaciones mejoradas**
   - Excel más limpio y profesional
   - PDF optimizado
   - Mejor formato visual

4. **Interfaz mejorada**
   - Scroll en la información
   - Vista de todos los usuarios
   - Mejor feedback visual

## 🐛 Solución de Problemas

### Error: "ModuleNotFoundError: No module named 'openpyxl'"
**Solución:**
```bash
pip install openpyxl
```

### Error: "ModuleNotFoundError: No module named 'reportlab'"
**Solución:**
```bash
pip install reportlab
```

### La base de datos no se crea
**Solución:**
- Verifica que tienes permisos de escritura en la carpeta
- Ejecuta el programa desde la línea de comandos para ver errores

## 📝 Personalización

### Cambiar colores corporativos
En `ExcelGenerator` y `PDFGenerator`, modifica:
```python
COLOR_HEADER = '1F4E78'  # Azul corporativo
COLOR_SECTION = 'D9E1F2'  # Azul claro
```

### Agregar más campos
1. Modifica la tabla `usuarios` en `DatabaseManager.init_database()`
2. Actualiza `buscar_usuario()` para incluir los nuevos campos
3. Modifica las plantillas Excel/PDF para mostrarlos

## 🆘 Soporte

Para reportar problemas o sugerencias:
- Revisa el código en las líneas indicadas en el error
- Verifica que todas las dependencias estén instaladas
- Asegúrate de tener Python 3.8+

## 📜 Licencia

Uso interno - VUBA LOGISTICS

---

**Versión:** 2.0  
**Última actualización:** Enero 2025  
**Desarrollado para:** Departamento de TI - VUBA LOGISTICS
