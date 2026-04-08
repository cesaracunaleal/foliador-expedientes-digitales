# Generador de Índice + Foliador Digital (v7.0)

Herramienta para la organización, foliado y verificación de integridad de expedientes digitales en entornos institucionales.

-----

## Descripción

La Comisión Regional de Uso del Borde Costero (CRUBC) de la Región de Los Ríos gestiona un volumen creciente de expedientes digitales vinculados a solicitudes de concesiones, informes técnicos y actas de sesión. La dispersión de criterios de nombrado, la ausencia de foliado sistemático y la dificultad para verificar la integridad de los expedientes generaban pérdida de tiempo y riesgos de errores en procesos de revisión y auditoría.

Esta aplicación automatiza la construcción de índices documentales y el foliado de expedientes, incorporando mecanismos de trazabilidad, validación y control de integridad. Su arquitectura permite adaptarla a otros contextos administrativos mediante configuración externa, sin modificar el código fuente.

-----

## Funcionalidades

### Procesamiento documental

- Recorrido automático de carpetas y subcarpetas
- Identificación de fecha (desde nombre, contenido o metadata), tipo documental y número de documento cuando aplica
- Exclusión configurable de archivos auxiliares (componentes SHP, carpetas de salida, etc.)

### Foliado estructurado

- Generación de nombres normalizados
- Eliminación de redundancias (fecha, tipo, número)
- Incorporación del número de documento cuando existe
- Ordenamiento consistente del expediente

### Integridad y trazabilidad

- Hash SHA-256 por archivo y hash global del expediente
- Registro de fuente de extracción (nombre, texto, metadata) y regla aplicada
- Historial de ediciones
- Generación de log de auditoría

### Rendimiento

- Procesamiento paralelo mediante `ThreadPoolExecutor`
- Sistema de caché basado en cambios de archivo (`mtime` + tamaño)
- Optimizado para expedientes con alto volumen documental

### Exportación

|Formato|Contenido                                |
|-------|-----------------------------------------|
|Excel  |Índice completo con metadatos y hash     |
|Word   |Versión formal del índice                |
|Log    |Advertencias, validaciones y trazabilidad|

-----

## Configuración externa

El sistema utiliza un archivo `config_indice.json` que se genera automáticamente al primer uso. Permite adaptar:

```json
{
  "institucion": {
    "nombre": "GOBIERNO REGIONAL DE LOS RÍOS",
    "subtitulo": "Comisión Regional de Uso del Borde Costero",
    "pie": "Oficina Técnica CRUBC"
  },
  "dominio_correo": "gorelosrios.cl",
  "tipos_adicionales": {
    "nombre": [],
    "texto": {}
  }
}
```

Sin necesidad de modificar el código fuente.

-----

## Pruebas automáticas

Las pruebas están integradas en el script principal y cubren:

- Detección de fechas
- Clasificación documental
- Foliado
- Integridad (hash SHA-256)
- Funcionamiento de caché

Ejecutar:

```bash
python generador_indice.py --test
```

-----

## Arquitectura

- Modelo de datos mediante `dataclass` (`DocumentoAnalizado`)
- Clase de dominio (`ExpedienteDigital`) que encapsula la lógica del expediente
- Separación conceptual entre análisis, foliado, exportación e interfaz

-----

## Tecnologías

|Componente |Tecnología          |
|-----------|--------------------|
|Lenguaje   |Python 3.10+        |
|Hashing    |`hashlib`           |
|Paralelismo|`concurrent.futures`|
|Excel      |`openpyxl`          |
|Word       |`python-docx`       |
|Interfaz   |`tkinter`           |
|PDF        |`pdfplumber`        |

-----

## Instalación

```bash
pip install pdfplumber python-docx openpyxl pillow
```

```bash
python generador_indice.py
```

Opciones adicionales:

- Lanzador `.vbs` para ejecución sin consola en Windows
- Archivo `.bat` para depuración

-----

## Aplicación en gestión pública

Esta herramienta contribuye a la estandarización documental, la reducción de errores manuales y la trazabilidad de expedientes. Está diseñada como una capa operativa complementaria a los sistemas documentales formales, útil especialmente en la preparación de expedientes previo a procesos de revisión, auditoría o entrega institucional.

-----

## Autor

**César Acuña Leal**  
Geógrafo — Especialista en gestión territorial y borde costero  
Gobierno Regional de Los Ríos, Chile

-----

## Licencia

Este proyecto está bajo la [Apache License 2.0](LICENSE). Permite uso, modificación y distribución libre con atribución al autor original, e incluye protección explícita de patentes.
