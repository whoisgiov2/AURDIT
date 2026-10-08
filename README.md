# Auditoría de Accesos Post-Baja

Herramienta en Python para la auditoría de accesos posteriores a la baja de colaboradores (*Leaver Access Review*). Cruza el padrón de bajas de la organización con un reporte de logins exportado desde un Identity Provider, clasifica cada caso según una matriz de riesgo y genera un libro Excel con el resultado.

## Contenido

1. [Descripción general](#1-descripción-general)
2. [Requisitos e instalación](#2-requisitos-e-instalación)
3. [Uso](#3-uso)
4. [Archivos de entrada](#4-archivos-de-entrada)
5. [Proceso de auditoría](#5-proceso-de-auditoría)
6. [Matriz de riesgo](#6-matriz-de-riesgo)
7. [Salida](#7-salida)
8. [Pruebas](#8-pruebas)
9. [Estructura del repositorio](#9-estructura-del-repositorio)

---

## 1. Descripción general

El sistema procesa de forma vectorizada dos orígenes de datos:

| Origen | Descripción |
| :--- | :--- |
| Archivo Universo | Padrón maestro de empleados y bajas. Debe contener la pestaña `BajasU`. |
| Reporte de Logins | Exportación (`report<timestamp>.*`) de un IdP o sistema de identidades (Salesforce, CRM, Active Directory, Okta, etc.) en `.xlsx`, `.xls`, tabla HTML con extensión `.xls`, o `.csv`. |

Sobre ambos aplica:

- Deduplicación cronológica del Universo y del reporte.
- Validación compuesta (correo + nombre) para evitar falsos positivos por cuentas compartidas o reasignadas.
- Matriz de riesgo con cinco categorías.

El resultado es el libro `Auditoria_Accesos_Resultado.xlsx` con cinco pestañas.

## 2. Requisitos e instalación

- Python 3.10 o superior.
- Dependencias (`requirements.txt`): `pandas`, `numpy`, `openpyxl`, `xlrd` (lectura de `.xls`) y `pytest`.

```bash
pip install -r requirements.txt
```

## 3. Uso

### Ejecución interactiva

```bash
python AUDIT.py
```

En Windows también puede ejecutarse `AUDIT.bat`, que reenvía los argumentos a `AUDIT.py`. El programa solicita, en orden:

1. **Ruta del archivo Universo.** Se ofrece como valor por defecto (Enter para aceptarlo) el archivo detectado automáticamente en la carpeta actual o, si no hay, la última ruta usada.
2. **Ruta del reporte de logins**, salvo que se haya detectado automáticamente (`report<dígitos>` en la carpeta actual) o se indique con `--reporte`.
3. **Carpeta de resultados:**
   - **Primera vez:** pide la ruta y, si la carpeta no existe, pregunta antes de crearla. La carpeta debe estar fuera de la carpeta de scripts; de lo contrario se rechaza.
   - **Ejecuciones siguientes:** pregunta si los resultados se guardan en la última carpeta usada. Si se responde que no, pide una nueva ruta (con las mismas validaciones) y la recuerda.

El archivo de resultados se guarda como `Auditoria_Accesos_Resultado.xlsx` dentro de la carpeta elegida.

### Configuración guardada

La última carpeta de resultados y la última ruta del Universo se guardan en `~/.auditoria_accesos/config.json` (fuera del repositorio). Puede cambiarse la ubicación con la variable de entorno `AUDITORIA_CONFIG`.

### Ejecución no interactiva

```bash
python AUDIT.py \
  --universo "Universo_Usuarios_PROD_2.xlsx" \
  --reporte "report1790717536569.xlsx" \
  --carpeta-resultados "D:/Auditorias/2026-10"
```

Con `--carpeta-resultados` no se hacen preguntas de carpeta (se crea si no existe) y la ruta queda guardada para la próxima vez. Con `--salida` se indica la ruta completa del Excel y no se consulta ni se guarda la carpeta. Sin consola interactiva y con datos faltantes, el programa termina con error.

### Argumentos de línea de comandos

| Argumento | Descripción | Valor por defecto |
| :--- | :--- | :--- |
| `-u`, `--universo` | Ruta al archivo Universo. | Se pregunta |
| `--hoja-universo` | Pestaña a leer en el Universo. | `BajasU` |
| `-r`, `--reporte` | Ruta al reporte de logins (`.xlsx`, `.xls`, `.csv`). | Detección automática o se pregunta |
| `--carpeta-resultados` | Carpeta de resultados (fuera de la carpeta de scripts). | Se pregunta |
| `-o`, `--salida` | Ruta completa del Excel de resultados. | Ninguno |
| `--auto` | Sin efecto funcional; se conserva por compatibilidad. | Desactivado |

Notas de comportamiento:

- La detección automática considera archivos `.xlsx`, `.xls` y `.csv` de la carpeta actual, ignorando temporales de Office (`~$*`) y nombres que contengan `resultado`. El reporte es el que empiece por `report` con dígitos (o contenga `report`/`login`); el Universo, el que contenga `prod`, `universo`, `padron`, `baja`, `empleado` o `master`.
- Si el archivo de salida está abierto en otra aplicación, se guarda con un sufijo de fecha y hora (`<nombre>_YYYYMMDD_HHMMSS.xlsx`).
- El proceso retorna código `1` ante errores de entrada o de ejecución y `0` en caso de éxito.
- Al finalizar se imprime en consola un resumen con las métricas de la ejecución.

## 4. Archivos de entrada

### 4.1 Archivo Universo

| Elemento | Detalle |
| :--- | :--- |
| Pestaña | `BajasU`: coincidencia exacta (sin distinguir mayúsculas) o, en su defecto, por subcadena. Si no existe, se aborta con un error que lista las pestañas halladas. |
| Columnas mínimas | 10. Las filas completamente vacías se descartan. |
| Nombre (Col. C, índice 2) | Se busca por nombre (`nombre`, `nombre_completo`, `empleado`, `trabajador`, `full name`, `name`) y, si no aparece, por posición. |
| `FECHA_BAJA` (Col. D, índice 3) | Se busca por nombre (`fecha_baja`, `fecha baja`, `fechabaja`, `fec_baja`, `baja`) y, si no aparece, por posición. |
| Correo (Col. J, índice 9) | Se busca por nombre (`Unnamed: 9`, `(No column name)`, `correo`, `email`, `mail`, `correo_corporativo`, `username`) y, si no aparece, por posición. |

Formatos de fecha admitidos: `DD/MM/YYYY` con o sin hora (día primero), ISO `YYYY-MM-DD` y valores de fecha nativos de Excel. Los valores `0`, vacíos y `nan` se tratan como "sin fecha".

### 4.2 Reporte de Logins

**Lectura del archivo**

1. Si los primeros bytes contienen etiquetas HTML, se interpreta como tabla HTML (parser propio, sin dependencias externas; prueba codificaciones UTF-8, ISO-8859-1 y Windows-1252).
2. Si la extensión es `.csv`, se lee como CSV.
3. En otro caso se lee como libro Excel, eligiendo la pestaña en este orden: `in`; la única pestaña del libro; una que contenga `report`, `login`, `hoja1` o `sheet1`; la primera disponible. Si la lectura falla, se reintenta como HTML.

**Preprocesamiento defensivo**

- `dropna(how='all', axis=1)` para eliminar columnas fantasma generadas por celdas combinadas o exportaciones web.
- Eliminación de filas vacías y recorte de espacios en los encabezados.
- Se requieren al menos dos columnas tras la limpieza.

**Mapeo de columnas**

La búsqueda no distingue mayúsculas: primero coincidencia exacta, luego subcadena y, como último recurso, posición.

| Campo lógico | Candidatos (en orden) | Posición por defecto |
| :--- | :--- | :---: |
| `name` (nombre completo) | `full name`, `full_name`, `nombre`, `nombre_completo`, `name`, `empleado` | 0 |
| `user` (llave de cruce) | `username`, `usuario`, `correo`, `email`, `login`, `user name` | 2 |
| `active` | `active`, `activo`, `estatus`, `status` | 3 |
| `login` (último acceso) | `last login`, `ultimo login`, `last_login`, `ultimo_login`, `login date`, `fecha_login`, `login` | 4 |
| `baja_rep` (baja en reporte) | `fecha de ba`, `fecha baja`, `fecha_baja`, `fechabaja`, `baja` (excluye columnas con `tipo` o `motivo`) | 5 |
| `created` | `created date`, `created`, `creacion`, `creac`, `fecha creacion` | 6 |

El indicador `Active` se normaliza a `1` o `0`: se consideran activos los valores `1`, `1.0`, `true`, `t`, `si`, `yes`, `y`, `activo` y `active`; cualquier otro valor equivale a `0`.

## 5. Proceso de auditoría

### 5.1 Normalización y llave compuesta

- Correo: recorte de espacios y minúsculas. Los valores `""`, `nan`, `none`, `null`, `nat`, `0` y `0.0` se tratan como ausentes.
- Nombre: minúsculas, eliminación de acentos y diacríticos (`NFKD` a ASCII) y colapso de espacios repetidos.
- Llave de cruce: `_key_compuesta = correo + "___" + nombre normalizado`.
- Si el reporte no contiene nombres utilizables, el cruce se realiza solo por correo.
- Todas las comparaciones de fecha se hacen a nivel de día calendario (se descarta la hora), para no penalizar accesos legítimos del día del cese.

### 5.2 Deduplicación del Universo (`deduplicate_universo`)

Los registros se agrupan por llave compuesta:

| Situación del grupo | Tratamiento |
| :--- | :--- |
| Un solo registro | Se conserva. |
| Registros con y sin `FECHA_BAJA` | Se identifica como reingreso: se conservan todos los registros y la llave queda marcada para clasificarse como `POSIBLE REINGRESO`. |
| Todos con `FECHA_BAJA` | Se ordenan por fecha descendente y se conserva solo el de baja más reciente. |
| Registros sin correo | Se conservan sin agrupar. |

### 5.3 Consolidación del reporte (`consolidate_logins`)

Por cada llave compuesta del reporte se obtiene un único registro:

- **Active:** se conserva el valor máximo; si la cuenta figura activa en cualquier registro, se considera activa.
- **Baja en reporte:** se conserva la fecha más reciente, usada como respaldo cuando el Universo no tiene fecha.
- **Login:** se guardan todos los logins válidos y se elige el más cercano a la fecha de baja del colaborador (ver siguiente punto).

### 5.4 Selección del login (`select_closest_login_event`)

Con la fecha de baja efectiva, se elige el login con menor `|fecha_login - fecha_baja|`.

- En empate exacto de distancia se prioriza el login posterior a la baja.
- Si no hay fecha de baja, se toma el login más reciente.
- Si no hay logins válidos, el resultado es vacío.

La fecha de baja efectiva es la del Universo; si falta, se usa la del reporte.

### 5.5 Validación de identidad

- Si la llave compuesta del Universo coincide con la del reporte, el colaborador queda enlazado a su login y estado de cuenta.
- Si el correo existe en el reporte pero asociado a un nombre distinto, no se enlaza el login y se marca `DISCREPANCIA_IDENTIDAD = True` para revisión de TI/IAM.
- Los colaboradores sin coincidencia se muestran como `Sin registro` en `ULTIMO_LOGIN_DETECTADO` y `ESTATUS_CUENTA_REPORTE`.

### 5.6 Detección de reingresos

Un registro se clasifica como reingreso si, para la misma identidad, existe al menos un registro con baja y otro sin baja. La comparación se evalúa por cualquiera de estos criterios:

- Correo.
- Nombre normalizado.
- Clave de trabajador (columna `cla_trab`, `id_empleado`, `num_emp` o `empleado_id`, si existe).
- Llave compuesta marcada durante la deduplicación del Universo.

Los reingresos no se penalizan por bajas de relaciones laborales anteriores.

## 6. Matriz de riesgo

Aplica a registros que no son reingreso. "Acceso post-baja" significa `fecha_login > fecha_baja`.

| Acceso post-baja | Cuenta (`Active`) | `ESTATUS_AUDITORIA` | `CATEGORIA_RIESGO` | Acción requerida |
| :---: | :---: | :--- | :--- | :--- |
| Sí | Activa | `REVISAR` | `CRÍTICO - RIESGO ACTIVO` | Bloqueo inmediato de la cuenta e inicio de investigación. |
| No o sin login | Activa | `REVISAR` | `ALTO - CUENTA HUÉRFANA` | Desactivación preventiva de la cuenta. |
| Sí | Inactiva | `INCIDENTE RESUELTO` | `MEDIO - INCIDENTE PASADO` | Documentar como evidencia histórica. |
| No o sin login | Inactiva | `OK` | `CONFORME` | Ninguna; control aplicado correctamente. |
| No aplica | Cualquiera | `OK` | `POSIBLE REINGRESO` | Informativa; conciliar con Recursos Humanos. |

Notas:

- Los registros sin fecha de baja válida se clasifican como `CONFORME`.
- Los colaboradores `Sin registro` en el reporte se consideran con cuenta inactiva.
- `DIAS_POST_BAJA` solo se calcula cuando existe acceso post-baja.

## 7. Salida

El archivo `Auditoria_Accesos_Resultado.xlsx` contiene cinco pestañas:

| # | Pestaña | Contenido |
| :---: | :--- | :--- |
| 1 | `Glosario_y_Criterios` | Estructura del libro, diccionario de estatus con condiciones, impacto y acciones, y reglas técnicas (corte calendario, baja más reciente, validación compuesta). |
| 2 | `Auditoria_Completa` | Padrón deduplicado con todas las columnas originales más las columnas de auditoría. |
| 3 | `Riesgos_Activos` | Casos `CRÍTICO - RIESGO ACTIVO` y `ALTO - CUENTA HUÉRFANA`, ordenados por severidad y luego por días post-baja descendente. |
| 4 | `Incidentes_Pasados` | Casos `MEDIO - INCIDENTE PASADO`, ordenados por último login descendente. |
| 5 | `Reingresos` | Casos `POSIBLE REINGRESO`. |

### Columnas de auditoría añadidas

| Columna | Descripción |
| :--- | :--- |
| `ULTIMO_LOGIN_DETECTADO` | Fecha del login seleccionado (`DD/MM/YYYY`), vacío si no hay login válido o `Sin registro`. |
| `ESTATUS_CUENTA_REPORTE` | `Activa`, `Inactiva` o `Sin registro`. |
| `DISCREPANCIA_IDENTIDAD` | `True` si el correo existe en el reporte con otro nombre. |
| `ESTATUS_AUDITORIA` | `REVISAR`, `INCIDENTE RESUELTO` u `OK`. |
| `CATEGORIA_RIESGO` | Categoría de la matriz de riesgo. |
| `TIPO_HALLAZGO` | Mismo valor que `CATEGORIA_RIESGO` (se mantiene por compatibilidad). |
| `DIAS_POST_BAJA` | Días entre la baja y el login, solo con acceso post-baja. |

`FECHA_BAJA` se reescribe en formato `DD/MM/YYYY` cuando existe una fecha válida.

### Formato visual

Encabezados oscuros (`#1F2937`), primera fila inmovilizada, ancho de columnas ajustado y resaltado semántico:

| Color | Relleno / Texto | Aplica a |
| :--- | :--- | :--- |
| Rojo | `#FEE2E2` / `#991B1B` | `CRÍTICO - RIESGO ACTIVO` |
| Ámbar | `#FEF3C7` / `#92400E` | `ALTO - CUENTA HUÉRFANA`, discrepancias de identidad |
| Gris Slate | `#F1F5F9` / `#334155` | `MEDIO - INCIDENTE PASADO`, `INCIDENTE RESUELTO` |
| Azul | `#E0F2FE` / `#0369A1` | `POSIBLE REINGRESO` |
| Verde | `#DCFCE7` / `#166534` | `CONFORME`, `OK` |

## 8. Pruebas

```bash
pytest -v
```

La suite contiene 17 pruebas en `tests/test_auditoria.py` y `tests/test_auditor.py`. Dos de ellas validan los archivos reales de producción y se omiten (`skipped`) si esos archivos no están presentes en el entorno, por lo que en un clon limpio el resultado esperado es 15 aprobadas y 2 omitidas.

| Archivo | Cobertura |
| :--- | :--- |
| `tests/test_auditoria.py` | Discrepancia de identidad con mismo correo y nombres distintos; normalización de acentos y espacios; cuenta activa post-baja; criterios de la matriz de riesgo; archivos reales del espacio de trabajo; contenido del glosario; múltiples fechas de baja; login más cercano; desempate hacia el login posterior. |
| `tests/test_auditor.py` | Configuración persistente y selección de carpeta de resultados; parseo de fechas (`DD/MM/YYYY`, ISO, vacíos); parseo de `Active`; columnas fantasma y cabeceras truncadas; consolidación de múltiples logins; matriz de riesgo integrada; archivos reales del repositorio; creación aislada del glosario. |

## 9. Estructura del repositorio

| Ruta | Descripción |
| :--- | :--- |
| `AUDIT.py` | Punto de entrada; delega en `auditor_bajas.main()`. |
| `AUDIT.bat` | Lanzador para Windows. |
| `auditor_bajas.py` | Motor completo: lectura, normalización, deduplicación, cruce, matriz de riesgo, exportación y CLI. |
| `requirements.txt` | Dependencias. |
| `tests/` | Suite de pruebas automatizadas. |
