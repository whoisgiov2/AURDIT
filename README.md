# Auditoría de Accesos Post-Baja (Leaver Access Review Automation)

Herramienta en Python de alto rendimiento y grado de producción para la automatización del proceso de auditoría de seguridad, detección de cuentas huérfanas, conciliación de accesos pos-desvinculación (*Leaver Access Review*) y clasificación tripartita de riesgos de ciberseguridad.

---

## 📋 Objetivo Principal

El sistema procesa y cruza automáticamente de forma **vectorizada** dos orígenes de datos:
1. **Archivo Universo:** Padrón maestro de empleados y bajas de la organización (pestaña obligatoria `'BajasU'`).
2. **Archivo Reporte de Logins:** Reporte con nombre dinámico (`report<timestamp>.*`) exportado desde Identity Providers (Salesforce, CRM, Active Directory, Okta, etc.) en formatos Excel (`.xlsx` / `.xls`), tablas HTML (`.xls`) o `.csv`.

Aplica un **Algoritmo Cronológico de Deduplicación**, **Validación Compuesta Anticolisiones** y una **Matriz de Riesgo Tripartita**, generando un libro Excel ejecutivo (`Auditoria_Accesos_Resultado.xlsx`) compuesto por **5 pestañas** con formateo profesional openpyxl, glosario metodológico integrado y semáforos visuales de riesgo.

---

## ⚙️ Estructura de Entradas y Preprocesamiento Defensivo

### 1. Archivo Universo (Padrón Maestro)
- **Pestaña obligatoria:** `'BajasU'` (detección exacta o por coincidencia de subcadena).
- **Columna C (Índice 2):** Nombre completo del colaborador (`NOMBRE`).
- **Columna D (Índice 3):** `FECHA_BAJA` (soporta DD/MM/YYYY con o sin hora, marcas temporales Unix y formato ISO YYYY-MM-DD).
- **Columna J (Índice 9):** Correo corporativo (incluso con cabeceras genéricas como `(No column name)` o `Unnamed: 9`).
- **Validación defensiva:** Requiere un mínimo de 10 columnas y descarta filas vacías automáticamente.

### 2. Archivo Reporte de Logins
- **Formatos:** Libros Excel (pestaña `'in'` o primera pestaña disponible), tablas HTML exportadas con extensión `.xls` y archivos `.csv`.
- **Preprocesamiento Defensivo Obligatorio:**
  - Aplica `df.dropna(how='all', axis=1)` inmediatamente para purgar columnas fantasma generadas por celdas combinadas o desalineadas en exportaciones web.
  - Elimina filas completamente vacías y normaliza nombres de columnas (recorte de espacios).
- **Mapeo Resiliente de Columnas (búsqueda insensible a mayúsculas y acentos):**
  - **`Full Name`:** Columna con `"full name"`, `"nombre"`, `"name"` o `"empleado"` (Índice 0 por defecto).
  - **`Username`:** Columna con `"username"`, `"usuario"`, `"email"` o `"correo"` (Índice 2 por defecto).
  - **`Active`:** Columna con `"active"` o `"activo"`. Parseo binario robusto (`1` o `0`), aceptando cadenas de texto (`'1'`, `'true'`, `'si'`, `'activo'`).
  - **`Last Login`:** Columna con `"last login"`, `"ultimo login"`, `"login"` o `"fecha_login"`.
  - **`Fecha de Baja (Reporte)`:** Columna con `"fecha de ba"` o `"baja"` (excluyendo motivos o tipos).
  - **`Created Date`:** Columna con `"created"` o `"creac"`.

---

## 🔄 Algoritmo Cronológico de Deduplicación y Selección Inteligente

Para resolver empates y múltiples registros de una misma identidad (`_key_compuesta = Username + "___" + Nombre Normalizado`):

### 1. Deduplicación en Universo (`deduplicate_universo`)
Al agrupar registros por colaborador:
* **Paso 1 (Detección de Reingreso Activo):**
  - Si la persona tiene registros donde `Fecha_Baja` es nula/vacía **Y** registros con `Fecha_Baja` confirmada, se identifica como **Reingreso Activo**.
  - Se clasifica con `ESTATUS_AUDITORIA = 'OK'` y `CATEGORIA_RIESGO = 'POSIBLE REINGRESO'`.
  - Se preservan sus registros para auditoría, garantizando que sus accesos laborales vigentes no sean penalizados por bajas de relaciones laborales anteriores.
* **Paso 2 (Múltiples Bajas Históricas):**
  - Si todos los registros de la persona tienen fecha de baja confirmada:
    1. Se ordenan cronológicamente por `_parsed_baja` descendente (`ascending=False`).
    2. Se conserva únicamente el registro con la **fecha de baja más reciente (máxima)**.
    3. Se descartan las fechas de baja antiguas (eliminando falsos positivos por bajas superadas).

### 2. Selección Cronológica de Login en Reporte (`select_closest_login_event`)
Determinada la `Fecha_Baja_Efectiva` del colaborador:
* Si el usuario registra múltiples eventos de login en el reporte:
  - Se calcula la distancia temporal absoluta en días:
    $$\text{distancia} = |\text{Fecha\_Login} - \text{Fecha\_Baja\_Efectiva}|$$
  - Se selecciona el registro de login que posea la **menor distancia temporal (el más cercano)** respecto a su fecha de baja.
  - **Criterio Estricto de Desempate (Tie-Breaker):** En caso de empate exacto de distancia (ej. un login 2 días antes y otro 2 días después de la baja), se prioriza el **login posterior** (`login_date > fecha_baja`) para garantizar que ningún acceso no autorizado post-cese sea omitido.
  - **Preservación de Estado de Cuenta:** Si la cuenta figura activa (`Active == 1`) en cualquiera de los registros del usuario en el reporte, se conserva el estado activo (`_parsed_active.max()`).

---

## 🛡️ Validación Compuesta y Protección Anticolisiones

Para evitar falsos positivos causados por cuentas compartidas, cuentas genéricas o buzones reasignados:
1. **Llave Compuesta de Cruce:** `_key_compuesta = _clean_email + "___" + _clean_nombre`.
2. **Normalización Fonética y Ortográfica Profunda:**
   - Eliminación de acentos, tildes y diacríticos (`unicodedata.normalize('NFKD')`).
   - Supresión de caracteres especiales, puntuación y unificación a minúsculas.
   - Colapso de dobles o triples espacios en blanco.
3. **Detección de Discrepancia de Identidad (`DISCREPANCIA_IDENTIDAD = True`):**
   - Si un correo existe en el reporte de logins pero está asignado a un nombre distinto al del padrón de bajas, el sistema **no vincula arbitrariamente el login** al empleado cesado.
   - Marca la alerta `DISCREPANCIA_IDENTIDAD = True` para revisión del equipo de TI / IAM y evita falsos incidentes post-baja.

---

## 🧠 Matriz de Evaluación de Riesgos Tripartita

La auditoría clasifica cada caso evaluado bajo la siguiente matriz operativa:

| Acceso Post-Baja | Estado Cuenta (`Active`) | Estatus Auditoría | Categoría de Riesgo | Acción Operativa Requerida |
| :---: | :---: | :--- | :--- | :--- |
| **SÍ** (`Login > Baja`) | `1` (Activa) | `REVISAR` | **CRÍTICO - RIESGO ACTIVO** | 🚨 **Bloqueo Inmediato:** Desactivar cuenta en Directorio Activo / IdP e iniciar investigación forense. |
| **NO** (`Login <= Baja` / Sin login) | `1` (Activa) | `REVISAR` | **ALTO - CUENTA HUÉRFANA** | ⚠️ **Desactivación Preventiva:** Cuenta aún encendida post-baja; mitigar riesgo de intrusión latente. |
| **SÍ** (`Login > Baja`) | `0` (Inactiva) | `INCIDENTE RESUELTO` | **MEDIO - INCIDENTE PASADO** | 📋 **Documentación:** El usuario accedió tras su cese, pero la cuenta ya fue apagada. Evidencia histórica. |
| **NO** (`Login <= Baja` / Sin login) | `0` (Inactiva) | `OK` | **CONFORME** | ✅ **Cumplimiento:** Control de acceso aplicado correctamente conforme a política. |
| *N/A* | *Cualquiera* | `OK` | **POSIBLE REINGRESO** | ℹ️ **Informativo:** Colaborador activo vigente con antecedente de baja previa. Sin riesgo. |

*Nota de Corte Calendario:* La comparación temporal se realiza normalizada a nivel día (`YYYY-MM-DD`), descartando horas para no penalizar accesos legítimos ocurridos durante la jornada laboral del día de cese.

---

## 📊 Entregables y Salidas (Estructura del Libro Excel Final)

El archivo generado (`Auditoria_Accesos_Resultado.xlsx`) cuenta con **5 pestañas corporativas** formateadas con estilos ejecutivos en OpenPyXL:

1. **Pestaña 1 (`Glosario_y_Criterios` - index=0):**
   - **Referencia ejecutiva y metodológica al abrir el archivo:**
     - **Estructura del Libro:** Alcance y propósito de cada una de las hojas.
     - **Diccionario de Estatus:** Matriz de severidad con condiciones lógicas, impacto y acciones.
     - **Reglas Técnicas:** Justificación del corte calendario, priorización de bajas más recientes y validación compuesta.
2. **Pestaña 2 (`Auditoria_Completa`):**
   - Padrón íntegro con el 100% de los colaboradores evaluados y todas las columnas originales intactas, más las 7 columnas de auditoría:
     `[ULTIMO_LOGIN_DETECTADO, ESTATUS_CUENTA_REPORTE, DISCREPANCIA_IDENTIDAD, ESTATUS_AUDITORIA, CATEGORIA_RIESGO, TIPO_HALLAZGO, DIAS_POST_BAJA]`.
3. **Pestaña 3 (`Riesgos_Activos`):**
   - Casos con cuenta encendida (`Active == 1`) que **requieren intervención inmediata de TI**:
     - `CRÍTICO - RIESGO ACTIVO` (relleno rojo suave `#FEE2E2` / texto `#991B1B`).
     - `ALTO - CUENTA HUÉRFANA` (relleno ámbar suave `#FEF3C7` / texto `#92400E`).
4. **Pestaña 4 (`Incidentes_Pasados`):**
   - Casos clasificados como `MEDIO - INCIDENTE PASADO` (`Active == 0` con login posterior). Formato gris neutro Slate (`#F1F5F9` / `#334155`), ordenados por último login descendente.
5. **Pestaña 5 (`Reingresos`):**
   - Registros clasificados como `POSIBLE REINGRESO` para transparencia de auditoría interna y conciliación con Recursos Humanos.

---

## 📈 Resultados Actuales de Auditoría (Datos Reales de Producción)

Al procesar los archivos de producción (`Universo_Usuarios_PROD_2.xlsx` y `report1790717536569.xlsx`):

| Métrica Ejecutiva | Cantidad | Descripción Operativa |
| :--- | :---: | :--- |
| **Registros Originales Universo** | 4,980 | Padrón histórico original en pestaña `BajasU`. |
| **Bajas Antiguas Depuradas** | 25 | Descarte de bajas de años anteriores de 25 personas con re-baja reciente. |
| **Colaboradores Evaluados** | **4,955** | Universo limpio deduplicado cronológicamente. |
| **Pestaña `Riesgos_Activos` (Acción Urgente)** | **4** | **2 Críticos** (Login + Cuenta Activa) + **2 Altos** (Cuentas Huérfanas). |
| **Pestaña `Incidentes_Pasados` (Remediados)** | **15** | Acceso post-baja con cuenta ya apagada (optimizado tras deduplicación). |
| **Pestaña `Reingresos` y Conformes** | **4,936** | Casos debidamente gestionados (`OK`). |
| **Discrepancias de Identidad** | **18** | Cuentas compartidas / reasignadas protegidas contra falsos positivos. |

---

## 🚀 Modos de Ejecución

### Opción 1: Ejecución Automática (Recomendada)
Coloca los archivos en la carpeta y ejecuta:
```bash
python AUDIT.py --auto
```
o ejecuta interactivamente con detección automática de orígenes:
```bash
python AUDIT.py
```
*(En Windows también puedes hacer doble clic en `AUDIT.bat`)*.

### Opción 2: Especificando Rutas por Argumentos CLI
```bash
python AUDIT.py --universo "Universo_Usuarios_PROD_2.xlsx" --hoja-universo "BajasU" --reporte "report1790717536569.xlsx" --salida "Auditoria_Accesos_Resultado.xlsx"
```

### Argumentos CLI Disponibles:
- `-u`, `--universo`: Ruta al archivo Universo.
- `--hoja-universo`: Nombre de la pestaña a leer en el Universo (por defecto: `BajasU`).
- `-r`, `--reporte`: Ruta al reporte de logins (`.xlsx`, `.xls` o `.csv`).
- `-o`, `--salida`: Ruta para el archivo Excel de resultados (por defecto: `Auditoria_Accesos_Resultado.xlsx`).
- `--auto`: Ejecuta con auto-detección y sin confirmaciones interactivas.

---

## 🧪 Pruebas Automatizadas

La solución cuenta con una suite completa de pruebas unitarias y de integración que validan el cruce, la deduplicación, los casos de borde y los archivos reales del proyecto:

```bash
pytest -v
```

### Cobertura de Pruebas (16/16 Aprobadas):
1. `test_caso_1_mismo_correo_nombres_distintos`: Validación de discrepancia de identidad.
2. `test_caso_2_acentos_y_espacios_match_perfecto`: Normalización fonética y caso de incidente resuelto.
3. `test_caso_3_cuenta_activa_post_baja`: Detección de cuenta huérfana (`Active == 1`).
4. `test_criterios_validacion_matriz_riesgo`: Validación de los 5 estados de la matriz.
5. `test_caso_4_real_workspace_files`: Validación con archivos reales de producción y conteo de filas de salida.
6. `test_caso_5_glosario_y_criterios_completo`: Verificación de secciones, estilos y diccionario en hoja 0.
7. `test_fechas_de_baja_multiples`: Priorización de baja más reciente ante múltiples bajas históricas.
8. `test_login_mas_cercano`: Selección del login más cercano respecto a la fecha de baja efectiva.
9. `test_login_empate_distancia_prioriza_posterior`: Desempate favoreciendo el login posterior ante equidistancia.
10. `test_robust_parse_dates`: Parseo defensivo de formatos de fecha DD/MM/YYYY, ISO y valores vacíos/nulos.
11. `test_parse_active_series`: Parseo robusto de estados de cuenta booleanos y textuales.
12. `test_ghost_columns_and_truncated_headers`: Manejo de columnas fantasma y cabeceras truncadas.
13. `test_multiple_logins_consolidation`: Consolidación y selección de login en reporte.
14. `test_double_risk_audit_matrix`: Integración de la matriz de riesgo.
15. `test_real_workspace_files`: Validación complementaria sobre archivos del repositorio.
16. `test_glosario_y_criterios_unitario`: Creación aislada de pestaña de glosario y metodología.

