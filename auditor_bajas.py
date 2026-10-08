"""
auditor_bajas.py
=============================================================================
Herramienta de Auditoría de Seguridad y Control de Accesos Post-Baja
(Leaver Access Review Automation con Validación Compuesta y Matriz de Riesgo).

Procesa y cruza de manera vectorizada y de alto rendimiento dos orígenes:
1. Archivo Universo: Padrón maestro de empleados y bajas.
   - Pestaña obligatoria: 'BajasU'
   - Columna C (Índice 2): NOMBRE (Nombre completo del colaborador)
   - Columna D (Índice 3): FECHA_BAJA (DD/MM/YYYY con o sin hora o ISO)
   - Columna J (Índice 9): Correo corporativo (a menudo 'Unnamed: 9' o '(No column name)')
2. Archivo Reporte de Logins: Reporte dinámico (report<timestamp>.*)
   - Formatos soportados: Excel (.xlsx / .xls con pestaña 'in' o primera pestaña),
     HTML table export (.xls / .html de Salesforce/CRM) y CSV (.csv).
   - Preprocesamiento defensivo: dropna(how='all', axis=1) para eliminar columnas fantasma.
   - Mapeo resiliente de columnas:
     * Full Name: subcadena 'full name' o 'nombre' (Índice 0 por defecto)
     * Username: subcadena 'username' o 'usuario' (Índice 2 por defecto)
     * Active: subcadena 'active' o 'activo' (Índice 3 por defecto, parseo a 1 o 0)
     * Last Login: subcadena 'last login' o 'ultimo login' (Índice 4 por defecto)
     * Fecha de Baja: subcadena 'fecha de ba' o 'baja' (Índice 5 por defecto)
     * Created Date: subcadena 'created' o 'creac' (Índice 6 por defecto)

VALIDACIÓN COMPUESTA Y PROTECCIÓN CONTRA FALSOS POSITIVOS:
- Clave Compuesta de Cruce: _key_correo + "___" + _key_nombre
- Normalización fonética y ortográfica profunda (eliminación de tildes/acentos,
  diacríticos, espacios redundantes y unificación a minúsculas).
- Detección y marcado de DISCREPANCIA_IDENTIDAD = True cuando un correo coincide
  en el reporte pero pertenece a otra persona (cuentas compartidas o reasignadas).

MATRIZ DE EVALUACIÓN DE RIESGOS Y SEGMENTACIÓN TRIPARTITA:
1. RIESGOS ACTIVOS / INMEDIATOS (Requieren intervención urgente de TI):
   - 'CRÍTICO - RIESGO ACTIVO': Acceso confirmado posterior a la baja (Fecha_Last_Login > Fecha_Baja)
     con cuenta aún habilitada (Active == 1). ESTATUS_AUDITORIA = 'REVISAR'.
   - 'ALTO - CUENTA HUÉRFANA': Cuenta aún habilitada (Active == 1) post-baja pero sin acceso posterior detectado.
     ESTATUS_AUDITORIA = 'REVISAR'.
2. INCIDENTES RESUELTOS / PASADOS (Evidencia histórica para auditoría, no riesgo vivo):
   - 'MEDIO - INCIDENTE PASADO': Acceso confirmado posterior a la baja (Fecha_Last_Login > Fecha_Baja),
     pero la cuenta ya fue deshabilitada (Active == 0). ESTATUS_AUDITORIA = 'INCIDENTE RESUELTO'.
3. REINGRESOS Y CASOS OK (Sin riesgo de seguridad):
   - 'POSIBLE REINGRESO': Usuario activo vigente sin fecha de baja en segundo registro. ESTATUS_AUDITORIA = 'OK'.
   - 'CONFORME': Control de baja aplicado correctamente (Active == 0 sin acceso post-baja) o colaborador activo vigente.
     ESTATUS_AUDITORIA = 'OK'.

ESTRUCTURA DE EXPORTACIÓN EN EXCEL (OPENPYXL):
- Pestaña 1: 'Resumen_Auditoria' (datos cuantitativos de la ejecución: volumen, resultados por categoría e indicadores).
- Pestaña 2: 'Auditoria_Completa' (padrón íntegro con todas las columnas de auditoría).
- Pestaña 3: 'Riesgos_Activos' (casos críticos y huérfanos que requieren acción inmediata de TI).
- Pestaña 4: 'Incidentes_Pasados' (casos medios/remediados ordenados por último login descendente).
- Pestaña 5: 'Reingresos' (registros clasificados como posible reingreso).
=============================================================================
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime
import glob
import json
from html.parser import HTMLParser
import os
from pathlib import Path
import sys
import time
from typing import Any, Dict, List, Optional, Tuple
import unicodedata

import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd


@dataclass
class AuditMetrics:
    """Métricas de resumen y desempeño del proceso de auditoría."""

    total_evaluados: int = 0
    total_revisar: int = 0
    total_ok: int = 0
    total_critico_activo: int = 0
    total_alto_huerfana: int = 0
    total_medio_incidente_pasado: int = 0
    total_conforme: int = 0
    total_posible_reingreso: int = 0
    # Compatibilidad con suite anterior:
    total_ambos: int = 0
    total_solo_acceso: int = 0
    total_solo_cuenta_activa: int = 0
    total_discrepancias_identidad: int = 0
    total_sin_registro: int = 0
    total_con_login_valido: int = 0
    total_sin_fecha_baja: int = 0
    tiempo_total_segundos: float = 0.0
    output_path: str = ""
    # Datos de la ejecución (los completa el pipeline)
    fecha_auditoria: str = ""
    universo_archivo: str = ""
    reporte_archivo: str = ""
    total_universo_original: int = 0
    total_registros_reporte: int = 0


# ============================================================================
# FUNCIONES DE NORMALIZACIÓN Y PARSEO
# ============================================================================


def normalize_text_name(val: Any) -> str:
    """
    Normalización fonética y ortográfica profunda para nombres:
    - Convierte a string.
    - Remueve diacríticos y acentos (FRÍAS -> frias, JOSÉ -> jose, PEÑA -> pena, etc.)
      usando descomposición canónica NFKD y codificación ASCII.
    - Convierte a minúsculas.
    - Colapsa espacios intermedios y recorta extremos (" ".join(texto.split())).
    """
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip().lower()
    s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("utf-8")
    return " ".join(s.split())


def clean_name_series(series: pd.Series | pd.DataFrame) -> pd.Series:
    """Aplica normalize_text_name sobre una serie de nombres."""
    if isinstance(series, pd.DataFrame):
        series = series.iloc[:, 0]
    return series.apply(normalize_text_name)


def clean_identifier_series(series: pd.Series | pd.DataFrame) -> pd.Series:
    """
    Normaliza identificadores de usuario / correos para cruce estricto:
    - Si es DataFrame (por columnas duplicadas), toma la primera columna.
    - Convierte a string
    - Elimina espacios en blanco iniciales/finales (.str.strip())
    - Convierte todo a minúsculas (.str.lower())
    - Convierte valores dummy ('nan', 'none', '', '0', etc.) a None
    """
    if isinstance(series, pd.DataFrame):
        series = series.iloc[:, 0]

    cleaned = (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )
    invalid_mask = cleaned.isin(["", "nan", "none", "null", "nat", "0", "0.0"])
    cleaned[invalid_mask] = None
    return cleaned


def parse_active_series(series: pd.Series | pd.DataFrame) -> pd.Series:
    """
    Normaliza el indicador de cuenta activa ('Active'):
    - Acepta representaciones numéricas o de texto ('1', '0', 1, 0, '1.0', 1.0,
      'true', 'si', 'yes', 'activo').
    - Retorna una pd.Series con enteros 1 o 0 (int64).
    """
    if isinstance(series, pd.DataFrame):
        series = series.iloc[:, 0]

    clean_str = series.fillna("0").astype(str).str.strip().str.lower()
    is_active = clean_str.isin(["1", "1.0", "true", "t", "si", "yes", "y", "activo", "active"])
    return pd.Series(np.where(is_active, 1, 0), index=series.index, dtype="int64")


def robust_parse_dates(series: pd.Series | pd.DataFrame) -> pd.Series:
    """
    Normaliza y parsea fechas asegurando formato día primero (dayfirst=True)
    para fechas latinoamericanas (DD/MM/YYYY) y formato año primero (dayfirst=False)
    para fechas ISO (YYYY-MM-DD).
    Elimina horas y minutos (.dt.normalize()) para comparación estricta de fecha calendario.
    Maneja de forma robusta valores vacíos, nulos, 0, '0', '0.0', timestamps con hora
    y casos donde la serie sea extraída de columnas con nombres duplicados.
    """
    if isinstance(series, pd.DataFrame):
        series = series.iloc[:, 0]

    if pd.api.types.is_datetime64_any_dtype(series):
        return series.dt.normalize()

    clean = series.copy()

    # Si es tipo objeto o string, recortar espacios en blanco
    if clean.dtype == object or str(clean.dtype).startswith("string"):
        is_str = clean.apply(lambda x: isinstance(x, str))
        clean.loc[is_str] = clean.loc[is_str].str.strip()

    # Reemplazar valores 0 o cadenas vacías por None para evitar parseos a época Unix (1970-01-01)
    clean = clean.replace(
        [0, 0.0, "0", "0.0", "", "nan", "None", "NAT", "NaT", "null", "NULL", "none"],
        None,
    )

    str_series = clean.fillna("").astype(str).str.strip()
    is_iso = str_series.str.match(r"^\d{4}-\d{1,2}-\d{1,2}")

    result = pd.Series(pd.NaT, index=series.index)

    if is_iso.any():
        result.loc[is_iso] = pd.to_datetime(
            clean.loc[is_iso],
            errors="coerce",
            dayfirst=False,
            format="mixed",
        )

    non_iso = ~is_iso & clean.notna()
    if non_iso.any():
        result.loc[non_iso] = pd.to_datetime(
            clean.loc[non_iso],
            dayfirst=True,
            errors="coerce",
            format="mixed",
        )

    return pd.to_datetime(result).dt.normalize()


def find_column_by_substring(
    df: pd.DataFrame,
    candidates: list[str],
    default_index: Optional[int] = None,
    column_role: str = "",
    exclude: Optional[list[str]] = None,
) -> Optional[str]:
    """
    Busca de manera resiliente una columna en el DataFrame por coincidencias
    de texto en minúsculas (búsqueda exacta o subcadena).
    Si no encuentra coincidencia y se proporciona un default_index válido, recurre a la posición.
    """
    normalized_cols = {str(col).strip().lower(): col for col in df.columns}
    cols_lower_list = [str(c).strip().lower() for c in df.columns]

    # 1. Búsqueda por coincidencia exacta
    for cand in candidates:
        cand_clean = cand.strip().lower()
        if cand_clean in normalized_cols:
            return normalized_cols[cand_clean]

    # 2. Búsqueda por subcadena
    for cand in candidates:
        cand_clean = cand.strip().lower()
        for idx, col_lower in enumerate(cols_lower_list):
            if cand_clean in col_lower:
                if exclude:
                    if any(ex.strip().lower() in col_lower for ex in exclude):
                        continue
                return df.columns[idx]

    # 3. Recurso al índice posicional por defecto
    if default_index is not None and 0 <= default_index < len(df.columns):
        return df.columns[default_index]

    return None


def find_column_by_name_or_index(
    df: pd.DataFrame,
    candidates: list[str],
    default_index: int,
    column_role: str,
) -> str:
    """Wrapper estricto que lanza ValueError si la columna no puede determinarse."""
    col = find_column_by_substring(df, candidates, default_index, column_role)
    if col is not None:
        return col

    raise ValueError(
        f"No se pudo identificar la columna requerida para '{column_role}'. "
        f"Se buscaron nombres con patrones {candidates} o la columna en posición {default_index}. "
        f"Columnas disponibles: {list(df.columns)}"
    )


# ============================================================================
# LECTURA Y VALIDACIÓN DEFENSIVA DE ENTRADAS
# ============================================================================


class HTMLTableReader(HTMLParser):
    """Parser sin dependencias externas para tablas HTML exportadas como .xls o .html (ej. Salesforce/CRM)."""

    def __init__(self) -> None:
        super().__init__()
        self.rows: list[list[str]] = []
        self.current_row: list[str] = []
        self.current_cell: list[str] = []
        self.in_cell: bool = False

    def handle_starttag(self, tag: str, attrs: list[tuple[str, Optional[str]]]) -> None:
        if tag.lower() in ("th", "td"):
            self.in_cell = True
            self.current_cell = []

    def handle_endtag(self, tag: str) -> None:
        tag_lower = tag.lower()
        if tag_lower in ("th", "td"):
            self.in_cell = False
            self.current_row.append("".join(self.current_cell).strip())
        elif tag_lower == "tr":
            if self.current_row:
                self.rows.append(self.current_row)
            self.current_row = []

    def handle_data(self, data: str) -> None:
        if self.in_cell:
            self.current_cell.append(data)


def parse_html_table_file(file_path: Path | str) -> pd.DataFrame:
    """Lee y extrae datos tabulares desde archivos HTML (incluso con extensión .xls)."""
    path = Path(file_path)
    content = ""
    for enc in ("utf-8", "iso-8859-1", "windows-1252", "latin-1"):
        try:
            with open(path, "r", encoding=enc) as f:
                content = f.read()
            break
        except UnicodeDecodeError:
            continue

    if not content:
        raise ValueError(f"No se pudo decodificar el contenido del archivo '{path.name}'.")

    parser = HTMLTableReader()
    parser.feed(content)
    if not parser.rows:
        raise ValueError(f"No se encontró una tabla válida en el archivo '{path.name}'.")

    header = parser.rows[0]
    data = parser.rows[1:]
    return pd.DataFrame(data, columns=header)


def read_universo(
    file_path: Path | str,
    target_sheet_name: str = "BajasU",
) -> tuple[pd.DataFrame, str, str]:
    """
    Carga el Archivo Universo (Entrada 1).
    Carga específicamente la pestaña obligatoria 'BajasU'.
    Aplica limpieza inicial, valida existencia y número mínimo de columnas (>= 10),
    detecta la columna de FECHA_BAJA (Índice 3 / Col D) y la columna de correo
    corporativo (Índice 9 / Col J / Unnamed: 9).
    """
    path = Path(file_path)
    if not path.is_file():
        raise FileNotFoundError(f"El archivo Universo no existe: '{path.resolve()}'")

    try:
        with pd.ExcelFile(path) as xls:
            sheet_names = xls.sheet_names
            target_sheet = None

            # 1. Búsqueda exacta de la pestaña 'BajasU' (insensible a mayúsculas)
            for s in sheet_names:
                if s.strip().lower() == target_sheet_name.strip().lower():
                    target_sheet = s
                    break

            # 2. Si no se encuentra exacta, buscar si contiene 'bajasu'
            if target_sheet is None:
                for s in sheet_names:
                    if target_sheet_name.strip().lower() in s.strip().lower():
                        target_sheet = s
                        break

            # 3. Si no existe la pestaña requerida, lanzar excepción clara
            if target_sheet is None:
                raise ValueError(
                    f"Validación fallida: El archivo Universo '{path.name}' no contiene la pestaña obligatoria '{target_sheet_name}'. "
                    f"Pestañas encontradas en el archivo: {sheet_names}"
                )

            df = pd.read_excel(xls, sheet_name=target_sheet)
    except ValueError:
        raise
    except Exception as exc:
        raise RuntimeError(f"Error al abrir el archivo Universo '{path.name}': {exc}") from exc

    # Preprocesamiento defensivo: eliminar filas completamente vacías
    df = df.dropna(how="all", axis=0).reset_index(drop=True)

    if df.empty:
        raise ValueError(f"El archivo Universo '{path.name}' está vacío.")

    # Validación defensiva del número mínimo de columnas
    if len(df.columns) < 10:
        raise ValueError(
            f"El archivo Universo '{path.name}' tiene {len(df.columns)} columnas, "
            f"pero se requieren al menos 10 columnas (incluyendo FECHA_BAJA en Col D / Índice 3 "
            f"y correo corporativo en Col J / Índice 9)."
        )

    # Col D (Índice 3): FECHA_BAJA
    baja_col = find_column_by_name_or_index(
        df,
        candidates=["fecha_baja", "fecha baja", "fechabaja", "fec_baja", "baja"],
        default_index=3,
        column_role="FECHA_BAJA (Columna D / Índice 3)",
    )

    # Col J (Índice 9): Correo corporativo (a menudo 'Unnamed: 9' o '(No column name)')
    email_col = find_column_by_name_or_index(
        df,
        candidates=[
            "unnamed: 9",
            "(no column name)",
            "correo",
            "email",
            "mail",
            "correo_corporativo",
            "username",
        ],
        default_index=9,
        column_role="Correo Corporativo (Columna J / Índice 9)",
    )

    return df, baja_col, email_col


def read_reporte_logins(file_path: Path | str) -> tuple[pd.DataFrame, Dict[str, Optional[str]]]:
    """
    Carga el Archivo Reporte de Logins (Entrada 2).
    Soporta:
    - Libros Excel (.xlsx / .xls) con pestaña 'in' o pestaña única/activa.
    - Exportaciones HTML con extensión .xls / .html (típicas de Salesforce/CRM).
    - Archivos CSV (.csv).

    Preprocesamiento defensivo mandatorio:
    - Aplica df.dropna(how='all', axis=1) inmediatamente para eliminar columnas
      fantasma creadas por celdas combinadas.
    - Aplica df.dropna(how='all', axis=0) para eliminar filas vacías.
    - Recorta nombres de columnas.

    Mapeo resiliente de columnas:
    - Full Name: subcadena 'full name' o 'nombre' (Índice 0)
    - Username: subcadena 'username' o 'usuario' (Índice 2)
    - Active: subcadena 'active' o 'activo' (Índice 3)
    - Last Login: subcadena 'last login' o 'ultimo login' (Índice 4)
    - Fecha de Baja (Reporte): subcadena 'fecha de ba' o 'baja' (Índice 5)
    - Created Date: subcadena 'created' o 'creac' (Índice 6)

    Retorna: (df, col_map)
    """
    path = Path(file_path)
    if not path.is_file():
        raise FileNotFoundError(f"El archivo de Reporte de Logins no existe: '{path.resolve()}'")

    df: Optional[pd.DataFrame] = None

    # 1. Comprobar si el archivo es en realidad HTML (muy común en reportes exportados con .xls)
    try:
        with open(path, "rb") as f:
            header_sample = f.read(512).lower()
        if any(tag in header_sample for tag in (b"<table", b"<head", b"<html", b"<!doctype")):
            df = parse_html_table_file(path)
    except Exception:
        df = None

    # 2. Si no es HTML, intentar como CSV si tiene esa extensión
    if df is None and path.suffix.lower() == ".csv":
        try:
            df = pd.read_csv(path)
        except Exception as exc:
            raise RuntimeError(f"Error al leer el archivo CSV '{path.name}': {exc}") from exc

    # 3. Si no es HTML ni CSV, intentar con pd.ExcelFile
    if df is None:
        try:
            with pd.ExcelFile(path) as excel_reader:
                sheet_names = excel_reader.sheet_names
                sheet_target = None

                # Prioridad 1: Pestaña 'in'
                for s in sheet_names:
                    if s.strip().lower() == "in":
                        sheet_target = s
                        break

                # Prioridad 2: Si hay una sola hoja en el libro
                if sheet_target is None and len(sheet_names) == 1:
                    sheet_target = sheet_names[0]

                # Prioridad 3: Hoja que contenga 'report', 'login', 'hoja1' o la primera disponible
                if sheet_target is None:
                    for s in sheet_names:
                        if any(k in s.strip().lower() for k in ["report", "login", "hoja1", "sheet1"]):
                            sheet_target = s
                            break

                if sheet_target is None and sheet_names:
                    sheet_target = sheet_names[0]

                if sheet_target is None:
                    raise ValueError(f"El archivo '{path.name}' no contiene hojas accesibles.")

                df = pd.read_excel(excel_reader, sheet_name=sheet_target)
        except Exception as exc:
            # Fallback final: intentar parsear como HTML si falló ExcelFile
            try:
                df = parse_html_table_file(path)
            except Exception:
                raise RuntimeError(f"Error al leer la estructura de '{path.name}': {exc}") from exc

    if df is None or df.empty:
        raise ValueError(f"El archivo de Reporte de Logins '{path.name}' está vacío o no contiene datos válidos.")

    # =========================================================================
    # PREPROCESAMIENTO DEFENSIVO OBLIGATORIO
    # =========================================================================
    # 1. Eliminar columnas completamente vacías (desfasamiento por celdas combinadas)
    df = df.dropna(how="all", axis=1)

    # 2. Eliminar filas completamente vacías
    df = df.dropna(how="all", axis=0).reset_index(drop=True)

    # 3. Limpieza de nombres de encabezados (quitar espacios sobrantes)
    df.columns = [str(c).strip() for c in df.columns]

    # Validación defensiva del número mínimo de columnas
    if len(df.columns) < 2:
        raise ValueError(
            f"El archivo '{path.name}' tiene {len(df.columns)} columnas después de limpiar vacíos, "
            f"pero se requieren al menos las columnas de usuario y login."
        )

    # =========================================================================
    # MAPEO RESILIENTE DE COLUMNAS POR COINCIDENCIA DE SUBQUERIES
    # =========================================================================
    col_name = find_column_by_substring(
        df,
        candidates=["full name", "full_name", "nombre", "nombre_completo", "name", "empleado"],
        default_index=0 if len(df.columns) > 0 else None,
        column_role="Full Name (Nombre Completo)",
    )

    col_user = find_column_by_name_or_index(
        df,
        candidates=["username", "usuario", "correo", "email", "login", "user name"],
        default_index=2 if len(df.columns) > 2 else 0,
        column_role="Username (Llave de cruce)",
    )

    col_active = find_column_by_substring(
        df,
        candidates=["active", "activo", "estatus", "status"],
        default_index=3 if len(df.columns) > 3 else None,
        column_role="Active (Indicador de cuenta activa)",
    )

    col_login = find_column_by_substring(
        df,
        candidates=["last login", "ultimo login", "last_login", "ultimo_login", "login date", "fecha_login", "login"],
        default_index=4 if len(df.columns) > 4 else None,
        column_role="Last Login (Último inicio de sesión)",
    )

    col_baja_rep = find_column_by_substring(
        df,
        candidates=["fecha de ba", "fecha baja", "fecha_baja", "fechabaja", "baja"],
        default_index=5 if len(df.columns) > 5 else None,
        column_role="Fecha de Baja (Reporte)",
        exclude=["tipo", "motivo"],
    )

    col_created = find_column_by_substring(
        df,
        candidates=["created date", "created", "creacion", "creac", "fecha creacion"],
        default_index=6 if len(df.columns) > 6 else None,
        column_role="Created Date",
    )

    col_map = {
        "name": col_name,
        "user": col_user,
        "active": col_active,
        "login": col_login,
        "baja_rep": col_baja_rep,
        "created": col_created,
    }

    return df, col_map


# ============================================================================
# CONSOLIDACIÓN Y MOTOR DE AUDITORÍA VECTORIZADO CON CLAVE COMPUESTA
# ============================================================================


def select_closest_login_event(
    candidate_logins: list[pd.Timestamp],
    fecha_baja: Optional[pd.Timestamp],
) -> pd.Timestamp:
    """
    Selecciona la fecha de login temporalmente más cercana a fecha_baja:
    distancia = abs(Fecha_Login - Fecha_Baja_Efectiva)

    Criterios de desempate y casos especiales:
    1. Si no hay logins válidos, retorna pd.NaT.
    2. Si solo hay un login válido, retorna ese login.
    3. Si fecha_baja es NaT o None: retorna el login más reciente (máximo).
    4. En caso de empate exacto de distancia (uno antes y uno después),
       prioriza el login posterior para no omitir un posible acceso no autorizado.
    5. Si hay dos fechas idénticas o empate del mismo lado, retorna el más reciente.
    """
    valid_logins = [dt for dt in candidate_logins if dt is not None and not pd.isna(dt)]
    if not valid_logins:
        return pd.NaT
    if len(valid_logins) == 1:
        return valid_logins[0]
    if fecha_baja is None or pd.isna(fecha_baja):
        return max(valid_logins)

    def _login_sort_key(dt: pd.Timestamp):
        dist = abs(dt - fecha_baja)
        # Prioridad de desempate: 0 si login > fecha_baja else 1 (priorizar posterior)
        tie_priority = 0 if dt > fecha_baja else 1
        return (dist, tie_priority, -dt.value)

    return min(valid_logins, key=_login_sort_key)


def deduplicate_universo(
    df_universo: pd.DataFrame,
    baja_col: str,
    email_col: str,
    nombre_col: Optional[str] = None,
) -> tuple[pd.DataFrame, set[str]]:
    """
    Implementa el algoritmo estricto de deduplicación y resolución cronológica en Universo:
    Al agrupar por _key_compuesta (_clean_email + '___' + _clean_nombre):

    * Paso 1 (Detección de Reingreso):
      Si la persona tiene registros donde Fecha_Baja es nula/vacía Y registros con Fecha_Baja válida,
      se identifica como Reingreso Activo (ESTATUS_AUDITORIA = 'OK', TIPO_HALLAZGO = 'POSIBLE REINGRESO').
      Se registran las claves en reingreso_keys y se conservan los registros para su trazabilidad
      en auditoría, garantizando que el usuario no sea penalizado por accesos de su historial previo.

    * Paso 2 (Múltiples Bajas):
      Si todos los registros de la persona tienen fecha de baja válida:
      - Se ordenan cronológicamente por Fecha_Baja descendente (más reciente primero).
      - Se conserva únicamente el registro con la fecha de baja más reciente (máxima).
      - Se descartan las fechas de baja antiguas (evita falsos positivos por bajas históricas superadas).

    Retorna:
    - df_deduplicado: DataFrame de Universo deduplicado.
    - reingreso_keys: Conjunto de _key_compuesta identificadas como Reingreso Activo.
    """
    df_work = df_universo.copy()

    if nombre_col is None or nombre_col not in df_work.columns:
        nombre_col = find_column_by_substring(
            df_work,
            candidates=["nombre", "nombre_completo", "empleado", "trabajador", "full name", "name"],
            default_index=2 if len(df_work.columns) > 2 else 0,
            column_role="Nombre Completo (Columna C / Índice 2)",
        )

    clean_email = clean_identifier_series(df_work[email_col])
    if nombre_col and nombre_col in df_work.columns:
        clean_name = clean_name_series(df_work[nombre_col])
    else:
        clean_name = pd.Series("", index=df_work.index)

    parsed_baja = robust_parse_dates(df_work[baja_col])

    df_work["_clean_email"] = clean_email
    df_work["_clean_name"] = clean_name
    df_work["_parsed_baja"] = parsed_baja
    df_work["_key_compuesta"] = np.where(
        clean_email.notna(),
        clean_email + "___" + clean_name,
        None,
    )

    indices_to_keep: list[int] = []
    reingreso_keys: set[str] = set()

    valid_key_mask = df_work["_key_compuesta"].notna()
    if valid_key_mask.any():
        for key, grp in df_work[valid_key_mask].groupby("_key_compuesta", sort=False):
            if len(grp) == 1:
                indices_to_keep.extend(grp.index.tolist())
            else:
                has_na = grp["_parsed_baja"].isna().any()
                has_valid = grp["_parsed_baja"].notna().any()

                if has_na and has_valid:
                    # Paso 1: Reingreso activo
                    reingreso_keys.add(key)
                    indices_to_keep.extend(grp.index.tolist())
                elif grp["_parsed_baja"].notna().all():
                    # Paso 2: Múltiples bajas -> conservar únicamente la más reciente (máxima)
                    sorted_grp = grp.sort_values(by="_parsed_baja", ascending=False)
                    indices_to_keep.append(sorted_grp.index[0])
                else:
                    indices_to_keep.append(grp.index[0])

    no_key_indices = df_work[~valid_key_mask].index.tolist()
    indices_to_keep.extend(no_key_indices)

    indices_to_keep.sort()
    df_deduplicado = df_universo.loc[indices_to_keep].copy().reset_index(drop=True)
    return df_deduplicado, reingreso_keys


def consolidate_logins(
    df_reporte: pd.DataFrame,
    col_map: Dict[str, Optional[str]],
    fecha_baja_map: Optional[Dict[str, pd.Timestamp]] = None,
) -> pd.DataFrame:
    """
    Normaliza identificadores, nombres, fechas y estatus de cuentas del reporte de logins.
    Genera la clave compuesta (_key_compuesta = _clean_username + '___' + _clean_name).
    Deduplicación inteligente con selección cronológica:
    - Si se proporciona fecha_baja_map, selecciona la fecha de login más cercana
      a la fecha de baja efectiva (priorizando la posterior en caso de empate exacto).
    - Si no se proporciona o no hay fecha de baja, conserva el login más reciente.
    - Conserva Active = 1 si la cuenta figura activa en cualquiera de los registros (max).
    - Conserva la Fecha de Baja reportada más reciente (max) para respaldo.
    - Almacena _candidate_logins con todos los logins válidos del usuario.
    """
    working_df = df_reporte.copy()
    user_col = col_map["user"]
    name_col = col_map.get("name")
    login_col = col_map.get("login")
    active_col = col_map.get("active")
    baja_rep_col = col_map.get("baja_rep")

    working_df["_clean_username"] = clean_identifier_series(working_df[user_col])

    if name_col and name_col in working_df.columns:
        working_df["_clean_name"] = clean_name_series(working_df[name_col])
    else:
        working_df["_clean_name"] = ""

    # Generación de Llave Compuesta: correo + "___" + nombre normalizado
    working_df["_key_compuesta"] = np.where(
        working_df["_clean_username"].notna(),
        working_df["_clean_username"] + "___" + working_df["_clean_name"],
        None,
    )

    if login_col and login_col in working_df.columns:
        working_df["_parsed_login_date"] = robust_parse_dates(working_df[login_col])
    else:
        working_df["_parsed_login_date"] = pd.NaT

    if active_col and active_col in working_df.columns:
        working_df["_parsed_active"] = parse_active_series(working_df[active_col])
    else:
        working_df["_parsed_active"] = 0

    if baja_rep_col and baja_rep_col in working_df.columns:
        working_df["_parsed_baja_rep"] = robust_parse_dates(working_df[baja_rep_col])
    else:
        working_df["_parsed_baja_rep"] = pd.NaT

    # Filtrar registros sin llave válida
    valid_entries = working_df[working_df["_key_compuesta"].notna()].copy()

    records = []
    for key, grp in valid_entries.groupby("_key_compuesta", sort=False):
        clean_user = grp["_clean_username"].iloc[0]
        clean_name = grp["_clean_name"].iloc[0]
        max_active = int(grp["_parsed_active"].max())

        valid_bajas_rep = grp["_parsed_baja_rep"].dropna()
        max_baja_rep = valid_bajas_rep.max() if not valid_bajas_rep.empty else pd.NaT

        valid_logins = [d for d in grp["_parsed_login_date"] if pd.notna(d)]

        target_baja = None
        if fecha_baja_map and key in fecha_baja_map:
            target_baja = fecha_baja_map[key]

        chosen_login = select_closest_login_event(valid_logins, target_baja)

        records.append({
            "_key_compuesta": key,
            "_clean_username": clean_user,
            "_clean_name": clean_name,
            "_parsed_login_date": chosen_login,
            "_parsed_active": max_active,
            "_parsed_baja_rep": max_baja_rep,
            "_candidate_logins": valid_logins,
            "_en_reporte": True,
        })

    consolidated = pd.DataFrame(records)
    if consolidated.empty:
        consolidated = pd.DataFrame(columns=[
            "_key_compuesta", "_clean_username", "_clean_name",
            "_parsed_login_date", "_parsed_active", "_parsed_baja_rep",
            "_candidate_logins", "_en_reporte"
        ])
    return consolidated


def execute_audit(
    df_universo: pd.DataFrame,
    baja_col: str,
    email_col: str,
    consolidated_logins: pd.DataFrame,
    nombre_col: Optional[str] = None,
) -> tuple[pd.DataFrame, pd.DataFrame, AuditMetrics]:
    """
    Ejecuta el cruce vectorizado con VALIDACIÓN COMPUESTA (Correo + Nombre Completo)
    y MATRIZ DE RIESGO TRIPARTITA:
    
    1. Generación de clave compuesta:
       _key_compuesta = _clean_email + "___" + _clean_nombre
       Garantiza que cuentas reusadas o compartidas no generen falsos positivos
       al comparar estrictamente la identidad del colaborador.

    2. Detección de Discrepancias de Identidad:
       Si un correo existe en el reporte pero bajo un nombre distinto:
       - No vincula arbitrariamente el login ni el estatus.
       - Marca DISCREPANCIA_IDENTIDAD = True para revisión manual.

    3. Detección de Posibles Reingresos (Transparencia de Auditoría):
       Un colaborador se identifica como reingreso vigente cuando para una misma
       identidad (correo, nombre o clave de trabajador) existe un registro con baja
       y al menos otro registro activo vigente SIN fecha de baja.
       -> ESTATUS_AUDITORIA: 'OK'
       -> CATEGORIA_RIESGO: 'POSIBLE REINGRESO'

    4. Matriz de Clasificación de Riesgo:
       Para registros con fecha de baja válida que no correspondan a un reingreso vigente:
       - Acceso Post-Baja (SÍ) & Active == 1:
         -> ESTATUS_AUDITORIA: 'REVISAR'
         -> CATEGORIA_RIESGO: 'CRÍTICO - RIESGO ACTIVO'
       - Acceso Post-Baja (NO) & Active == 1:
         -> ESTATUS_AUDITORIA: 'REVISAR'
         -> CATEGORIA_RIESGO: 'ALTO - CUENTA HUÉRFANA'
       - Acceso Post-Baja (SÍ) & Active == 0:
         -> ESTATUS_AUDITORIA: 'INCIDENTE RESUELTO'
         -> CATEGORIA_RIESGO: 'MEDIO - INCIDENTE PASADO'
       - Acceso Post-Baja (NO) & Active == 0:
         -> ESTATUS_AUDITORIA: 'OK'
         -> CATEGORIA_RIESGO: 'CONFORME'

    Retorna:
    - df_completo: Universo original íntegro + columnas calculadas:
      [ULTIMO_LOGIN_DETECTADO, ESTATUS_CUENTA_REPORTE, DISCREPANCIA_IDENTIDAD, ESTATUS_AUDITORIA, CATEGORIA_RIESGO, TIPO_HALLAZGO, DIAS_POST_BAJA]
    - df_riesgos_activos: Únicamente los registros con estatus REVISAR (CRÍTICO y ALTO).
    - metrics: Métricas consolidadas.
    """
    # 0. Deduplicación cronológica del Universo (Paso 1 y Paso 2)
    uni, reingreso_keys_dedup = deduplicate_universo(
        df_universo=df_universo,
        baja_col=baja_col,
        email_col=email_col,
        nombre_col=nombre_col,
    )

    # Detección de columna de Nombre si no fue provista
    if nombre_col is None or nombre_col not in uni.columns:
        nombre_col = find_column_by_substring(
            uni,
            candidates=["nombre", "nombre_completo", "empleado", "trabajador", "full name", "name"],
            default_index=2 if len(uni.columns) > 2 else 0,
            column_role="Nombre Completo (Columna C / Índice 2)",
        )

    # Normalización de correo y nombre completo
    uni["_clean_email"] = clean_identifier_series(uni[email_col])
    has_report_names = (
        "_clean_name" in consolidated_logins.columns
        and (consolidated_logins["_clean_name"].astype(str).str.len() > 0).any()
    )
    if has_report_names and nombre_col:
        uni["_clean_name"] = clean_name_series(uni[nombre_col])
    else:
        uni["_clean_name"] = ""
    uni["_parsed_baja_uni"] = robust_parse_dates(uni[baja_col])

    # Generación de la Llave Compuesta en Universo
    uni["_key_compuesta"] = np.where(
        uni["_clean_email"].notna(),
        uni["_clean_email"] + "___" + uni["_clean_name"],
        None,
    )

    # Cruce vectorizado LEFT JOIN sobre la Llave Compuesta
    cols_to_merge = [
        c for c in consolidated_logins.columns
        if c not in ("_clean_name", "_clean_username") or c == "_key_compuesta"
    ]
    merged = pd.merge(
        uni,
        consolidated_logins[cols_to_merge],
        on="_key_compuesta",
        how="left",
    )

    # 1. Prioridad de Fecha de Baja: Universo primero, combine_first con reporte
    if "_parsed_baja_rep" in merged.columns:
        fecha_baja_final = merged["_parsed_baja_uni"].combine_first(merged["_parsed_baja_rep"])
    else:
        fecha_baja_final = merged["_parsed_baja_uni"]

    # Refinamiento cronológico de login más cercano respecto a fecha_baja_final
    if "_candidate_logins" in merged.columns:
        refined_logins = []
        for idx, row in merged.iterrows():
            cands = row["_candidate_logins"]
            fb = fecha_baja_final.loc[idx]
            if isinstance(cands, list) and len(cands) > 1 and pd.notna(fb):
                refined_logins.append(select_closest_login_event(cands, fb))
            else:
                refined_logins.append(row["_parsed_login_date"])
        merged["_parsed_login_date"] = pd.Series(refined_logins, index=merged.index)

    # Identificación de todos los correos presentes en el reporte
    report_emails = set(consolidated_logins["_clean_username"].dropna().unique())

    has_email = merged["_clean_email"].notna()
    en_reporte = has_email & merged["_en_reporte"].fillna(False).astype(bool)
    has_valid_login = en_reporte & merged["_parsed_login_date"].notna()
    has_valid_baja = fecha_baja_final.notna()
    is_active_account = en_reporte & (merged["_parsed_active"].fillna(0).astype(int) == 1)

    # Detección de Discrepancia de Identidad:
    # El correo existe en el reporte, pero el nombre no correspondió con la clave compuesta
    discrepancia_identidad = ~en_reporte & merged["_clean_email"].isin(report_emails)

    # -------------------------------------------------------------------------
    # Detección de Posibles Reingresos (Usuario activo vigente en 2do registro)
    # -------------------------------------------------------------------------
    reingreso_emails = set()
    if has_email.any():
        for email_val, group in merged[has_email].groupby("_clean_email"):
            bajas_in_group = group["_parsed_baja_uni"].combine_first(
                group["_parsed_baja_rep"] if "_parsed_baja_rep" in group.columns else pd.Series(pd.NaT, index=group.index)
            )
            has_baja_grp = bajas_in_group.notna()
            no_baja_grp = bajas_in_group.isna()
            if len(group) > 1 and has_baja_grp.any() and no_baja_grp.any():
                reingreso_emails.add(email_val)

    reingreso_names = set()
    if "_clean_name" in merged.columns and (merged["_clean_name"].astype(str).str.len() > 0).any():
        valid_names = merged["_clean_name"].astype(str).str.len() > 0
        for name_val, group in merged[valid_names].groupby("_clean_name"):
            bajas_in_group = group["_parsed_baja_uni"].combine_first(
                group["_parsed_baja_rep"] if "_parsed_baja_rep" in group.columns else pd.Series(pd.NaT, index=group.index)
            )
            has_baja_grp = bajas_in_group.notna()
            no_baja_grp = bajas_in_group.isna()
            if len(group) > 1 and has_baja_grp.any() and no_baja_grp.any():
                reingreso_names.add(name_val)

    trab_col = find_column_by_substring(uni, candidates=["cla_trab", "id_empleado", "num_emp", "empleado_id"])
    reingreso_trab = set()
    if trab_col and trab_col in merged.columns:
        valid_trab = merged[trab_col].notna() & ~merged[trab_col].astype(str).str.strip().isin(["", "0", "0.0", "nan"])
        for trab_val, group in merged[valid_trab].groupby(trab_col):
            bajas_in_group = group["_parsed_baja_uni"].combine_first(
                group["_parsed_baja_rep"] if "_parsed_baja_rep" in group.columns else pd.Series(pd.NaT, index=group.index)
            )
            has_baja_grp = bajas_in_group.notna()
            no_baja_grp = bajas_in_group.isna()
            if len(group) > 1 and has_baja_grp.any() and no_baja_grp.any():
                reingreso_trab.add(trab_val)

    is_reingreso = (
        (merged["_clean_email"].isin(reingreso_emails))
        | (merged["_clean_name"].isin(reingreso_names))
        | (merged[trab_col].isin(reingreso_trab) if trab_col and trab_col in merged.columns else pd.Series(False, index=merged.index))
        | (merged["_key_compuesta"].isin(reingreso_keys_dedup))
    )

    # -------------------------------------------------------------------------
    # Reglas de Auditoría y Matriz de Riesgo Tripartita
    # -------------------------------------------------------------------------
    acceso_post_baja = has_valid_baja & has_valid_login & (merged["_parsed_login_date"] > fecha_baja_final) & ~is_reingreso

    cond_reingreso = is_reingreso
    cond_critico = ~is_reingreso & has_valid_baja & acceso_post_baja & is_active_account
    cond_alto = ~is_reingreso & has_valid_baja & ~acceso_post_baja & is_active_account
    cond_medio = ~is_reingreso & has_valid_baja & acceso_post_baja & ~is_active_account
    cond_conforme = ~is_reingreso & (~has_valid_baja | (~acceso_post_baja & ~is_active_account))

    # ESTATUS_AUDITORIA: 'REVISAR', 'INCIDENTE RESUELTO', 'OK'
    estatus_conditions = [
        cond_critico,
        cond_alto,
        cond_medio,
        cond_reingreso,
    ]
    estatus_choices = [
        "REVISAR",
        "REVISAR",
        "INCIDENTE RESUELTO",
        "OK",
    ]
    estatus_series = np.select(estatus_conditions, estatus_choices, default="OK")

    # CATEGORIA_RIESGO:
    # - 'CRÍTICO - RIESGO ACTIVO'
    # - 'ALTO - CUENTA HUÉRFANA'
    # - 'MEDIO - INCIDENTE PASADO'
    # - 'POSIBLE REINGRESO'
    # - 'CONFORME'
    riesgo_conditions = [
        cond_critico,
        cond_alto,
        cond_medio,
        cond_reingreso,
    ]
    riesgo_choices = [
        "CRÍTICO - RIESGO ACTIVO",
        "ALTO - CUENTA HUÉRFANA",
        "MEDIO - INCIDENTE PASADO",
        "POSIBLE REINGRESO",
    ]
    categoria_riesgo_series = np.select(riesgo_conditions, riesgo_choices, default="CONFORME")

    # DIAS_POST_BAJA
    dias_diff = (merged["_parsed_login_date"] - fecha_baja_final).dt.days
    dias_post_baja = pd.Series(
        np.where(acceso_post_baja, dias_diff, pd.NA),
        dtype="Int64",
        index=merged.index,
    )

    # ULTIMO_LOGIN_DETECTADO
    formatted_logins = merged["_parsed_login_date"].dt.strftime("%d/%m/%Y").fillna("")
    ultimo_login_series = np.where(
        ~en_reporte,
        "Sin registro",
        np.where(has_valid_login, formatted_logins, ""),
    )

    # ESTATUS_CUENTA_REPORTE
    estatus_cuenta_series = np.where(
        ~en_reporte,
        "Sin registro",
        np.where(is_active_account, "Activa", "Inactiva"),
    )

    # Construcción de DataFrame Completo (conserva todas las columnas originales intactas)
    df_completo = uni.copy()

    # Normalizar FECHA_BAJA visualmente a DD/MM/YYYY donde exista fecha válida
    if isinstance(df_completo[baja_col], pd.DataFrame):
        target_baja_col = df_completo.columns[3]
    else:
        target_baja_col = baja_col

    valid_baja_mask = fecha_baja_final.notna()
    df_completo[target_baja_col] = df_completo[target_baja_col].astype(object)
    df_completo.loc[valid_baja_mask, target_baja_col] = fecha_baja_final.loc[valid_baja_mask].dt.strftime("%d/%m/%Y")

    # Columnas finales ordenadas y limpias requeridas:
    df_completo["ULTIMO_LOGIN_DETECTADO"] = ultimo_login_series
    df_completo["ESTATUS_CUENTA_REPORTE"] = estatus_cuenta_series
    df_completo["DISCREPANCIA_IDENTIDAD"] = discrepancia_identidad
    df_completo["ESTATUS_AUDITORIA"] = estatus_series
    df_completo["CATEGORIA_RIESGO"] = categoria_riesgo_series
    df_completo["TIPO_HALLAZGO"] = categoria_riesgo_series
    df_completo["DIAS_POST_BAJA"] = dias_post_baja

    # -------------------------------------------------------------------------
    # PESTAÑA 2: Riesgos_Activos (Requieren intervención urgente de TI: Active == 1)
    # -------------------------------------------------------------------------
    mask_riesgos_activos = df_completo["CATEGORIA_RIESGO"].isin([
        "CRÍTICO - RIESGO ACTIVO",
        "ALTO - CUENTA HUÉRFANA",
    ])
    df_riesgos_activos = df_completo[mask_riesgos_activos].copy()

    if not df_riesgos_activos.empty:
        df_riesgos_activos["_sort_priority"] = np.where(
            df_riesgos_activos["CATEGORIA_RIESGO"] == "CRÍTICO - RIESGO ACTIVO",
            1,
            2,
        )
        df_riesgos_activos["_sort_dias"] = df_riesgos_activos["DIAS_POST_BAJA"].fillna(-1)

        df_riesgos_activos = (
            df_riesgos_activos.sort_values(
                by=["_sort_priority", "_sort_dias"],
                ascending=[True, False],
            )
            .drop(columns=["_sort_priority", "_sort_dias"])
            .reset_index(drop=True)
        )

    # -------------------------------------------------------------------------
    # PESTAÑA 3: Incidentes_Pasados (Acceso post-baja con cuenta ya desactivada: Active == 0)
    # -------------------------------------------------------------------------
    mask_incidentes_pasados = df_completo["CATEGORIA_RIESGO"] == "MEDIO - INCIDENTE PASADO"
    df_incidentes_pasados = df_completo[mask_incidentes_pasados].copy()

    if not df_incidentes_pasados.empty:
        df_incidentes_pasados["_sort_login"] = merged.loc[mask_incidentes_pasados, "_parsed_login_date"]
        df_incidentes_pasados = (
            df_incidentes_pasados.sort_values(by="_sort_login", ascending=False)
            .drop(columns=["_sort_login"])
            .reset_index(drop=True)
        )

    # -------------------------------------------------------------------------
    # PESTAÑA 4: Reingresos (POSIBLE REINGRESO)
    # -------------------------------------------------------------------------
    mask_reingresos = df_completo["CATEGORIA_RIESGO"] == "POSIBLE REINGRESO"
    df_reingresos = df_completo[mask_reingresos].copy()
    if not df_reingresos.empty:
        df_reingresos = df_reingresos.reset_index(drop=True)

    # Métricas consolidadas
    metrics = AuditMetrics(
        total_evaluados=len(df_completo),
        total_revisar=int(mask_riesgos_activos.sum()),
        total_ok=int((df_completo["ESTATUS_AUDITORIA"] == "OK").sum()),
        total_critico_activo=int(cond_critico.sum()),
        total_alto_huerfana=int(cond_alto.sum()),
        total_medio_incidente_pasado=int(cond_medio.sum()),
        total_conforme=int(cond_conforme.sum()),
        total_posible_reingreso=int(cond_reingreso.sum()),
        total_ambos=int(cond_critico.sum()),
        total_solo_acceso=int(cond_medio.sum()),
        total_solo_cuenta_activa=int(cond_alto.sum()),
        total_discrepancias_identidad=int(discrepancia_identidad.sum()),
        total_sin_registro=int((~en_reporte).sum()),
        total_con_login_valido=int(has_valid_login.sum()),
        total_sin_fecha_baja=int((~has_valid_baja).sum()),
    )

    return df_completo, df_riesgos_activos, metrics


# ============================================================================
# EXPORTACIÓN Y FORMATO PROFESIONAL EN EXCEL (OPENPYXL)
# ============================================================================


def _write_merged_cell(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    cell_range: str,
    value: Any,
    font: Optional[Font] = None,
    fill: Optional[PatternFill] = None,
    border: Optional[Border] = None,
    alignment: Optional[Alignment] = None,
) -> None:
    """Combina el rango de celdas indicado y aplica estilo uniforme a todas sus celdas."""
    ws.merge_cells(cell_range)
    cells = ws[cell_range]
    cells[0][0].value = value
    for row in cells:
        for cell in row:
            if font is not None:
                cell.font = font
            if fill is not None:
                cell.fill = fill
            if border is not None:
                cell.border = border
            if alignment is not None:
                cell.alignment = alignment


def _write_cell(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    row: int,
    col: int,
    value: Any,
    font: Optional[Font] = None,
    fill: Optional[PatternFill] = None,
    border: Optional[Border] = None,
    alignment: Optional[Alignment] = None,
) -> None:
    """Escribe un valor y aplica estilos a una celda individual."""
    cell = ws.cell(row=row, column=col, value=value)
    if font is not None:
        cell.font = font
    if fill is not None:
        cell.fill = fill
    if border is not None:
        cell.border = border
    if alignment is not None:
        cell.alignment = alignment


RESUMEN_SHEET = "Resumen_Auditoria"

# (categoría, estatus, criterio breve, relleno, texto)
_RESUMEN_CATEGORIAS = [
    ("CRÍTICO - RIESGO ACTIVO", "REVISAR", "Login > baja y cuenta activa", "FFFEE2E2", "FF991B1B"),
    ("ALTO - CUENTA HUÉRFANA", "REVISAR", "Sin login > baja y cuenta activa", "FFFEF3C7", "FF92400E"),
    ("MEDIO - INCIDENTE PASADO", "INCIDENTE RESUELTO", "Login > baja y cuenta inactiva", "FFF1F5F9", "FF334155"),
    ("POSIBLE REINGRESO", "OK", "Registros con y sin baja", "FFE0F2FE", "FF0369A1"),
    ("CONFORME", "OK", "Sin login > baja y cuenta inactiva", "FFDCFCE7", "FF166534"),
]


def create_resumen_sheet(
    workbook: openpyxl.Workbook,
    df_completo: pd.DataFrame,
    metrics: Optional[AuditMetrics] = None,
) -> openpyxl.worksheet.worksheet.Worksheet:
    """
    Crea como primera pestaña (index=0) la hoja 'Resumen_Auditoria' con datos
    cuantitativos de la auditoría ejecutada:
    1. Datos de la ejecución (fecha y archivos de entrada).
    2. Volumen procesado (registros del Universo, deduplicación y reporte).
    3. Resultados por categoría de riesgo (cantidad y porcentaje).
    4. Indicadores adicionales (discrepancias, sin registro, días post-baja).
    Las filas que dependen de `metrics` solo se escriben si se proporcionó.
    """
    if RESUMEN_SHEET in workbook.sheetnames:
        del workbook[RESUMEN_SHEET]
    ws = workbook.create_sheet(title=RESUMEN_SHEET, index=0)
    ws.views.sheetView[0].showGridLines = True

    font_name = "Segoe UI"
    title_font = Font(name=font_name, size=14, bold=True, color="FFFFFFFF")
    title_fill = PatternFill(start_color="FF1F2937", end_color="FF1F2937", fill_type="solid")
    head_font = Font(name=font_name, size=10, bold=True, color="FFFFFFFF")
    head_fill = PatternFill(start_color="FF374151", end_color="FF374151", fill_type="solid")
    body_font = Font(name=font_name, size=10, color="FF1F2937")
    bold_font = Font(name=font_name, size=10, bold=True, color="FF1F2937")
    side = Side(border_style="thin", color="FFE5E7EB")
    border = Border(left=side, right=side, top=side, bottom=side)
    left = Alignment(horizontal="left", vertical="center")
    center = Alignment(horizontal="center", vertical="center")

    total = len(df_completo)
    fecha = (metrics.fecha_auditoria if metrics and metrics.fecha_auditoria else datetime.now().strftime("%d/%m/%Y"))

    _write_merged_cell(ws, "A1:D1", f"RESUMEN DE AUDITORÍA - {fecha}", title_font, title_fill, border, left)
    ws.row_dimensions[1].height = 28
    row = 3

    def section(title: str, headers: list[str]) -> None:
        nonlocal row
        for i, h in enumerate(headers, start=1):
            _write_cell(ws, row, i, h if i > 1 else title, head_font, head_fill, border, left if i == 1 else center)
        row += 1

    def line(label: str, *values: Any, bold: bool = False, fill: Optional[PatternFill] = None, fonts: Optional[Font] = None) -> None:
        nonlocal row
        _write_cell(ws, row, 1, label, fonts or (bold_font if bold else body_font), fill, border, left)
        for i, v in enumerate(values, start=2):
            _write_cell(ws, row, i, v, fonts or (bold_font if bold else body_font), fill, border, center)
        row += 1

    # 1. Volumen procesado
    section("Volumen procesado", ["Concepto", "Cantidad"])
    if metrics is not None:
        if metrics.universo_archivo:
            line("Archivo Universo", metrics.universo_archivo)
        if metrics.reporte_archivo:
            line("Archivo Reporte de Logins", metrics.reporte_archivo)
        if metrics.total_universo_original:
            line("Registros originales del Universo", metrics.total_universo_original)
            line("Bajas antiguas depuradas", metrics.total_universo_original - total)
        if metrics.total_registros_reporte:
            line("Registros en el reporte de logins", metrics.total_registros_reporte)
    line("Colaboradores evaluados", total, bold=True)
    row += 1

    # 2. Resultados por categoría
    section("Resultados por categoría", ["Categoría", "Cantidad", "% del total", "Criterio"])
    counts = df_completo["CATEGORIA_RIESGO"].value_counts() if total else pd.Series(dtype=int)
    for categoria, estatus, criterio, bg, fg in _RESUMEN_CATEGORIAS:
        n = int(counts.get(categoria, 0))
        pct = (n / total) if total else 0
        _write_cell(
            ws, row, 1, categoria,
            Font(name=font_name, size=10, bold=True, color=fg),
            PatternFill(start_color=bg, end_color=bg, fill_type="solid"), border, left,
        )
        _write_cell(ws, row, 2, n, bold_font, None, border, center)
        c = ws.cell(row=row, column=3, value=pct)
        c.number_format = "0.0%"
        c.font, c.border, c.alignment = body_font, border, center
        _write_cell(ws, row, 4, criterio, body_font, None, border, left)
        row += 1
    _write_cell(ws, row, 1, "TOTAL", bold_font, None, border, left)
    _write_cell(ws, row, 2, total, bold_font, None, border, center)
    c = ws.cell(row=row, column=3, value=1 if total else 0)
    c.number_format = "0.0%"
    c.font, c.border, c.alignment = bold_font, border, center
    _write_cell(ws, row, 4, "", body_font, None, border, left)
    row += 2

    # 3. Indicadores adicionales
    section("Indicadores adicionales", ["Indicador", "Valor"])
    n_revisar = int(df_completo["ESTATUS_AUDITORIA"].eq("REVISAR").sum()) if total else 0
    line("Casos a revisar (críticos + huérfanas)", n_revisar, bold=True)
    line("Discrepancias de identidad", int(df_completo["DISCREPANCIA_IDENTIDAD"].fillna(False).astype(bool).sum()) if total else 0)
    line("Sin registro en el reporte", int(df_completo["ESTATUS_CUENTA_REPORTE"].eq("Sin registro").sum()) if total else 0)
    line("Cuentas activas en el reporte", int(df_completo["ESTATUS_CUENTA_REPORTE"].eq("Activa").sum()) if total else 0)
    if metrics is not None:
        line("Con login válido", metrics.total_con_login_valido)
        line("Sin fecha de baja válida", metrics.total_sin_fecha_baja)
    dias = pd.to_numeric(df_completo["DIAS_POST_BAJA"], errors="coerce").dropna() if total else pd.Series(dtype=float)
    line("Accesos post-baja detectados", int(len(dias)))
    if len(dias):
        line("Días post-baja (máximo)", int(dias.max()))
        line("Días post-baja (promedio)", round(float(dias.mean()), 1))

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 36
    return ws


def apply_professional_excel_styling(workbook: openpyxl.Workbook) -> None:
    """
    Aplica una estética corporativa de alta calidad a las hojas de datos del libro Excel:
    - Encabezados oscuros (#1F2937) con tipografía blanca negrita
    - Bordes limpios y discretos (#E5E7EB)
    - Inmovilización de la fila superior (encabezados siempre visibles)
    - Auto-ajuste dinámico del ancho de columnas
    - Resaltado condicional semántico:
      * CRÍTICO - RIESGO ACTIVO / REVISAR: Rojo suave (#FEE2E2 / #991B1B)
      * ALTO - CUENTA HUÉRFANA: Ámbar / Naranja suave (#FEF3C7 / #92400E)
      * MEDIO - INCIDENTE PASADO / INCIDENTE RESUELTO: Gris / Azul neutro Slate (#F1F5F9 / #334155)
      * POSIBLE REINGRESO: Azul suave / Cian (#E0F2FE / #0369A1)
      * CONFORME / OK: Verde suave (#DCFCE7 / #166534)
    - Líneas de cuadrícula activas y formato estándar de fechas
    """
    header_fill = PatternFill(start_color="FF1F2937", end_color="FF1F2937", fill_type="solid")
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFFFF")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    body_font = Font(name="Segoe UI", size=9, bold=False, color="FF1F2937")
    regular_alignment = Alignment(horizontal="left", vertical="center")
    center_alignment = Alignment(horizontal="center", vertical="center")

    thin_border_side = Side(border_style="thin", color="FFE5E7EB")
    cell_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side,
    )

    # Estilos de resaltado
    rojo_fill = PatternFill(start_color="FFFEE2E2", end_color="FFFEE2E2", fill_type="solid")
    rojo_font = Font(name="Segoe UI", size=9, bold=True, color="FF991B1B")

    ambar_fill = PatternFill(start_color="FFFEF3C7", end_color="FFFEF3C7", fill_type="solid")
    ambar_font = Font(name="Segoe UI", size=9, bold=True, color="FF92400E")

    slate_fill = PatternFill(start_color="FFF1F5F9", end_color="FFF1F5F9", fill_type="solid")
    slate_font = Font(name="Segoe UI", size=9, bold=True, color="FF334155")

    azul_fill = PatternFill(start_color="FFE0F2FE", end_color="FFE0F2FE", fill_type="solid")
    azul_font = Font(name="Segoe UI", size=9, bold=True, color="FF0369A1")

    verde_fill = PatternFill(start_color="FFDCFCE7", end_color="FFDCFCE7", fill_type="solid")
    verde_font = Font(name="Segoe UI", size=9, bold=True, color="FF166534")

    sin_registro_font = Font(name="Segoe UI", size=9, italic=True, color="FF6B7280")

    for sheetname in workbook.sheetnames:
        if sheetname == RESUMEN_SHEET:
            continue
        ws = workbook[sheetname]
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 28

        # 1. Dar formato a la fila de encabezados
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = cell_border

        # 2. Localizar índices de columnas especiales (búsqueda insensible a mayúsculas)
        header_map = {str(cell.value).strip().upper(): idx + 1 for idx, cell in enumerate(ws[1])}
        col_estatus = header_map.get("ESTATUS_AUDITORIA")
        col_riesgo = header_map.get("CATEGORIA_RIESGO")
        col_tipo = header_map.get("TIPO_HALLAZGO")
        col_login = header_map.get("ULTIMO_LOGIN_DETECTADO")
        col_cuenta = header_map.get("ESTATUS_CUENTA_REPORTE")
        col_discrepancia = header_map.get("DISCREPANCIA_IDENTIDAD")
        col_dias = header_map.get("DIAS_POST_BAJA") or header_map.get("DIAS_DESPUES_DE_BAJA")

        # 3. Dar formato a filas de datos
        max_row = ws.max_row
        max_col = ws.max_column

        for r_idx in range(2, max_row + 1):
            ws.row_dimensions[r_idx].height = 20
            riesgo_val = str(ws.cell(row=r_idx, column=col_riesgo).value or "") if col_riesgo else ""
            if not riesgo_val and col_tipo:
                riesgo_val = str(ws.cell(row=r_idx, column=col_tipo).value or "")

            for c_idx in range(1, max_col + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.font = body_font
                cell.border = cell_border
                cell.alignment = regular_alignment

                col_name = str(ws.cell(row=1, column=c_idx).value or "").upper()
                val_str = str(cell.value or "")

                # Centrar fechas, claves, estatus y números
                if (
                    "FECHA" in col_name
                    or "DATE" in col_name
                    or c_idx in [col_login, col_estatus, col_riesgo, col_tipo, col_cuenta, col_discrepancia, col_dias]
                    or isinstance(cell.value, (int, float))
                ):
                    cell.alignment = center_alignment

                # Formato de fecha
                if "FECHA" in col_name or "DATE" in col_name:
                    if cell.is_date:
                        cell.number_format = "DD/MM/YYYY"

                # Resaltado condicional en ESTATUS_AUDITORIA
                if c_idx == col_estatus:
                    if val_str == "REVISAR":
                        if "HUÉRFANA" in riesgo_val:
                            cell.fill = ambar_fill
                            cell.font = ambar_font
                        else:
                            cell.fill = rojo_fill
                            cell.font = rojo_font
                    elif val_str == "INCIDENTE RESUELTO":
                        cell.fill = slate_fill
                        cell.font = slate_font
                    elif val_str == "OK":
                        if "REINGRESO" in riesgo_val:
                            cell.fill = azul_fill
                            cell.font = azul_font
                        else:
                            cell.fill = verde_fill
                            cell.font = verde_font

                # Resaltado condicional en CATEGORIA_RIESGO y TIPO_HALLAZGO
                if c_idx in [col_riesgo, col_tipo]:
                    if "CRÍTICO" in val_str or val_str in ["Acceso Post-Baja", "Ambos"]:
                        cell.fill = rojo_fill
                        cell.font = rojo_font
                    elif "ALTO" in val_str or "HUÉRFANA" in val_str or val_str == "Cuenta Activa":
                        cell.fill = ambar_fill
                        cell.font = ambar_font
                    elif "MEDIO" in val_str or "PASADO" in val_str or val_str == "INCIDENTE RESUELTO":
                        cell.fill = slate_fill
                        cell.font = slate_font
                    elif "REINGRESO" in val_str:
                        cell.fill = azul_fill
                        cell.font = azul_font
                    elif "CONFORME" in val_str or val_str in ["OK", "Ninguno"]:
                        cell.fill = verde_fill
                        cell.font = verde_font

                # Resaltado en DISCREPANCIA_IDENTIDAD
                if c_idx == col_discrepancia:
                    if cell.value is True or val_str.lower() in ["true", "si", "sí"]:
                        cell.fill = ambar_fill
                        cell.font = ambar_font

                # Formato especial en ESTATUS_CUENTA_REPORTE
                if c_idx == col_cuenta:
                    if val_str == "Activa":
                        cell.font = ambar_font
                    elif val_str == "Sin registro":
                        cell.font = sin_registro_font

                # Formato a ULTIMO_LOGIN_DETECTADO
                if c_idx == col_login:
                    if val_str == "Sin registro":
                        cell.font = sin_registro_font

                # Formato a DIAS_POST_BAJA
                if c_idx == col_dias:
                    if sheetname == "Riesgos_Activos" or "CRÍTICO" in riesgo_val:
                        cell.fill = rojo_fill
                        cell.font = rojo_font
                    elif sheetname == "Incidentes_Pasados" or "MEDIO" in riesgo_val:
                        cell.fill = slate_fill
                        cell.font = slate_font

        # 4. Auto-ajuste dinámico de ancho de columnas
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = cell.value
                if val is not None:
                    lines = str(val).split("\n")
                    line_len = max(len(l) for l in lines)
                    max_len = max(max_len, line_len)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 13)


def export_to_excel(
    df_completo: pd.DataFrame,
    df_riesgos_activos: pd.DataFrame,
    output_path: Path | str,
    df_incidentes_pasados: Optional[pd.DataFrame] = None,
    df_reingresos: Optional[pd.DataFrame] = None,
    metrics: Optional[AuditMetrics] = None,
) -> Path:
    """
    Genera el archivo final Excel con cinco pestañas:
    - 'Resumen_Auditoria' (primera pestaña: datos cuantitativos de la ejecución)
    - 'Auditoria_Completa'
    - 'Riesgos_Activos'
    - 'Incidentes_Pasados'
    - 'Reingresos'
    Aplica estilo visual con openpyxl y maneja excepciones de archivo bloqueado.
    """
    dest = Path(output_path)
    dest.parent.mkdir(parents=True, exist_ok=True)

    # Manejo de archivo abierto o bloqueado por otra aplicación
    try:
        if dest.exists():
            with open(dest, "r+b"):
                pass
    except (PermissionError, IOError):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        alt_dest = dest.parent / f"{dest.stem}_{timestamp}{dest.suffix}"
        print(
            f"\n[AVISO] El archivo destino '{dest.name}' está abierto en otra aplicación."
            f"\nSe guardará con nombre alternativo: '{alt_dest.name}'"
        )
        dest = alt_dest

    if df_incidentes_pasados is None:
        mask_pasados = df_completo["CATEGORIA_RIESGO"] == "MEDIO - INCIDENTE PASADO"
        df_incidentes_pasados = df_completo[mask_pasados].copy()

    if df_reingresos is None:
        mask_reingresos = df_completo["CATEGORIA_RIESGO"] == "POSIBLE REINGRESO"
        df_reingresos = df_completo[mask_reingresos].copy()

    # 1. Escribir DataFrames en sus respectivas 4 pestañas de datos
    with pd.ExcelWriter(dest, engine="openpyxl") as writer:
        df_completo.to_excel(writer, sheet_name="Auditoria_Completa", index=False)
        df_riesgos_activos.to_excel(writer, sheet_name="Riesgos_Activos", index=False)
        df_incidentes_pasados.to_excel(writer, sheet_name="Incidentes_Pasados", index=False)
        df_reingresos.to_excel(writer, sheet_name="Reingresos", index=False)

    # 2. Cargar con openpyxl para inyectar el Resumen en index 0 y estilizar
    wb = openpyxl.load_workbook(dest)
    create_resumen_sheet(wb, df_completo, metrics)
    apply_professional_excel_styling(wb)
    wb.active = 0
    wb.save(dest)
    return dest


# ============================================================================
# DETECCIÓN AUTOMÁTICA Y FLUJO INTERACTIVO
# ============================================================================


def auto_detect_input_files(directory: Path | str = ".") -> tuple[Optional[Path], Optional[Path]]:
    """
    Intenta descubrir automáticamente los archivos Universo y Reporte de Logins
    en el directorio especificado mediante patrones de nombres.
    Soporta extensiones .xlsx, .xls y .csv.
    """
    search_dir = Path(directory)
    if not search_dir.is_dir():
        return None, None

    all_files: list[Path] = []
    for pattern in ("*.xlsx", "*.xls", "*.csv"):
        all_files.extend(search_dir.glob(pattern))

    # Descartar archivos temporales (~$...) y archivos de salida previos
    valid_files = [
        f for f in all_files
        if not f.name.startswith("~$") and "resultado" not in f.name.lower()
    ]

    universo_candidate: Optional[Path] = None
    reporte_candidate: Optional[Path] = None

    # Buscar reporte con formato report<números> o similar
    for f in valid_files:
        name_lower = f.name.lower()
        if name_lower.startswith("report") and any(ch.isdigit() for ch in name_lower):
            reporte_candidate = f
            break

    # Si no tiene números, buscar por nombre 'report' o 'login'
    if reporte_candidate is None:
        for f in valid_files:
            if "report" in f.name.lower() or "login" in f.name.lower():
                reporte_candidate = f
                break

    # Buscar universo: priorizar archivos de producción o padrón real
    uni_candidates: list[Path] = []
    for f in valid_files:
        if f == reporte_candidate:
            continue
        name_lower = f.name.lower()
        if any(keyword in name_lower for keyword in ["prod", "universo", "padron", "baja", "empleado", "master"]):
            uni_candidates.append(f)

    if uni_candidates:
        prod_candidates = [f for f in uni_candidates if "prod" in f.name.lower()]
        non_example = [f for f in uni_candidates if "ejemplo" not in f.name.lower() and "sample" not in f.name.lower()]
        if prod_candidates:
            universo_candidate = prod_candidates[0]
        elif non_example:
            universo_candidate = non_example[0]
        else:
            universo_candidate = uni_candidates[0]

    # Si hay 2 archivos y solo uno es reporte, el otro es probablemente universo
    if reporte_candidate and not universo_candidate and len(valid_files) == 2:
        for f in valid_files:
            if f != reporte_candidate:
                universo_candidate = f
                break

    # Si hay universo y no reporte, y hay 2 archivos válidos
    if universo_candidate and not reporte_candidate and len(valid_files) == 2:
        for f in valid_files:
            if f != universo_candidate:
                reporte_candidate = f
                break

    return universo_candidate, reporte_candidate


def print_summary_box(metrics: AuditMetrics, output_file: Path) -> None:
    """Imprime un resumen visual profesional en terminal con métricas de ejecución."""
    border = "=" * 78
    divider = "-" * 78
    print("\n" + border)
    print("        RESUMEN EJECUTIVO DE AUDITORÍA DE ACCESOS POST-BAJA        ")
    print("     (Matriz de Riesgo Tripartita y Clasificación de Incidentes)   ")
    print(border)
    print(f" Filas procesadas / evaluadas                 : {metrics.total_evaluados:>7,}")
    print(divider)
    print(f" PESTAÑA 'Riesgos_Activos' (ACCIÓN INMEDIATA) : {metrics.total_revisar:>7,}")
    print(f"   * [CRÍTICO] Riesgo Activo (Login + Activa) : {metrics.total_critico_activo:>7,}")
    print(f"   * [ALTO]    Cuenta Huérfana (Solo Activa)  : {metrics.total_alto_huerfana:>7,}")
    print(divider)
    print(f" PESTAÑA 'Incidentes_Pasados' (REMEDIADOS)    : {metrics.total_medio_incidente_pasado:>7,}")
    print(f"   * [MEDIO]   Incidente Pasado (Login + Inact): {metrics.total_medio_incidente_pasado:>7,}")
    print(divider)
    print(f" PESTAÑA 'Reingresos' Y REGISTROS CONFORMES   : {metrics.total_ok:>7,}")
    print(f"   * [REINGRESO] Posible Reingreso            : {metrics.total_posible_reingreso:>7,}")
    print(f"   * [CONFORME]  Conformes (OK)               : {metrics.total_conforme:>7,}")
    print(f"   * Usuarios 'Sin registro' en reporte       : {metrics.total_sin_registro:>7,}")
    print(divider)
    print(f" Discrepancias de Identidad detectadas        : {metrics.total_discrepancias_identidad:>7,}")
    print(f" (Mismo correo en reporte pero con distinto nombre / cuenta reasignada)")
    print(divider)
    print(f" Tiempo total de ejecución                    : {metrics.tiempo_total_segundos:>7.3f} s")
    print(f" Archivo Excel generado satisfactoriamente (5 pestañas, primera: 'Resumen_Auditoria'):")
    print(f" -> {output_file.resolve()}")
    print(border + "\n")


# ============================================================================
# CONFIGURACIÓN PERSISTENTE Y RUTAS INTERACTIVAS
# ============================================================================

OUTPUT_NAME_PREFIX = "Auditoria_Accesos_Resultado"


def build_output_name(day: Optional[datetime] = None) -> str:
    """Nombre del archivo de resultados con la fecha de la auditoría (AAAA-MM-DD)."""
    return f"{OUTPUT_NAME_PREFIX}_{(day or datetime.now()).strftime('%Y-%m-%d')}.xlsx"
SCRIPT_DIR = Path(__file__).resolve().parent


def get_config_path() -> Path:
    """Ruta del archivo de configuración (fuera de la carpeta de scripts).

    Se puede sobrescribir con la variable de entorno AUDITORIA_CONFIG.
    """
    override = os.environ.get("AUDITORIA_CONFIG")
    if override:
        return Path(override).expanduser()
    return Path.home() / ".auditoria_accesos" / "config.json"


def load_config() -> Dict[str, str]:
    """Lee la configuración guardada; devuelve {} si no existe o es inválida."""
    try:
        data = json.loads(get_config_path().read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_config(config: Dict[str, str]) -> None:
    """Guarda la configuración; un fallo de escritura nunca interrumpe la auditoría."""
    path = get_config_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError as exc:
        print(f"[AVISO] No se pudo guardar la configuración en '{path}': {exc}")


def _clean_path_input(raw: str) -> str:
    """Quita espacios y comillas (típicas al arrastrar archivos a la consola)."""
    return raw.strip().strip('"').strip("'").strip()


def _ask(prompt: str) -> str:
    """input() que lanza EOFError si no hay consola interactiva."""
    return input(prompt)


def _ask_yes_no(prompt: str, default: bool = True) -> bool:
    suffix = "[S/n]" if default else "[s/N]"
    while True:
        answer = _ask(f"{prompt} {suffix}: ").strip().lower()
        if not answer:
            return default
        if answer in ("s", "si", "sí", "y", "yes"):
            return True
        if answer in ("n", "no"):
            return False
        print("    Responda 's' o 'n'.")


def _is_inside(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


def prompt_universo_path(suggested: Optional[str] = None) -> str:
    """Solicita la ruta del archivo Universo hasta obtener un archivo existente."""
    while True:
        hint = f" [{suggested}]" if suggested else ""
        raw = _clean_path_input(_ask(f"Ruta del archivo Universo{hint}: "))
        candidate = raw or (suggested or "")
        if not candidate:
            print("    Debe indicar una ruta.")
            continue
        path = Path(candidate).expanduser()
        if path.is_file():
            return str(path)
        print(f"    [ERROR] No existe el archivo: '{path}'")


def prompt_new_output_dir() -> Path:
    """Solicita una carpeta de resultados fuera de la carpeta de scripts y la crea previa confirmación."""
    while True:
        raw = _clean_path_input(_ask("Ruta de la carpeta de resultados: "))
        if not raw:
            print("    Debe indicar una ruta.")
            continue
        folder = Path(raw).expanduser()
        if _is_inside(folder, SCRIPT_DIR):
            print(f"    [ERROR] La carpeta debe estar fuera de la carpeta de scripts ('{SCRIPT_DIR}').")
            continue
        if folder.exists() and not folder.is_dir():
            print(f"    [ERROR] '{folder}' existe y no es una carpeta.")
            continue
        if not folder.exists():
            if not _ask_yes_no(f"La carpeta '{folder}' no existe. ¿Crearla?"):
                continue
            try:
                folder.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                print(f"    [ERROR] No se pudo crear la carpeta: {exc}")
                continue
        return folder


def resolve_output_dir(config: Dict[str, str]) -> Path:
    """Primera vez: pregunta la ruta. Siguientes: confirma la última o permite cambiarla."""
    last = config.get("output_dir")
    if last:
        last_path = Path(last)
        if _ask_yes_no(f"¿Guardar los resultados en '{last_path}'?"):
            if not last_path.exists():
                if not _ask_yes_no(f"La carpeta '{last_path}' ya no existe. ¿Crearla?"):
                    return resolve_output_dir({})
                last_path.mkdir(parents=True, exist_ok=True)
            return last_path
    return prompt_new_output_dir()


# ============================================================================
# FUNCIÓN PRINCIPAL Y PIPELINE COMPLETO
# ============================================================================


def run_audit_pipeline(
    universo_path: Path | str,
    reporte_path: Path | str,
    output_path: Optional[Path | str] = None,
    hoja_universo: str = "BajasU",
) -> tuple[Path, AuditMetrics]:
    """
    Ejecuta el pipeline completo de auditoría con validación compuesta y matriz de riesgo:
    1. Carga y validación defensiva de Universo en pestaña específica ('BajasU')
    2. Carga y preprocesamiento defensivo del Reporte de Logins (dropna de columnas vacías)
    3. Mapeo resiliente de columnas por texto (Username, Full Name, Active, Login, etc.)
    4. Consolidación de logins por llave compuesta (Correo + Nombre)
    5. Evaluación de matriz de riesgo tripartita (Riesgos Activos, Incidentes Pasados, Reingresos)
    6. Generación y formateo profesional de Excel (5 pestañas: Resumen_Auditoria,
       Auditoria_Completa, Riesgos_Activos, Incidentes_Pasados, Reingresos)
    """
    start_time = time.perf_counter()

    uni_path = Path(universo_path)
    rep_path = Path(reporte_path)
    out_path = Path(output_path) if output_path else Path(build_output_name())

    print(f"\n[+] Iniciando proceso de auditoría con validación compuesta y matriz de riesgo...")
    print(f"    * Archivo Universo          : {uni_path.name}")
    print(f"    * Archivo Reporte de Logins : {rep_path.name}")

    # 1. Lectura de Universo y deduplicación cronológica inicial
    df_uni, baja_col, email_col = read_universo(uni_path, target_sheet_name=hoja_universo)
    df_uni_dedup, reingreso_keys = deduplicate_universo(df_uni, baja_col=baja_col, email_col=email_col)
    print(f"    -> Universo cargado (pestaña '{hoja_universo}'): {len(df_uni)} registros originales -> {len(df_uni_dedup)} deduplicados.")
    print(f"       Columna Fecha Baja: '{baja_col}' | Columna Correo: '{email_col}'")

    # Mapeo de fecha de baja efectiva por clave compuesta para el reporte
    nombre_col_uni = find_column_by_substring(
        df_uni_dedup,
        candidates=["nombre", "nombre_completo", "empleado", "trabajador", "full name", "name"],
        default_index=2 if len(df_uni_dedup.columns) > 2 else 0,
        column_role="Nombre Completo (Columna C / Índice 2)",
    )
    clean_emails = clean_identifier_series(df_uni_dedup[email_col])
    clean_names = clean_name_series(df_uni_dedup[nombre_col_uni]) if nombre_col_uni else pd.Series("", index=df_uni_dedup.index)
    parsed_bajas = robust_parse_dates(df_uni_dedup[baja_col])
    fecha_baja_map: Dict[str, pd.Timestamp] = {}
    for em, nm, fb in zip(clean_emails, clean_names, parsed_bajas):
        if pd.notna(em) and pd.notna(fb):
            fecha_baja_map[f"{em}___{nm}"] = fb

    # 2. Lectura defensiva de Reporte
    df_rep, col_map = read_reporte_logins(rep_path)
    print(f"    -> Reporte cargado: {len(df_rep)} registros (columnas activas: {len(df_rep.columns)}).")
    print(f"       Mapeo detectado: Nombre='{col_map.get('name')}', Usuario='{col_map['user']}', Active='{col_map.get('active')}', Login='{col_map.get('login')}'")

    # 3. Consolidación de Logins por Llave Compuesta con selección cronológica
    consolidated_logins = consolidate_logins(df_rep, col_map, fecha_baja_map=fecha_baja_map)
    print(f"    -> Logins consolidados: {len(consolidated_logins)} llaves compuestas únicas en reporte.")

    # 4. Evaluación vectorizada con matriz de riesgo
    df_completo, df_riesgos_activos, metrics = execute_audit(
        df_universo=df_uni_dedup,
        baja_col=baja_col,
        email_col=email_col,
        consolidated_logins=consolidated_logins,
    )

    # 5. Exportación y formateo profesional en 5 pestañas
    metrics.fecha_auditoria = datetime.now().strftime("%d/%m/%Y")
    metrics.universo_archivo = uni_path.name
    metrics.reporte_archivo = rep_path.name
    metrics.total_universo_original = len(df_uni)
    metrics.total_registros_reporte = len(df_rep)
    saved_file = export_to_excel(df_completo, df_riesgos_activos, out_path, metrics=metrics)
    metrics.output_path = str(saved_file.resolve())
    metrics.tiempo_total_segundos = round(time.perf_counter() - start_time, 4)

    # 6. Resumen ejecutivo
    print_summary_box(metrics, saved_file)
    return saved_file, metrics


def main() -> int:
    """Punto de entrada de línea de comandos (CLI)."""
    parser = argparse.ArgumentParser(
        description="Automatización de Auditoría de Accesos Post-Baja con Validación Compuesta",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "-u",
        "--universo",
        type=str,
        default=None,
        help="Ruta al Archivo Universo (padrón maestro de empleados y bajas)",
    )
    parser.add_argument(
        "--hoja-universo",
        type=str,
        default="BajasU",
        help="Nombre de la pestaña a leer en el archivo Universo (por defecto: BajasU)",
    )
    parser.add_argument(
        "-r",
        "--reporte",
        type=str,
        default=None,
        help="Ruta al Archivo Reporte de Logins (report<id>.*)",
    )
    parser.add_argument(
        "-o",
        "--salida",
        type=str,
        default=None,
        help="Ruta completa del Excel de resultados (omite las preguntas de carpeta)",
    )
    parser.add_argument(
        "--carpeta-resultados",
        type=str,
        default=None,
        help="Carpeta de resultados (omite la pregunta interactiva y se recuerda para la próxima vez)",
    )
    parser.add_argument(
        "--auto",
        action="store_true",
        help="Intentar auto-descubrir los archivos en la carpeta actual sin solicitar confirmación",
    )

    args = parser.parse_args()
    config = load_config()

    try:
        # 1. Archivo Universo: argumento, o pregunta (sugiriendo detección automática / última ruta)
        universo_file = args.universo
        if not universo_file:
            auto_uni, auto_rep = auto_detect_input_files(".")
            suggested = str(auto_uni) if auto_uni else config.get("universo")
            universo_file = prompt_universo_path(suggested)
            if not args.reporte and auto_rep:
                args.reporte = str(auto_rep)
                print(f"[INFO] Reporte de Logins detectado automáticamente: {auto_rep.name}")

        # 2. Reporte de logins
        reporte_file = args.reporte
        if not reporte_file:
            reporte_file = _clean_path_input(_ask("Ingrese la ruta del Archivo Reporte de Logins: "))

        if not Path(universo_file).is_file():
            print(f"\n[ERROR] El archivo Universo '{universo_file}' no existe.", file=sys.stderr)
            return 1
        if not Path(reporte_file).is_file():
            print(f"\n[ERROR] El archivo Reporte de Logins '{reporte_file}' no existe.", file=sys.stderr)
            return 1

        # 3. Destino de resultados
        if args.salida:
            output_path = Path(args.salida)
        else:
            if args.carpeta_resultados:
                output_dir = Path(args.carpeta_resultados).expanduser()
                if _is_inside(output_dir, SCRIPT_DIR):
                    print("\n[ERROR] La carpeta de resultados debe estar fuera de la carpeta de scripts.", file=sys.stderr)
                    return 1
                output_dir.mkdir(parents=True, exist_ok=True)
            else:
                output_dir = resolve_output_dir(config)
            config["output_dir"] = str(output_dir.resolve())
            output_path = output_dir / build_output_name()

        config["universo"] = str(Path(universo_file).resolve())
        save_config(config)
    except EOFError:
        print(
            "\n[ERROR] Sin consola interactiva: indique --universo, --reporte y "
            "--carpeta-resultados (o --salida).",
            file=sys.stderr,
        )
        return 1

    try:
        run_audit_pipeline(
            universo_path=universo_file,
            reporte_path=reporte_file,
            output_path=output_path,
            hoja_universo=args.hoja_universo,
        )
        return 0
    except Exception as exc:
        print(f"\n[ERROR CRÍTICO] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
