#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 SIVREP CLOUD — Sistema Institucional de Validación de Repuestos
              y Levantamiento de Planos Eléctricos (STREAMLIT EDITION)
================================================================================
Versión    : 3.1.0-Cloud
Plataforma : Streamlit Cloud / Web
Python     : 3.9+

DEPENDENCIAS
────────────
  pip install streamlit pandas openpyxl PyPDF2 pdfplumber pytesseract pillow 
              opencv-python-headless numpy matplotlib seaborn 
              fuzzywuzzy python-Levenshtein PyMuPDF plotly
================================================================================
"""

import os
import sys
import re
import json
import logging
import io
import base64
from datetime import datetime
from pathlib import Path
from collections import defaultdict
from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional, Any, Callable
import traceback
import time

import streamlit as st
import pandas as pd
import numpy as np

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN GLOBAL
# ═══════════════════════════════════════════════════════════════════════════════
@dataclass(frozen=True)
class AppConfig:
    APP_NAME: str = "SIVREP CLOUD"
    VERSION: str = "3.1.0-Cloud"
    ORGANIZATION: str = "Departamento de Ingeniería"
    DEFAULT_THRESHOLD: int = 75
    MAX_WORDS_MATCH: int = 4
    OUTPUT_DIR: str = "output"
    TEMP_DIR: str = "temp"
    PDF_DPI: int = 300
    OCR_LANG: str = "eng+spa"

CONFIG = AppConfig()
Path(CONFIG.OUTPUT_DIR).mkdir(exist_ok=True)
Path(CONFIG.TEMP_DIR).mkdir(exist_ok=True)

log_buffer = io.StringIO()
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    handlers=[logging.StreamHandler(log_buffer), logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("SIVREP")

# ═══════════════════════════════════════════════════════════════════════════════
# PALETA DE COLORES INSTITUCIONAL
# ═══════════════════════════════════════════════════════════════════════════════
class Palette:
    BG_PRIMARY   = "#f0f4f8"
    BG_SECONDARY = "#ffffff"
    BG_TERTIARY  = "#e2e8f0"
    BG_HEADER    = "#1e3a5f"
    PRIMARY      = "#2b6cb0"
    PRIMARY_HOVER= "#2c5282"
    ACCENT       = "#3182ce"
    TEXT_PRIMARY   = "#1a202c"
    TEXT_SECONDARY = "#4a5568"
    TEXT_MUTED     = "#718096"
    SUCCESS      = "#276749"
    DANGER       = "#c53030"
    WARNING      = "#c05621"
    INFO         = "#2b6cb0"
    BORDER       = "#cbd5e0"

# ═══════════════════════════════════════════════════════════════════════════════
# CSS PERSONALIZADO
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
    .main-header { font-size: 2.2rem; font-weight: 700; color: #1e3a5f; text-align: center; 
                   padding: 1rem 0; border-bottom: 3px solid #2b6cb0; margin-bottom: 1.5rem; }
    .sub-header { font-size: 1.15rem; font-weight: 600; color: #1e3a5f; margin-top: 1.5rem; 
                  margin-bottom: 0.8rem; padding-left: 0.6rem; border-left: 4px solid #3182ce; }
    .metric-card { background: linear-gradient(145deg, #ffffff 0%, #f0f4f8 100%); border-radius: 10px; 
                   padding: 1.1rem; border: 1px solid #cbd5e0; box-shadow: 0 2px 8px rgba(30,58,95,0.08); }
    .metric-label { font-size: 0.78rem; color: #718096; text-transform: uppercase; letter-spacing: 1.2px; }
    .metric-value { font-size: 1.6rem; font-weight: 700; color: #1a202c; }
    .status-ok { color: #276749; } .status-warning { color: #c05621; } .status-danger { color: #c53030; }
    .info-box { background: linear-gradient(135deg, #ebf5fb 0%, #d4e6f1 100%); border-left: 4px solid #2b6cb0; 
                padding: 1.2rem; border-radius: 0 10px 10px 0; margin: 1rem 0; }
    .stTabs [data-baseweb="tab-list"] { gap: 6px; background-color: #f0f4f8; padding: 4px; border-radius: 10px; }
    .stTabs [data-baseweb="tab"] { background-color: #ffffff; border-radius: 8px; padding: 10px 20px; 
                                     font-weight: 600; color: #4a5568; }
    .stTabs [aria-selected="true"] { background-color: #2b6cb0 !important; color: #ffffff !important; }
    .footer { text-align: center; padding: 2rem 0 1rem 0; color: #718096; font-size: 0.82rem; 
              border-top: 1px solid #cbd5e0; margin-top: 2rem; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# IMPORTACIÓN DE LIBRERÍAS CON MANEJO GRACEFUL
# ═══════════════════════════════════════════════════════════════════════════════
PANDAS_OK = True
PDF_OK = PIL_OK = TESSERACT_OK = CV2_OK = FUZZY_OK = False
MATPLOTLIB_OK = SEABORN_OK = PYMUPDF_OK = PLOTLY_OK = False

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    import PyPDF2, pdfplumber
    PDF_OK = True
except ImportError as e:
    logger.error(f"PyPDF2/pdfplumber no disponibles: {e}")

try:
    import fitz
    PYMUPDF_OK = True
except ImportError:
    logger.warning("PyMuPDF no instalado.")

try:
    from PIL import Image, ImageEnhance, ImageFilter
    PIL_OK = True
except ImportError as e:
    logger.error(f"Pillow no disponible: {e}")

try:
    import pytesseract
    TESSERACT_OK = True
except ImportError as e:
    logger.error(f"pytesseract no disponible: {e}")

try:
    import cv2
    CV2_OK = True
except ImportError as e:
    logger.error(f"OpenCV no disponible: {e}")

try:
    from fuzzywuzzy import fuzz
    FUZZY_OK = True
except ImportError as e:
    logger.error(f"fuzzywuzzy no disponible: {e}")

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    MATPLOTLIB_OK = True
except ImportError as e:
    logger.error(f"matplotlib no disponible: {e}")

try:
    import seaborn as sns
    SEABORN_OK = True
except ImportError:
    pass

try:
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
    PLOTLY_OK = True
except ImportError:
    logger.warning("Plotly no disponible.")


# ═══════════════════════════════════════════════════════════════════════════════
# UTILIDADES GENERALES
# ═══════════════════════════════════════════════════════════════════════════════
class Utils:
    """Herramientas utilitarias estáticas."""

    @staticmethod
    def limpiar_texto(texto) -> str:
        if texto is None or (isinstance(texto, float) and pd.isna(texto)):
            return ""
        t = str(texto).upper().strip()
        t = re.sub(r"\s+", " ", t)
        t = re.sub(r"[^A-Z0-9\-_/\.\s]", "", t)
        return t

    @staticmethod
    def extraer_palabras_clave(texto: str, max_palabras: int = 4) -> List[str]:
        texto = Utils.limpiar_texto(texto)
        tokens = [t for t in texto.split() if len(t) > 2 or any(c.isdigit() for c in t)]
        def peso(t):
            return (any(c.isdigit() for c in t), "-" in t, len(t))
        return sorted(tokens, key=peso, reverse=True)[:max_palabras]

    @staticmethod
    def normalizar_codigo_parte(texto: str) -> str:
        texto = Utils.limpiar_texto(texto)
        patrones = [
            r"[A-Z]{2,6}[-\d]{3,20}[A-Z0-9\-]*",
            r"\d{3,6}[A-Z]{1,4}\d{0,4}",
            r"[A-Z]\d{2,4}[A-Z]\d{2,4}",
        ]
        for p in patrones:
            m = re.search(p, texto)
            if m:
                return m.group(0)
        return texto

    @staticmethod
    def calcular_similitud(s1: str, s2: str, metodo: str = "token_sort") -> float:
        if not FUZZY_OK:
            set1, set2 = set(s1.split()), set(s2.split())
            if not set1 or not set2:
                return 0.0
            return (len(set1 & set2) / len(set1 | set2)) * 100
        s1 = Utils.limpiar_texto(s1)
        s2 = Utils.limpiar_texto(s2)
        if metodo == "ratio":
            return float(fuzz.ratio(s1, s2))
        elif metodo == "partial":
            return float(fuzz.partial_ratio(s1, s2))
        elif metodo == "token_sort":
            return float(fuzz.token_sort_ratio(s1, s2))
        elif metodo == "token_set":
            return float(fuzz.token_set_ratio(s1, s2))
        elif metodo == "weighted":
            r1 = fuzz.token_set_ratio(s1, s2)
            r2 = fuzz.partial_ratio(s1, s2)
            r3 = fuzz.ratio(s1, s2)
            return r1 * 0.5 + r2 * 0.3 + r3 * 0.2
        return float(fuzz.token_sort_ratio(s1, s2))

    @staticmethod
    def es_codigo_parte(texto: str) -> bool:
        t = Utils.limpiar_texto(texto)
        if len(t) < 5:
            return False
        tiene_num = any(c.isdigit() for c in t)
        ratio = sum(1 for c in t if c.isdigit()) / len(t)
        return tiene_num and ("-" in t or ratio > 0.3)

    @staticmethod
    def ahora() -> str:
        return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


# ═══════════════════════════════════════════════════════════════════════════════
# MOTOR DE LECTURA DE ARCHIVOS
# ═══════════════════════════════════════════════════════════════════════════════
class FileReaderEngine:
    """Motor unificado de lectura de archivos maestros."""

    def __init__(self):
        self.historial: List[Dict] = []
        logger.info("FileReaderEngine inicializado.")

    def read_excel(self, file_obj, hoja: Optional[str] = None) -> pd.DataFrame:
        logger.info("Leyendo Excel desde stream")
        if hoja:
            df = pd.read_excel(file_obj, sheet_name=hoja, engine="openpyxl")
        else:
            xl = pd.ExcelFile(file_obj, engine="openpyxl")
            hoja = xl.sheet_names[0]
            df = pd.read_excel(file_obj, sheet_name=hoja, engine="openpyxl")
        df = self._limpiar_df(df)
        self.historial.append({"tipo": "excel", "filas": len(df), "columnas": len(df.columns)})
        logger.info(f"Excel cargado: {len(df)} filas x {len(df.columns)} columnas")
        return df

    def read_pdf_tablas(self, file_obj) -> Optional[pd.DataFrame]:
        if not PDF_OK:
            raise RuntimeError("Librerías PDF no disponibles.")
        logger.info("Extrayendo tablas de PDF")
        tablas = []
        with pdfplumber.open(file_obj) as pdf:
            for i, pagina in enumerate(pdf.pages):
                for tabla in pagina.extract_tables() or []:
                    if tabla and len(tabla) > 1:
                        tablas.append(pd.DataFrame(tabla[1:], columns=tabla[0]))
        if tablas:
            df = pd.concat(tablas, ignore_index=True)
            df = self._limpiar_df(df)
            self.historial.append({"tipo": "pdf_tabla", "filas": len(df)})
            return df
        logger.warning("No se detectaron tablas en el PDF.")
        return None

    def pdf_a_imagenes(self, file_obj, dpi: int = 300) -> List[str]:
        if not PIL_OK:
            raise RuntimeError("Pillow no disponible.")
        rutas_salida = []
        if PYMUPDF_OK:
            logger.info("Convirtiendo PDF con PyMuPDF")
            file_obj.seek(0)
            doc = fitz.open(stream=file_obj.read(), filetype="pdf")
            zoom = dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)
            for i in range(len(doc)):
                pagina = doc.load_page(i)
                pix = pagina.get_pixmap(matrix=mat)
                out = Path(CONFIG.TEMP_DIR) / f"pdf_page_{i:03d}.png"
                pix.save(str(out))
                rutas_salida.append(str(out))
            doc.close()
            return rutas_salida
        else:
            raise RuntimeError("Para procesar PDFs instale PyMuPDF.")

    def _limpiar_df(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.dropna(how="all")
        df = df.dropna(axis=1, how="all")
        df.columns = [str(c).strip().upper() for c in df.columns]
        return df.reset_index(drop=True)


# ═══════════════════════════════════════════════════════════════════════════════
# MOTOR OCR
# ═══════════════════════════════════════════════════════════════════════════════
class OCREngine:
    """Extracción inteligente de TAGs y Modelos desde planos eléctricos."""

    def __init__(self):
        self.config_tesseract = f"--psm 6 -l {CONFIG.OCR_LANG}"
        self.patrones = [
            ("TAG_EXPLICITO", r"TAG\s*[:\-]?\s*([A-Z0-9\-]{3,20})"),
            ("TAG_GENERICO", r"([A-Z]{1,4}[-\.]?\d{2,5}[A-Z0-9\-]*)"),
            ("NPARTE_EXPLICITO", r"N[°\s]*PARTE\s*[:\-]?\s*([A-Z0-9\-]{3,25})"),
            ("PARTNO", r"PART\s*NO\s*[:\-]?\s*([A-Z0-9\-]{3,25})"),
            ("MODELO_EXPLICITO", r"MODELO\s*[:\-]?\s*([A-Z0-9\-\s]{3,30})"),
            ("MODEL_EXPLICITO", r"MODEL\s*[:\-]?\s*([A-Z0-9\-\s]{3,30})"),
            ("SERIE", r"SERIE\s*[:\-]?\s*([A-Z0-9\-]{3,20})"),
            ("BANNER_SENSOR", r"S18[A-Z0-9]{5,15}"),
            ("CODIGO_3LETRAS", r"\d{3}[A-Z]{1,4}\d{0,4}"),
        ]
        logger.info("OCREngine inicializado.")

    def preprocesar(self, ruta_imagen: str) -> np.ndarray:
        if not CV2_OK:
            raise RuntimeError("OpenCV no disponible.")
        img = cv2.imread(ruta_imagen, cv2.IMREAD_GRAYSCALE)
        if img is None:
            raise ValueError(f"No se pudo cargar: {ruta_imagen}")
        img = cv2.resize(img, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC)
        img = cv2.fastNlMeansDenoising(img, None, 10, 7, 21)
        img = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                     cv2.THRESH_BINARY, 11, 2)
        kernel = np.ones((1, 1), np.uint8)
        img = cv2.dilate(img, kernel, iterations=1)
        return img

    def extraer_texto(self, ruta_imagen: str, preprocesar: bool = True) -> str:
        if not TESSERACT_OK:
            raise RuntimeError("Tesseract OCR no disponible.")
        if preprocesar and CV2_OK:
            img = self.preprocesar(ruta_imagen)
            temp = Path(CONFIG.TEMP_DIR) / f"ocr_preproc_{Path(ruta_imagen).stem}.png"
            cv2.imwrite(str(temp), img)
            return pytesseract.image_to_string(img, config=self.config_tesseract).upper()
        return pytesseract.image_to_string(Image.open(ruta_imagen), config=self.config_tesseract).upper()

    def extraer_datos(self, texto: str) -> List[Dict[str, str]]:
        resultados = []
        lineas = [l.strip() for l in texto.split("
") if len(l.strip()) > 2]
        for i, linea in enumerate(lineas):
            tag = None
            modelo = None
            for nombre, patron in self.patrones[:2]:
                m = re.search(patron, linea)
                if m:
                    candidato = m.group(1) if m.groups() else m.group(0)
                    if len(candidato) > 2:
                        tag = candidato
                        break
            for nombre, patron in self.patrones[2:]:
                m = re.search(patron, linea)
                if m:
                    candidato = m.group(1) if m.groups() else m.group(0)
                    if Utils.es_codigo_parte(candidato):
                        modelo = candidato
                        break
            if tag and not modelo and i + 1 < len(lineas):
                for nombre, patron in self.patrones[2:]:
                    m = re.search(patron, lineas[i + 1])
                    if m:
                        candidato = m.group(1) if m.groups() else m.group(0)
                        if Utils.es_codigo_parte(candidato):
                            modelo = candidato
                            break
            if tag or modelo:
                resultados.append({
                    "TAG": tag or "NO_IDENTIFICADO",
                    "MODELO": modelo or "NO_IDENTIFICADO",
                    "TEXTO_ORIGINAL": linea,
                    "METODO_EXTRACCION": "OCR_ESTRUCTURADO"
                })
        if not resultados:
            resultados = self._extraer_por_bloques(texto)
        return resultados

    def _extraer_por_bloques(self, texto: str) -> List[Dict[str, str]]:
        resultados = []
        palabras = texto.split()
        i = 0
        while i < len(palabras) - 1:
            w1, w2 = palabras[i], palabras[i + 1] if i + 1 < len(palabras) else ""
            if re.match(r"^[A-Z]{1,4}[-\.]?\d{2,5}", w1) and Utils.es_codigo_parte(w2):
                resultados.append({"TAG": w1, "MODELO": w2, "TEXTO_ORIGINAL": f"{w1} {w2}", "METODO_EXTRACCION": "OCR_BLOQUE"})
                i += 2
                continue
            if Utils.es_codigo_parte(w1) and len(w1) > 5:
                resultados.append({"TAG": w1, "MODELO": w2 if Utils.es_codigo_parte(w2) else "NO_IDENTIFICADO",
                                   "TEXTO_ORIGINAL": w1, "METODO_EXTRACCION": "OCR_AISLADO"})
            i += 1
        return resultados

    def procesar_imagen(self, ruta: str) -> List[Dict[str, str]]:
        logger.info(f"OCR sobre imagen: {ruta}")
        texto = self.extraer_texto(ruta)
        return self.extraer_datos(texto)

    def procesar_pdf(self, file_obj) -> List[Dict[str, str]]:
        logger.info("OCR sobre PDF")
        motor = FileReaderEngine()
        imagenes = motor.pdf_a_imagenes(file_obj, dpi=CONFIG.PDF_DPI)
        todos = []
        for idx, img in enumerate(imagenes):
            try:
                res = self.procesar_imagen(img)
                for r in res:
                    r["PAGINA_PDF"] = idx + 1
                todos.extend(res)
            except Exception as e:
                logger.error(f"Error página {idx + 1}: {e}")
        logger.info(f"OCR PDF finalizado: {len(todos)} componentes.")
        return todos


# ═══════════════════════════════════════════════════════════════════════════════
# MOTOR FUZZY MATCHING
# ═══════════════════════════════════════════════════════════════════════════════
class SparePartMatcher:
    """Comparación inteligente con múltiples estrategias."""

    def __init__(self, umbral: int = 75, max_palabras: int = 4, estrategia: str = "weighted"):
        self.umbral = umbral
        self.max_palabras = max_palabras
        self.estrategia = estrategia
        self.estadisticas = {
            "total_comparaciones": 0,
            "exactas": 0,
            "fuzzy": 0,
            "sin_coincidencia": 0,
            "tiempo_seg": 0.0
        }
        logger.info(f"Matcher listo (umbral={umbral}, estrategia={estrategia}).")

    def comparar_uno(self, origen: str, destino: str) -> Tuple[bool, float, str]:
        self.estadisticas["total_comparaciones"] += 1
        s = Utils.limpiar_texto(origen)
        t = Utils.limpiar_texto(destino)
        if not s or not t:
            return False, 0.0, "vacio"
        if s == t:
            self.estadisticas["exactas"] += 1
            return True, 100.0, "exacto"
        cs = Utils.normalizar_codigo_parte(s)
        ct = Utils.normalizar_codigo_parte(t)
        if cs and ct and cs == ct:
            self.estadisticas["exactas"] += 1
            return True, 100.0, "codigo_parte_exacto"
        ws = set(Utils.extraer_palabras_clave(s, self.max_palabras))
        wt = set(Utils.extraer_palabras_clave(t, self.max_palabras))
        if ws and wt:
            inter = ws & wt
            if len(inter) >= min(2, len(ws), len(wt)):
                score = (len(inter) / max(len(ws), len(wt))) * 100
                if score >= self.umbral:
                    self.estadisticas["fuzzy"] += 1
                    return True, round(score, 2), "palabras_clave"
        score = Utils.calcular_similitud(s, t, self.estrategia)
        if score >= self.umbral:
            self.estadisticas["fuzzy"] += 1
            return True, round(score, 2), f"fuzzy_{self.estrategia}"
        self.estadisticas["sin_coincidencia"] += 1
        return False, round(score, 2), "sin_coincidencia"

    def comparar_dataframes(self, df_origen: pd.DataFrame, col_origen: str,
                            df_destino: pd.DataFrame, col_destino: str) -> pd.DataFrame:
        t0 = time.time()
        resultados = []
        total = len(df_origen)
        valores_destino = [(i, Utils.limpiar_texto(str(v))) for i, v in enumerate(df_destino[col_destino].astype(str))]

        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, fila in df_origen.iterrows():
            val_origen = str(fila[col_origen]) if pd.notna(fila[col_origen]) else ""
            if not val_origen.strip():
                continue
            best_score = 0.0
            best_match = ""
            best_metodo = ""
            best_idx = -1
            for t_idx, t_val in valores_destino:
                match, score, metodo = self.comparar_uno(val_origen, t_val)
                if score > best_score:
                    best_score = score
                    best_match = t_val
                    best_metodo = metodo
                    best_idx = t_idx
                    if score == 100.0:
                        break
            coincide = best_score >= self.umbral
            fila_res = {
                "IDX_ORIGEN": idx,
                "VALOR_ORIGEN": val_origen,
                "VALOR_DESTINO": best_match,
                "SCORE": best_score,
                "METODO": best_metodo,
                "COINCIDE": "SÍ" if coincide else "NO",
                "IDX_DESTINO": best_idx if coincide else -1
            }
            for c in df_origen.columns:
                fila_res[f"SRC_{c}"] = fila[c]
            resultados.append(fila_res)

            if int(idx) % 10 == 0:
                progress = min((int(idx) + 1) / total, 1.0)
                progress_bar.progress(progress)
                status_text.text(f"Progreso: {int(idx) + 1}/{total} ({progress*100:.1f}%)")

        progress_bar.empty()
        status_text.empty()
        self.estadisticas["tiempo_seg"] = time.time() - t0
        logger.info(f"Comparación lista en {self.estadisticas['tiempo_seg']:.2f}s")
        return pd.DataFrame(resultados)


# ═══════════════════════════════════════════════════════════════════════════════
# MOTOR ESTADÍSTICO
# ═══════════════════════════════════════════════════════════════════════════════
class StatisticsEngine:
    """Genera métricas, gráficos y reportes institucionales."""

    def __init__(self):
        self.reporte: Dict[str, Any] = {}
        logger.info("StatisticsEngine inicializado.")

    def analizar(self, df_resultado: pd.DataFrame, df_origen: pd.DataFrame,
                 df_destino: pd.DataFrame) -> Dict[str, Any]:
        total_o, total_d = len(df_origen), len(df_destino)
        if "COINCIDE" not in df_resultado.columns:
            raise ValueError("DataFrame sin columna COINCIDE")
        coin = df_resultado[df_resultado["COINCIDE"] == "SÍ"]
        no_coin = df_resultado[df_resultado["COINCIDE"] == "NO"]
        metricas = {
            "total_origen": total_o,
            "total_destino": total_d,
            "comparados": len(df_resultado),
            "coincidencias": len(coin),
            "no_coincidencias": len(no_coin),
            "tasa_exito": round(len(coin) / len(df_resultado) * 100, 2) if len(df_resultado) else 0,
            "cobertura": round(len(coin) / total_o * 100, 2) if total_o else 0,
            "score_promedio": round(coin["SCORE"].mean(), 2) if len(coin) else 0,
            "score_min": round(coin["SCORE"].min(), 2) if len(coin) else 0,
            "score_max": round(coin["SCORE"].max(), 2) if len(coin) else 0,
        }
        if "METODO" in df_resultado.columns:
            metricas["metodos"] = df_resultado["METODO"].value_counts().to_dict()
        metricas["vacios_origen"] = {c: int(df_origen[c].isna().sum()) for c in df_origen.columns}
        metricas["vacios_destino"] = {c: int(df_destino[c].isna().sum()) for c in df_destino.columns}
        self.reporte = metricas
        return metricas

    def generar_graficos_plotly(self):
        """Genera gráficos interactivos con Plotly."""
        if not self.reporte:
            return None

        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=("Distribución de Resultados", "Métricas de Desempeño",
                           "Métodos de Coincidencia", "Resumen Ejecutivo"),
            specs=[[{"type": "pie"}, {"type": "bar"}],
                   [{"type": "bar"}, {"type": "domain"}]]
        )

        # 1. Pie chart
        sizes = [self.reporte["coincidencias"], self.reporte["no_coincidencias"]]
        fig.add_trace(go.Pie(
            labels=["Coincidencias", "No Coincidencias"],
            values=sizes,
            marker_colors=["#38a169", "#e53e3e"],
            textinfo="label+percent",
            hole=0.4,
        ), row=1, col=1)

        # 2. Barras horizontales
        metricas = ["tasa_exito", "cobertura", "score_promedio"]
        valores = [self.reporte.get(m, 0) for m in metricas]
        nombres = ["Tasa Éxito (%)", "Cobertura (%)", "Score Promedio"]
        fig.add_trace(go.Bar(
            x=valores,
            y=nombres,
            orientation='h',
            marker_color=["#3182ce", "#805ad5", "#dd6b20"],
            text=[f"{v:.1f}" for v in valores],
            textposition="outside",
        ), row=1, col=2)

        # 3. Métodos
        if "metodos" in self.reporte:
            met = self.reporte["metodos"]
            fig.add_trace(go.Bar(
                x=list(met.keys()),
                y=list(met.values()),
                marker_color="#2b6cb0",
            ), row=2, col=1)

        # 4. Resumen como indicadores
        fig.add_trace(go.Indicator(
            mode="number+delta",
            value=self.reporte["coincidencias"],
            title={"text": "Coincidencias"},
            domain={'row': 1, 'column': 1},
        ), row=2, col=2)

        fig.update_layout(
            title_text="SIVREP — Análisis Estadístico de Validación",
            title_font_size=18,
            title_font_color="#1e3a5f",
            template="plotly_white",
            height=700,
            showlegend=False,
            paper_bgcolor='#ffffff',
            plot_bgcolor='#f0f4f8',
            font=dict(family="Inter, Arial, sans-serif", color="#1a202c"),
        )

        return fig

    def generar_excel(self, df_resultado: pd.DataFrame, df_unificado: Optional[pd.DataFrame] = None) -> bytes:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as w:
            df_resultado.to_excel(w, sheet_name="RESULTADOS_DETALLADOS", index=False)
            if "COINCIDE" in df_resultado.columns:
                df_resultado[df_resultado["COINCIDE"] == "SÍ"].to_excel(w, sheet_name="COINCIDENCIAS", index=False)
                df_resultado[df_resultado["COINCIDE"] == "NO"].to_excel(w, sheet_name="NO_COINCIDENCIAS", index=False)
            if df_unificado is not None:
                df_unificado.to_excel(w, sheet_name="LISTADO_UNIFICADO", index=False)
            pd.DataFrame([self.reporte]).to_excel(w, sheet_name="ESTADISTICAS", index=False)
        self._formatear_excel(buffer)
        logger.info("Excel generado en memoria")
        buffer.seek(0)
        return buffer.getvalue()

    def _formatear_excel(self, buffer):
        try:
            wb = load_workbook(buffer)
            header_fill = PatternFill(start_color="1e3a5f", end_color="1e3a5f", fill_type="solid")
            header_font = Font(color="f7fafc", bold=True, size=11, name="Segoe UI")
            header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            thin = Border(left=Side(style="thin", color="cbd5e0"),
                          right=Side(style="thin", color="cbd5e0"),
                          top=Side(style="thin", color="cbd5e0"),
                          bottom=Side(style="thin", color="cbd5e0"))
            for nombre in wb.sheetnames:
                ws = wb[nombre]
                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = header_align
                    cell.border = thin
                for col in ws.columns:
                    max_len = 0
                    letra = col[0].column_letter
                    for cell in col:
                        try:
                            if cell.value:
                                max_len = max(max_len, len(str(cell.value)))
                        except:
                            pass
                    ws.column_dimensions[letra].width = min(max_len + 3, 55)
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
            wb.save(buffer)
        except Exception as e:
            logger.error(f"Error formateando Excel: {e}")


# ═══════════════════════════════════════════════════════════════════════════════
# CONSOLIDADOR UNIFICADO
# ═══════════════════════════════════════════════════════════════════════════════
class UnifiedConsolidator:
    """Genera un listado maestro único a partir de dos fuentes."""

    COLUMNAS_PRIORITARIAS = [
        "SISTEMA", "EQUIPO Y SUBEQUIPO", "ESPECIALIDAD", "EQUIPO O REPUESTO",
        "ARTICULO", "FAMILIA", "TAG", "FABRICANTE", "MODELO", "SERIE",
        "N° PARTE", "NOMBRE", "DESCRIPCION"
    ]

    def __init__(self):
        logger.info("UnifiedConsolidator inicializado.")

    def consolidar(self, df_o: pd.DataFrame, df_d: pd.DataFrame,
                   df_res: pd.DataFrame, col_o: str, col_d: str) -> pd.DataFrame:
        idx_coin = df_res[df_res["COINCIDE"] == "SÍ"]["IDX_ORIGEN"].tolist()
        df_o_m = df_o.copy()
        df_o_m["_ORIGEN"] = "ARCHIVO_1"
        df_o_m["_ESTADO"] = df_o_m.index.map(lambda x: "COINCIDE" if x in idx_coin else "UNICO_A1")
        idx_t_usados = df_res[df_res["COINCIDE"] == "SÍ"]["IDX_DESTINO"].dropna().astype(int).unique()
        df_d_u = df_d.copy()
        df_d_u = df_d_u[~df_d_u.index.isin(idx_t_usados)]
        df_d_u["_ORIGEN"] = "ARCHIVO_2"
        df_d_u["_ESTADO"] = "UNICO_A2"
        no1 = self._normalizar_cols(df_o_m)
        no2 = self._normalizar_cols(df_d_u)
        todas = list(set(no1.columns) | set(no2.columns))
        for c in todas:
            if c not in no1.columns:
                no1[c] = ""
            if c not in no2.columns:
                no2[c] = ""
        orden = [c for c in self.COLUMNAS_PRIORITARIAS if c in todas]
        orden += sorted([c for c in todas if c not in self.COLUMNAS_PRIORITARIAS])
        no1, no2 = no1[orden], no2[orden]
        uni = pd.concat([no1, no2], ignore_index=True)
        uni["_FECHA_CONSOLIDACION"] = Utils.ahora()
        uni["_ID_UNICO"] = [f"REP-{i+1:06d}" for i in range(len(uni))]
        logger.info(f"Consolidado: {len(uni)} repuestos únicos.")
        return uni

    def _normalizar_cols(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        mapping = {
            r"SISTEMA|SIST": "SISTEMA",
            r"EQUIPO\s*Y\s*SUBEQUIPO|EQUIPO\s*SUB|EQUIPO(?!\s*O)": "EQUIPO Y SUBEQUIPO",
            r"ESPECIALIDAD|ESP": "ESPECIALIDAD",
            r"EQUIPO\s*O\s*REPUESTO|REPUESTO|EQUIPO\s*REP": "EQUIPO O REPUESTO",
            r"ARTICULO|ARTÍCULO|ART(?!I)": "ARTICULO",
            r"FAMILIA|FAM": "FAMILIA",
            r"^TAG$|TAG\s*ID": "TAG",
            r"FABRICANTE|FABR|MARCA|MANUFACTURER|MAKER": "FABRICANTE",
            r"MODELO|MODEL(?!O)|MOD(?=ELO)": "MODELO",
            r"SERIE|SER(?!I)|SERIAL": "SERIE",
            r"N[°\s]*PARTE|PART\s*NO|PART\s*NUMBER|NUMERO\s*PARTE|N\s*PARTE|N°\s*PARTE|PARTE": "N° PARTE",
            r"NOMBRE|DESCRIPCION|DESCRIPTION|DESC(?!R)": "NOMBRE",
        }
        nuevas = {}
        for col in df.columns:
            asignado = False
            for patron, estandar in mapping.items():
                if re.search(patron, col, re.IGNORECASE):
                    nuevas[col] = estandar
                    asignado = True
                    break
            if not asignado:
                nuevas[col] = col
        df.rename(columns=nuevas, inplace=True)
        return df


# ═══════════════════════════════════════════════════════════════════════════════
# INTERFAZ STREAMLIT
# ═══════════════════════════════════════════════════════════════════════════════
def init_session_state():
    """Inicializa variables de sesión."""
    defaults = {
        'df_origen': None,
        'df_destino': None,
        'df_resultado': None,
        'df_unificado': None,
        'df_ocr': None,
        'motor_archivos': FileReaderEngine(),
        'motor_ocr': OCREngine(),
        'matcher': SparePartMatcher(),
        'stats': StatisticsEngine(),
        'consolidador': UnifiedConsolidator(),
        'metricas': None,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


def render_header():
    """Renderiza encabezado institucional."""
    st.markdown(f"""
    <div class='main-header'>
        🔧 SIVREP CLOUD
        <div style='font-size: 0.85rem; color: #4a5568; font-weight: 400; margin-top: 0.3rem;'>
            Sistema Institucional de Validación de Repuestos y Planos Eléctricos v{CONFIG.VERSION}
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_sidebar():
    """Panel lateral de navegación."""
    st.sidebar.markdown(f"""
    <div style='text-align: center; padding: 1rem 0; border-bottom: 2px solid #2b6cb0; margin-bottom: 1rem;'>
        <h2 style='color: #1e3a5f; margin: 0; font-size: 1.4rem;'>🔧 SIVREP</h2>
        <p style='color: #718096; font-size: 0.8rem; margin: 0.3rem 0 0 0;'>Validación de Repuestos</p>
    </div>
    """, unsafe_allow_html=True)

    st.sidebar.markdown("#### 📋 Navegación")
    pagina = st.sidebar.radio("", [
        "🏠 Inicio",
        "📊 Validación de Repuestos",
        "⚡ Planos Eléctricos (OCR)",
        "📈 Estadísticas y Reportes"
    ], label_visibility="collapsed")

    st.sidebar.markdown("---")
    st.sidebar.markdown("#### ℹ️ Estado del Sistema")

    deps = {
        "Pandas/OpenPyXL": PANDAS_OK,
        "PyPDF2/PDFPlumber": PDF_OK,
        "Pillow": PIL_OK,
        "Tesseract OCR": TESSERACT_OK,
        "OpenCV": CV2_OK,
        "FuzzyWuzzy": FUZZY_OK,
        "Matplotlib": MATPLOTLIB_OK,
        "PyMuPDF": PYMUPDF_OK,
        "Plotly": PLOTLY_OK,
    }

    for nombre, ok in deps.items():
        color = "🟢" if ok else "🔴"
        st.sidebar.markdown(f"{color} {nombre}")

    return pagina


def render_home():
    """Página de inicio."""
    st.markdown("""
    <div class='info-box'>
        <h4 style='margin: 0 0 0.5rem 0; color: #1e3a5f;'>👋 Bienvenido a SIVREP Cloud</h4>
        <p style='margin: 0; color: #4a5568;'>Sistema Institucional de Validación de Repuestos y Planos Eléctricos</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class='metric-card'>
            <div class='metric-label'>Módulo 1</div>
            <div style='font-size: 1.1rem; font-weight: 600; color: #1e3a5f;'>📊 Validación de Repuestos</div>
            <p style='color: #718096; font-size: 0.85rem; margin-top: 0.5rem;'>Comparación cruzada de maestros Excel/PDF con fuzzy matching inteligente.</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class='metric-card'>
            <div class='metric-label'>Módulo 2</div>
            <div style='font-size: 1.1rem; font-weight: 600; color: #1e3a5f;'>⚡ OCR Planos Eléctricos</div>
            <p style='color: #718096; font-size: 0.85rem; margin-top: 0.5rem;'>Extracción automática de TAGs y Modelos desde planos PDF e imágenes.</p>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class='metric-card'>
            <div class='metric-label'>Módulo 3</div>
            <div style='font-size: 1.1rem; font-weight: 600; color: #1e3a5f;'>📈 Estadísticas</div>
            <p style='color: #718096; font-size: 0.85rem; margin-top: 0.5rem;'>Análisis de desempeño, gráficos interactivos y reportes institucionales.</p>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("""
    <div style='margin-top: 2rem;'>
        <h3 style='color: #2b6cb0;'>✨ Funcionalidades Principales</h3>
        <ul style='color: #4a5568; line-height: 2;'>
            <li>📁 Carga de archivos <b>Excel (.xlsx, .xls)</b> y <b>PDF tabulares</b></li>
            <li>🔍 Comparación semántica difusa (<b>fuzzy matching</b>) con múltiples estrategias</li>
            <li>📊 Estadísticas completas: coincidencias, cobertura, scores, métodos</li>
            <li>🏭 Consolidación unificada de listados maestros sin duplicados</li>
            <li>⚡ <b>OCR</b> de planos eléctricos: extracción de TAGs y Modelos</li>
            <li>📋 Exportación a <b>Excel con formato institucional</b></li>
            <li>📈 Gráficos interactivos de análisis de desempeño</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)


def render_validacion():
    """Pestaña de validación de repuestos."""
    st.markdown("<div class='sub-header'>📊 Validación Cruzada de Repuestos</div>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 📁 Archivo 1 (Base)")
        file1 = st.file_uploader("Cargar Excel o PDF", type=['xlsx', 'xls', 'pdf'], key="file1")
        if file1:
            try:
                ext = Path(file1.name).suffix.lower()
                if ext in ('.xlsx', '.xls'):
                    st.session_state.df_origen = st.session_state.motor_archivos.read_excel(file1)
                elif ext == '.pdf':
                    st.session_state.df_origen = st.session_state.motor_archivos.read_pdf_tablas(file1)
                st.success(f"✅ {file1.name}: {len(st.session_state.df_origen)} filas x {len(st.session_state.df_origen.columns)} columnas")
            except Exception as e:
                st.error(f"Error cargando archivo 1: {e}")

    with col2:
        st.markdown("#### 📁 Archivo 2 (Comparar)")
        file2 = st.file_uploader("Cargar Excel o PDF", type=['xlsx', 'xls', 'pdf'], key="file2")
        if file2:
            try:
                ext = Path(file2.name).suffix.lower()
                if ext in ('.xlsx', '.xls'):
                    st.session_state.df_destino = st.session_state.motor_archivos.read_excel(file2)
                elif ext == '.pdf':
                    st.session_state.df_destino = st.session_state.motor_archivos.read_pdf_tablas(file2)
                st.success(f"✅ {file2.name}: {len(st.session_state.df_destino)} filas x {len(st.session_state.df_destino.columns)} columnas")
            except Exception as e:
                st.error(f"Error cargando archivo 2: {e}")

    if st.session_state.df_origen is not None and st.session_state.df_destino is not None:
        st.markdown("---")
        st.markdown("#### ⚙️ Parámetros de Comparación")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            col_origen = st.selectbox("Columna Archivo 1", st.session_state.df_origen.columns, key="col_o")
        with col2:
            col_destino = st.selectbox("Columna Archivo 2", st.session_state.df_destino.columns, key="col_d")
        with col3:
            umbral = st.slider("Umbral mínimo (%)", 0, 100, 75)
        with col4:
            estrategia = st.selectbox("Estrategia fuzzy", ["weighted", "ratio", "partial", "token_sort", "token_set"])

        max_pal = st.slider("Máx. palabras clave", 1, 10, 4)

        if st.button("▶ EJECUTAR VALIDACIÓN", type="primary", use_container_width=True):
            with st.spinner("Ejecutando comparación..."):
                try:
                    st.session_state.matcher = SparePartMatcher(umbral=umbral, max_palabras=max_pal, estrategia=estrategia)
                    st.session_state.df_resultado = st.session_state.matcher.comparar_dataframes(
                        st.session_state.df_origen, col_origen,
                        st.session_state.df_destino, col_destino
                    )
                    st.session_state.df_unificado = st.session_state.consolidador.consolidar(
                        st.session_state.df_origen, st.session_state.df_destino,
                        st.session_state.df_resultado, col_origen, col_destino
                    )
                    st.session_state.metricas = st.session_state.stats.analizar(
                        st.session_state.df_resultado,
                        st.session_state.df_origen,
                        st.session_state.df_destino
                    )
                    st.success("✅ Validación completada exitosamente")
                except Exception as e:
                    st.error(f"Error en validación: {e}")

        if st.session_state.df_resultado is not None:
            st.markdown("---")
            st.markdown("#### 📋 Resultados de la Comparación")

            m = st.session_state.metricas
            if m:
                c1, c2, c3, c4, c5 = st.columns(5)
                with c1:
                    st.markdown(f"""
                    <div class='metric-card'>
                        <div class='metric-label'>Comparados</div>
                        <div class='metric-value'>{m['comparados']:,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                with c2:
                    st.markdown(f"""
                    <div class='metric-card'>
                        <div class='metric-label'>Coincidencias</div>
                        <div class='metric-value status-ok'>{m['coincidencias']:,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                with c3:
                    st.markdown(f"""
                    <div class='metric-card'>
                        <div class='metric-label'>No Coincidencias</div>
                        <div class='metric-value status-danger'>{m['no_coincidencias']:,}</div>
                    </div>
                    """, unsafe_allow_html=True)
                with c4:
                    st.markdown(f"""
                    <div class='metric-card'>
                        <div class='metric-label'>Tasa Éxito</div>
                        <div class='metric-value'>{m['tasa_exito']:.1f}%</div>
                    </div>
                    """, unsafe_allow_html=True)
                with c5:
                    st.markdown(f"""
                    <div class='metric-card'>
                        <div class='metric-label'>Score Promedio</div>
                        <div class='metric-value'>{m['score_promedio']:.1f}</div>
                    </div>
                    """, unsafe_allow_html=True)

            tab_res, tab_coin, tab_nocoin, tab_unif = st.tabs([
                "📋 Resultados Detallados",
                "✅ Coincidencias",
                "❌ No Coincidencias",
                "📝 Listado Unificado"
            ])

            with tab_res:
                st.dataframe(st.session_state.df_resultado, use_container_width=True)
            with tab_coin:
                coin_df = st.session_state.df_resultado[st.session_state.df_resultado["COINCIDE"] == "SÍ"]
                st.dataframe(coin_df, use_container_width=True)
            with tab_nocoin:
                nocoin_df = st.session_state.df_resultado[st.session_state.df_resultado["COINCIDE"] == "NO"]
                st.dataframe(nocoin_df, use_container_width=True)
            with tab_unif:
                if st.session_state.df_unificado is not None:
                    st.dataframe(st.session_state.df_unificado, use_container_width=True)

            st.markdown("---")
            st.markdown("#### 💾 Exportar Resultados")
            col1, col2 = st.columns(2)
            with col1:
                if st.button("📊 Exportar Reporte Completo (Excel)", use_container_width=True):
                    excel_bytes = st.session_state.stats.generar_excel(
                        st.session_state.df_resultado,
                        st.session_state.df_unificado
                    )
                    st.download_button(
                        label="⬇️ Descargar Excel",
                        data=excel_bytes,
                        file_name=f"REPORTE_SIVREP_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
            with col2:
                if st.session_state.df_unificado is not None:
                    buffer = io.BytesIO()
                    st.session_state.df_unificado.to_excel(buffer, index=False, sheet_name="UNIFICADO")
                    st.download_button(
                        label="⬇️ Descargar Listado Unificado",
                        data=buffer.getvalue(),
                        file_name=f"LISTADO_UNIFICADO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )



def render_ocr():
    """Pestaña de OCR para planos eléctricos."""
    st.markdown("<div class='sub-header'>⚡ Extracción OCR de Planos Eléctricos</div>", unsafe_allow_html=True)

    if not TESSERACT_OK:
        st.error("🔴 Tesseract OCR no está disponible. En Streamlit Cloud, OCR requiere instalación del motor Tesseract en el sistema operativo, lo cual no está soportado en la versión gratuita.")
        st.info("💡 Alternativa: Use la versión local de SIVREP para procesar planos con OCR completo.")
        return

    if not PYMUPDF_OK:
        st.warning("🟡 PyMuPDF no instalado. La conversión de PDF a imágenes puede no funcionar.")

    st.markdown("#### 📁 Cargar Plano Eléctrico")
    file_plano = st.file_uploader("PDF o Imagen (PNG, JPG, TIFF)", type=['pdf', 'png', 'jpg', 'jpeg', 'tiff'], key="plano")

    if file_plano:
        st.success(f"✅ Archivo cargado: {file_plano.name}")

        if st.button("▶ PROCESAR OCR", type="primary", use_container_width=True):
            with st.spinner("Procesando OCR... Esto puede tomar varios minutos para archivos grandes."):
                try:
                    ext = Path(file_plano.name).suffix.lower()
                    if ext == '.pdf':
                        resultados = st.session_state.motor_ocr.procesar_pdf(file_plano)
                    else:
                        # Guardar imagen temporal
                        img_path = Path(CONFIG.TEMP_DIR) / file_plano.name
                        with open(img_path, "wb") as f:
                            f.write(file_plano.getvalue())
                        resultados = st.session_state.motor_ocr.procesar_imagen(str(img_path))

                    st.session_state.df_ocr = pd.DataFrame(resultados)
                    st.success(f"✅ OCR completado: {len(resultados)} componentes detectados")
                except Exception as e:
                    st.error(f"Error en OCR: {e}")
                    logger.error(traceback.format_exc())

    if st.session_state.df_ocr is not None and len(st.session_state.df_ocr) > 0:
        st.markdown("---")
        st.markdown("#### 📋 Componentes Detectados")

        col1, col2, col3 = st.columns(3)
        with col1:
            tags_ok = (st.session_state.df_ocr['TAG'] != 'NO_IDENTIFICADO').sum()
            st.markdown(f"""
            <div class='metric-card'>
                <div class='metric-label'>TAGs Identificados</div>
                <div class='metric-value status-ok'>{tags_ok}</div>
            </div>
            """, unsafe_allow_html=True)
        with col2:
            modelos_ok = (st.session_state.df_ocr['MODELO'] != 'NO_IDENTIFICADO').sum()
            st.markdown(f"""
            <div class='metric-card'>
                <div class='metric-label'>Modelos Identificados</div>
                <div class='metric-value status-ok'>{modelos_ok}</div>
            </div>
            """, unsafe_allow_html=True)
        with col3:
            st.markdown(f"""
            <div class='metric-card'>
                <div class='metric-label'>Total Componentes</div>
                <div class='metric-value'>{len(st.session_state.df_ocr)}</div>
            </div>
            """, unsafe_allow_html=True)

        st.dataframe(st.session_state.df_ocr, use_container_width=True)

        # Log de procesamiento
        with st.expander("📝 Log de Procesamiento"):
            for _, row in st.session_state.df_ocr.head(50).iterrows():
                st.text(f"TAG: {row['TAG']:<22} | MODELO: {row['MODELO']:<20} | MÉTODO: {row['METODO_EXTRACCION']}")

        # Exportar OCR
        st.markdown("---")
        if st.button("💾 Exportar Resultados OCR (Excel)", use_container_width=True):
            buffer = io.BytesIO()
            st.session_state.df_ocr.to_excel(buffer, index=False, sheet_name="COMPONENTES")
            st.download_button(
                label="⬇️ Descargar Excel OCR",
                data=buffer.getvalue(),
                file_name=f"OCR_PLANOS_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


def render_estadisticas():
    """Pestaña de estadísticas y reportes."""
    st.markdown("<div class='sub-header'>📈 Estadísticas y Reportes Institucionales</div>", unsafe_allow_html=True)

    if st.session_state.metricas is None:
        st.info("Ejecute una validación de repuestos para generar estadísticas.")
        return

    m = st.session_state.metricas

    # Métricas principales
    st.markdown("#### 📊 Métricas de Desempeño")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>Total Origen</div>
            <div class='metric-value'>{m['total_origen']:,}</div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>Total Destino</div>
            <div class='metric-value'>{m['total_destino']:,}</div>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>Comparados</div>
            <div class='metric-value'>{m['comparados']:,}</div>
        </div>
        """, unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>Coincidencias</div>
            <div class='metric-value status-ok'>{m['coincidencias']:,}</div>
        </div>
        """, unsafe_allow_html=True)
    with c5:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>Tasa Éxito</div>
            <div class='metric-value'>{m['tasa_exito']:.2f}%</div>
        </div>
        """, unsafe_allow_html=True)
    with c6:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>Cobertura</div>
            <div class='metric-value'>{m['cobertura']:.2f}%</div>
        </div>
        """, unsafe_allow_html=True)

    # Scores
    st.markdown("#### 📉 Distribución de Scores")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>Score Promedio</div>
            <div class='metric-value'>{m['score_promedio']:.2f}</div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>Score Mínimo</div>
            <div class='metric-value'>{m['score_min']:.2f}</div>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class='metric-card'>
            <div class='metric-label'>Score Máximo</div>
            <div class='metric-value'>{m['score_max']:.2f}</div>
        </div>
        """, unsafe_allow_html=True)

    # Gráficos interactivos
    if PLOTLY_OK and st.session_state.stats.reporte:
        st.markdown("---")
        st.markdown("#### 📈 Gráficos Interactivos")
        fig = st.session_state.stats.generar_graficos_plotly()
        if fig:
            st.plotly_chart(fig, use_container_width=True)

    # Distribución de métodos
    if "metodos" in m:
        st.markdown("---")
        st.markdown("#### 🔍 Distribución de Métodos de Coincidencia")
        met_df = pd.DataFrame(list(m["metodos"].items()), columns=["Método", "Cantidad"])
        st.bar_chart(met_df.set_index("Método"))

    # Reporte JSON
    st.markdown("---")
    st.markdown("#### 📋 Reporte JSON Completo")
    with st.expander("Ver reporte JSON"):
        st.json(m)

    # Exportar reporte
    st.markdown("---")
    if st.button("💾 Exportar Reporte Completo (Excel)", use_container_width=True):
        if st.session_state.df_resultado is not None:
            excel_bytes = st.session_state.stats.generar_excel(
                st.session_state.df_resultado,
                st.session_state.df_unificado
            )
            st.download_button(
                label="⬇️ Descargar Excel",
                data=excel_bytes,
                file_name=f"REPORTE_SIVREP_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


# ═══════════════════════════════════════════════════════════════════════════════
# PUNTO DE ENTRADA
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    st.set_page_config(
        page_title="SIVREP Cloud — Validación de Repuestos",
        page_icon="🔧",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    init_session_state()
    render_header()
    pagina = render_sidebar()

    if pagina == "🏠 Inicio":
        render_home()
    elif pagina == "📊 Validación de Repuestos":
        render_validacion()
    elif pagina == "⚡ Planos Eléctricos (OCR)":
        render_ocr()
    elif pagina == "📈 Estadísticas y Reportes":
        render_estadisticas()

    st.markdown("""
    <div class='footer'>
        SIVREP Cloud v3.1.0 | Departamento de Ingeniería<br>
        Sistema Institucional de Validación de Repuestos y Planos Eléctricos
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
