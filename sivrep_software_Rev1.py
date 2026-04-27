#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SIVREP CLOUD v3.1.1 - Sistema Institucional de Validacion de Repuestos
Plataforma: Streamlit Cloud / Web
"""

import os
import sys
import re
import json
import logging
import io
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Any
import traceback
import time

import streamlit as st
import pandas as pd
import numpy as np

# Configuracion
from dataclasses import dataclass

@dataclass(frozen=True)
class AppConfig:
    APP_NAME: str = "SIVREP CLOUD"
    VERSION: str = "3.1.1"
    ORGANIZATION: str = "Departamento de Ingenieria"
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

# CSS Institucional
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

# Importaciones con manejo graceful
PANDAS_OK = True
PDF_OK = False
PIL_OK = False
TESSERACT_OK = False
CV2_OK = False
FUZZY_OK = False
MATPLOTLIB_OK = False
SEABORN_OK = False
PYMUPDF_OK = False
PLOTLY_OK = False

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    import PyPDF2
    import pdfplumber
    PDF_OK = True
except ImportError as e:
    logger.error("PyPDF2/pdfplumber no disponibles: " + str(e))

try:
    import fitz
    PYMUPDF_OK = True
except ImportError:
    logger.warning("PyMuPDF no instalado.")

try:
    from PIL import Image, ImageEnhance, ImageFilter
    PIL_OK = True
except ImportError as e:
    logger.error("Pillow no disponible: " + str(e))

try:
    import pytesseract
    TESSERACT_OK = True
except ImportError as e:
    logger.error("pytesseract no disponible: " + str(e))

try:
    import cv2
    CV2_OK = True
except ImportError as e:
    logger.error("OpenCV no disponible: " + str(e))

try:
    from fuzzywuzzy import fuzz
    FUZZY_OK = True
except ImportError as e:
    logger.error("fuzzywuzzy no disponible: " + str(e))

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    MATPLOTLIB_OK = True
except ImportError as e:
    logger.error("matplotlib no disponible: " + str(e))

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


# =============================================================================
# UTILIDADES GENERALES
# =============================================================================
class Utils:
    """Herramientas utilitarias estaticas."""

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


# =============================================================================
# MOTOR DE LECTURA DE ARCHIVOS
# =============================================================================
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
        logger.info("Excel cargado: " + str(len(df)) + " filas x " + str(len(df.columns)) + " columnas")
        return df

    def read_pdf_tablas(self, file_obj) -> Optional[pd.DataFrame]:
        if not PDF_OK:
            raise RuntimeError("Librerias PDF no disponibles.")
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
                out = Path(CONFIG.TEMP_DIR) / ("pdf_page_" + str(i).zfill(3) + ".png")
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


# =============================================================================
# MOTOR OCR
# =============================================================================
class OCREngine:
    """Extraccion inteligente de TAGs y Modelos desde planos electricos."""

    def __init__(self):
        self.config_tesseract = "--psm 6 -l " + CONFIG.OCR_LANG
        self.patrones = [
            ("TAG_EXPLICITO", r"TAG\s*[:\-]?\s*([A-Z0-9\-]{3,20})"),
            ("TAG_GENERICO", r"([A-Z]{1,4}[-\.]?\d{2,5}[A-Z0-9\-]*)"),
            ("NPARTE_EXPLICITO", r"N[\u00b0\s]*PARTE\s*[:\-]?\s*([A-Z0-9\-]{3,25})"),
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
            raise ValueError("No se pudo cargar: " + ruta_imagen)
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
            temp = Path(CONFIG.TEMP_DIR) / ("ocr_preproc_" + Path(ruta_imagen).stem + ".png")
            cv2.imwrite(str(temp), img)
            return pytesseract.image_to_string(img, config=self.config_tesseract).upper()
        return pytesseract.image_to_string(Image.open(ruta_imagen), config=self.config_tesseract).upper()

    def extraer_datos(self, texto: str) -> List[Dict[str, str]]:
        resultados = []
        lineas = [l.strip() for l in texto.split("\n") if len(l.strip()) > 2]
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
                resultados.append({"TAG": w1, "MODELO": w2, "TEXTO_ORIGINAL": w1 + " " + w2, "METODO_EXTRACCION": "OCR_BLOQUE"})
                i += 2
                continue
            if Utils.es_codigo_parte(w1) and len(w1) > 5:
                resultados.append({"TAG": w1, "MODELO": w2 if Utils.es_codigo_parte(w2) else "NO_IDENTIFICADO",
                                   "TEXTO_ORIGINAL": w1, "METODO_EXTRACCION": "OCR_AISLADO"})
            i += 1
        return resultados

    def procesar_imagen(self, ruta: str) -> List[Dict[str, str]]:
        logger.info("OCR sobre imagen: " + ruta)
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
                logger.error("Error pagina " + str(idx + 1) + ": " + str(e))
        logger.info("OCR PDF finalizado: " + str(len(todos)) + " componentes.")
        return todos


# =============================================================================
# MOTOR FUZZY MATCHING
# =============================================================================
class SparePartMatcher:
    """Comparacion inteligente con multiples estrategias."""

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
        logger.info("Matcher listo (umbral=" + str(umbral) + ", estrategia=" + estrategia + ").")

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
            return True, round(score, 2), "fuzzy_" + self.estrategia
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
                "COINCIDE": "SI" if coincide else "NO",
                "IDX_DESTINO": best_idx if coincide else -1
            }
            for c in df_origen.columns:
                fila_res["SRC_" + c] = fila[c]
            resultados.append(fila_res)

            if int(idx) % 10 == 0:
                progress = min((int(idx) + 1) / total, 1.0)
                progress_bar.progress(progress)
                status_text.text("Progreso: " + str(int(idx) + 1) + "/" + str(total) + " (" + str(round(progress*100, 1)) + "%)")

        progress_bar.empty()
        status_text.empty()
        self.estadisticas["tiempo_seg"] = time.time() - t0
        logger.info("Comparacion lista en " + str(round(self.estadisticas["tiempo_seg"], 2)) + "s")
        return pd.DataFrame(resultados)


# =============================================================================
# MOTOR ESTADISTICO
# =============================================================================
class StatisticsEngine:
    """Genera metricas, graficos y reportes institucionales."""

    def __init__(self):
        self.reporte: Dict[str, Any] = {}
        logger.info("StatisticsEngine inicializado.")

    def analizar(self, df_resultado: pd.DataFrame, df_origen: pd.DataFrame,
                 df_destino: pd.DataFrame) -> Dict[str, Any]:
        total_o, total_d = len(df_origen), len(df_destino)
        if "COINCIDE" not in df_resultado.columns:
            raise ValueError("DataFrame sin columna COINCIDE")
        coin = df_resultado[df_resultado["COINCIDE"] == "SI"]
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
        """Genera graficos interactivos con Plotly."""
        if not self.reporte:
            return None

        fig = make_subplots(
            rows=2, cols=2,
            subplot_titles=("Distribucion de Resultados", "Metricas de Desempeno",
                           "Metodos de Coincidencia", "Resumen Ejecutivo"),
            specs=[[{"type": "pie"}, {"type": "bar"}],
                   [{"type": "bar"}, {"type": "domain"}]]
        )

        # Pie chart
        sizes = [self.reporte["coincidencias"], self.reporte["no_coincidencias"]]
        fig.add_trace(go.Pie(
            labels=["Coincidencias", "No Coincidencias"],
            values=sizes,
            marker_colors=["#38a169", "#e53e3e"],
            textinfo="label+percent",
            hole=0.4,
        ), row=1, col=1)

        # Barras horizontales
        metricas = ["tasa_exito", "cobertura", "score_promedio"]
        valores = [self.reporte.get(m, 0) for m in metricas]
        nombres = ["Tasa Exito (%)", "Cobertura (%)", "Score Promedio"]
        fig.add_trace(go.Bar(
            x=valores,
            y=nombres,
            orientation='h',
            marker_color=["#3182ce", "#805ad5", "#dd6b20"],
            text=[str(round(v, 1)) for v in valores],
            textposition="outside",
        ), row=1, col=2)

        # Metodos
        if "metodos" in self.reporte:
            met = self.reporte["metodos"]
            fig.add_trace(go.Bar(
                x=list(met.keys()),
                y=list(met.values()),
                marker_color="#2b6cb0",
            ), row=2, col=1)

        # Resumen como indicadores
        fig.add_trace(go.Indicator(
            mode="number",
            value=self.reporte["coincidencias"],
            title={"text": "Coincidencias"},
        ), row=2, col=2)

        fig.update_layout(
            title_text="SIVREP - Analisis Estadistico de Validacion",
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
                df_resultado[df_resultado["COINCIDE"] == "SI"].to_excel(w, sheet_name="COINCIDENCIAS", index=False)
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
            logger.error("Error formateando Excel: " + str(e))


# =============================================================================
# CONSOLIDADOR UNIFICADO
# =============================================================================
class UnifiedConsolidator:
    """Genera un listado maestro unico a partir de dos fuentes."""

    COLUMNAS_PRIORITARIAS = [
        "SISTEMA", "EQUIPO Y SUBEQUIPO", "ESPECIALIDAD", "EQUIPO O REPUESTO",
        "ARTICULO", "FAMILIA", "TAG", "FABRICANTE", "MODELO", "SERIE",
        "N PARTE", "NOMBRE", "DESCRIPCION"
    ]

    def __init__(self):
        logger.info("UnifiedConsolidator inicializado.")

    def consolidar(self, df_o: pd.DataFrame, df_d: pd.DataFrame,
                   df_res: pd.DataFrame, col_o: str, col_d: str) -> pd.DataFrame:
        idx_coin = df_res[df_res["COINCIDE"] == "SI"]["IDX_ORIGEN"].tolist()
        df_o_m = df_o.copy()
        df_o_m["_ORIGEN"] = "ARCHIVO_1"
        df_o_m["_ESTADO"] = df_o_m.index.map(lambda x: "COINCIDE" if x in idx_coin else "UNICO_A1")
        idx_t_usados = df_res[df_res["COINCIDE"] == "SI"]["IDX_DESTINO"].dropna().astype(int).unique()
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
        uni["_ID_UNICO"] = ["REP-" + str(i+1).zfill(6) for i in range(len(uni))]
        logger.info("Consolidado: " + str(len(uni)) + " repuestos unicos.")
        return uni

    def _normalizar_cols(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        mapping = {
            r"SISTEMA|SIST": "SISTEMA",
            r"EQUIPO\s*Y\s*SUBEQUIPO|EQUIPO\s*SUB|EQUIPO(?!\s*O)": "EQUIPO Y SUBEQUIPO",
            r"ESPECIALIDAD|ESP": "ESPECIALIDAD",
            r"EQUIPO\s*O\s*REPUESTO|REPUESTO|EQUIPO\s*REP": "EQUIPO O REPUESTO",
            r"ARTICULO|ARTICULO|ART(?!I)": "ARTICULO",
            r"FAMILIA|FAM": "FAMILIA",
            r"^TAG$|TAG\s*ID": "TAG",
            r"FABRICANTE|FABR|MARCA|MANUFACTURER|MAKER": "FABRICANTE",
            r"MODELO|MODEL(?!O)|MOD(?=ELO)": "MODELO",
            r"SERIE|SER(?!I)|SERIAL": "SERIE",
            r"N[\u00b0\s]*PARTE|PART\s*NO|PART\s*NUMBER|NUMERO\s*PARTE|N\s*PARTE|N\u00b0\s*PARTE|PARTE": "N PARTE",
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
