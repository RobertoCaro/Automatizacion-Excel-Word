#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
extraccion_planos_8_procesos.py

Pipeline de 8 procesos por PDF:

  Detección de recuadros
  1) Proceso 1 → Canvas vectorial (pdfplumber) → detectar recuadros → Recuadros_1
  2) Proceso 2 → Imagen raster (pdf2image) → detectar recuadros → Recuadros_2

  Extracción de texto
  3) Proceso 3 → OCR (Tesseract) en Recuadros_1
  4) Proceso 4 → pdfplumber en Recuadros_1
  5) Proceso 5 → OCR (Tesseract) en Recuadros_2
  6) Proceso 6 → pdfplumber en Recuadros_2

  Consolidación
  7) Proceso 7 → Comparar OCR vs pdfplumber para cada recuadro (por set)
  8) Proceso 8 → Clasificar (TÍTULO / CÓDIGO / REVISIÓN), guardar imágenes y Excel

Requisitos:
  pip install opencv-python numpy pandas pdfplumber pdf2image pytesseract pillow tqdm

Nota:
  - Asegúrate de tener Tesseract instalado y accesible en PATH
  - En Windows, si es necesario: pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
"""

import os
import re
import cv2
import math
import json
import numpy as np
import pdfplumber
import pandas as pd
import pytesseract
import datetime
from tqdm import tqdm
from pdf2image import convert_from_path
from collections import defaultdict

# ---------------- CONFIGURACIÓN ----------------
carpeta_pdf = r'D:'
carpeta_dst = r"D:\Resultados"
#carpeta_pdf = r'C:\Users\Roberto Caro\Desktop\tipico auto'
#carpeta_dst = r"C:\Users\Roberto Caro\Desktop\Licitacion OOEE\macro test\Plantilla Macro\test2"
os.makedirs(carpeta_dst, exist_ok=True)

dpi = 400  # alta resolución para imágenes nítidas
zona_relativa = (0.65, 0.80)  # (x_frac, y_frac) esquina inferior derecha
max_recuadros = 20
area_minima_rel = 0.001
area_max_rel = 0.4
revisiones_validas = ['A','B','N','P','Q','R','S','T','AB'] + [str(i) for i in range(0,10)]
palabras_prohibidas = ['CODELCO','CONTRATISTA','VICEPRESIDENCIA','NOTA IMPORTANTE','FECHA']

# Grosor de líneas al dibujar canvas vectorial (relativo al ancho)
LINE_THICKNESS_REL = 0.002

def display(img, frameName="OpenCV Image"):
    h, w = img.shape[0:2]
    neww = 1200#

    newh = int(neww*(h/w))
    img = cv2.resize(img, (neww, newh))
    cv2.imshow(frameName, img)
    cv2.waitKey(0)
    cv2.destroyAllWindows()

# ---------------- UTILIDADES ----------------
def ensure_int(v): return int(round(v))

def distancia_metric(x, y, w, h, W, H):
    dx = W - (x + w)
    dy = H - (y + h)
    return dx**2 + 4 * dy**2

def es_codigo_valido(txt):
    txt = txt.strip()
    if len(txt) < 5: return False
    if ' ' in txt: return False
    if '-' not in txt: return False
    if not any(c.isalpha() for c in txt): return False
    if not any(c.isdigit() for c in txt): return False
    return True



# ---------------- PREPROCESADO IMAGEN ----------------
def binarizar_para_lineas(img):
    """Prepara dos imágenes: una binarizada normal y otra invertida para detección de líneas."""
    gris = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # Normalizar brillo y contraste
    gris = cv2.equalizeHist(gris)
    
    # Umbral adaptativo
    binaria_normal = cv2.adaptiveThreshold(
        gris, 255,
        cv2.ADAPTIVE_THRESH_MEAN_C,
        cv2.THRESH_BINARY,
        21,  # tamaño de bloque
        5   # constante
    )
    
    # Versión invertida solo para detección de líneas
    binaria_invertida = cv2.bitwise_not(binaria_normal)
    
    # Cerrar huecos en ambas
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    binaria_normal = cv2.morphologyEx(binaria_normal, cv2.MORPH_CLOSE, kernel, iterations=1)
    binaria_invertida = cv2.morphologyEx(binaria_invertida, cv2.MORPH_CLOSE, kernel, iterations=1)
    
    return binaria_normal
# ---------------- DIBUJAR CANVAS VECTORIAL (pdfplumber) ----------------

def es_linea_recta(x0, y0, x1, y1, tolerancia=3):
    """Devuelve True si la línea es casi horizontal o vertical."""
    return abs(y0 - y1) <= tolerancia or abs(x0 - x1) <= tolerancia

def dibujar_lineas_pdf_a_canvas(page, img_w, img_h):
    """Dibuja líneas/rects/edges/curves de pdfplumber sobre canvas blanco."""
    canvas = np.ones((img_h, img_w, 3), dtype=np.uint8) * 255
    thickness = max(1, ensure_int(LINE_THICKNESS_REL * img_w))

    def mapx(x): return ensure_int(x / page.width * img_w)
    def mapy(y): return ensure_int((1 - y / page.height) * img_h)

    # pages have: lines, rects, edges, curves (algunos PDFs no traen curves)
    pools = [page.lines, page.rects, page.edges]
    if hasattr(page, "curves"): pools.append(page.curves)

    for group in pools:
        for el in group:
            if all(k in el for k in ('x0','y0','x1','y1')):
                x0 = mapx(el['x0']); y0 = mapy(el['y0'])
                x1 = mapx(el['x1']); y1 = mapy(el['y1'])
                if es_linea_recta(x0, y0, x1, y1):
                    cv2.line(canvas, (x0, y0), (x1, y1), (0, 0, 0), thickness)



    # cierre para unir pequeños gaps
    gray = cv2.cvtColor(canvas, cv2.COLOR_RGB2GRAY)
    k = cv2.getStructuringElement(cv2.MORPH_RECT, (3,3))
    closed = cv2.morphologyEx(gray, cv2.MORPH_CLOSE, k, iterations=2)
    return cv2.cvtColor(closed, cv2.COLOR_GRAY2RGB)

# ---------------- ROI (esquina inferior derecha) ----------------
def recortar_roi(img, zona_rel):
    H, W = img.shape[:2]
    x0 = ensure_int(W * zona_rel[0])
    y0 = ensure_int(H * zona_rel[1])
    return img[y0:H, x0:W], (x0, y0), (W, H)

def detectar_recuadros_en_zona(bin_img_global, zona_rel, min_area_rel, max_area_rel, max_cands):
    """Detecta recuadros usando minAreaRect para tolerar bordes redondeados o rotados."""
    H, W = bin_img_global.shape[:2]
    x0 = ensure_int(W * zona_rel[0])
    y0 = ensure_int(H * zona_rel[1])
    zona = bin_img_global[y0:, x0:]
    area_z = max(1, zona.shape[0] * zona.shape[1])

    # Preprocesamiento
    blurred = cv2.GaussianBlur(zona, (3, 3), 0)
    edges = cv2.Canny(blurred, 30, 100)
    k = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
    #dilated = cv2.dilate(edges, k, iterations=1)
    #closed = cv2.morphologyEx(dilated, cv2.MORPH_CLOSE, k, iterations=2)

    #cnts, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    cnts, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    cand = []
    for c in cnts:
        rect = cv2.minAreaRect(c)
        box = cv2.boxPoints(rect)
        box = np.array(box, dtype=int)
        x, y, w, h = cv2.boundingRect(box)
        
        x_abs = x0 + x
        y_abs = y0 + y
        x1_abs = min(x_abs + w, W - 1)
        y1_abs = min(y_abs + h, H - 1)

        # Recalcular ancho y alto recortados
        w = x1_abs - x_abs
        h = y1_abs - y_abs

        # Verificar que el área siga siendo válida
        area = w * h
        aspect_ratio = max(w, h) / (min(w, h) + 1e-5)

        if min_area_rel * area_z <= area <= max_area_rel * area_z and aspect_ratio <= 20:
            cand.append((x_abs, y_abs, w, h))




    # Filtrar contenedores
    filtered = []
    rects = [(c[0], c[1], c[0]+c[2], c[1]+c[3]) for c in cand]
    for i, r in enumerate(rects):
        contains_other = False
        for j, s in enumerate(rects):
            if i == j: continue
            if r[0] <= s[0] and r[1] <= s[1] and r[2] >= s[2] and r[3] >= s[3]:
                contains_other = True
                break
        if not contains_other:
            filtered.append(cand[i])

    ordered = sorted(filtered, key=lambda c: distancia_metric(c[0], c[1], c[2], c[3], W, H))[:max_cands]
    return ordered



import math


def agrupar_por_linea(letras, tolerancia_vertical=0.5):
    """Agrupa letras en líneas según su posición vertical."""
    letras_ordenadas = sorted(letras, key=lambda w: (float(w['top']), float(w['x0'])))
    lineas = []
    cur = [letras_ordenadas[0]]
    centro_y = (float(cur[0]['top']) + float(cur[0]['bottom'])) / 2
    altura = float(cur[0]['bottom']) - float(cur[0]['top'])

    for w in letras_ordenadas[1:]:
        centro_w = (float(w['top']) + float(w['bottom'])) / 2
        if abs(centro_w - centro_y) <= tolerancia_vertical * altura:
            cur.append(w)
        else:
            lineas.append(cur)
            cur = [w]
            centro_y = (float(w['top']) + float(w['bottom'])) / 2
            altura = float(w['bottom']) - float(w['top'])

    if cur:
        lineas.append(cur)
    return lineas

def distancia_entre_centros(a, b):
    """Calcula la distancia euclidiana entre los centros de dos letras."""
    cx_a = (float(a['x0']) + float(a['x1'])) / 2
    cy_a = (float(a['top']) + float(a['bottom'])) / 2
    cx_b = (float(b['x0']) + float(b['x1'])) / 2
    cy_b = (float(b['top']) + float(b['bottom'])) / 2
    return np.sqrt((cx_b - cx_a)**2 + (cy_b - cy_a)**2)


def unir_letras_por_distancia(letras):
    """Une letras en palabras usando criterios adaptativos, excluyendo letras angostas como 'I'."""
    if not letras:
        return ""

    letras.sort(key=lambda z: float(z['x0']))

    # Filtrar letras angostas (por ejemplo, 'I' o símbolos estrechos)
    letras_filtradas = [l for l in letras if l['text'].strip() not in ['I', 'l', '|', '1','(',')','/','-',',','.',';','"']]

    # Estadísticas de tamaño
    anchos = [float(l['x1']) - float(l['x0']) for l in letras_filtradas if 'x1' in l and 'x0' in l]
    ancho_med = np.median(anchos) if anchos else 10

    # Espacios horizontales entre letras
    espacios = [
        float(actual['x0']) - float(prev['x1']) if 'x1' in prev else float(actual['x0']) - float(prev['x0'])
        for prev, actual in zip(letras, letras[1:])
    ]
    if not espacios:
        return letras[0]['text']

    # Usar percentil alto para detectar espacios entre palabras
    umbral = np.percentile(espacios, 75)

    resultado = letras[0]['text']
    for i in range(1, len(letras)):
        prev = letras[i - 1]
        actual = letras[i]
        dx = float(actual['x0']) - float(prev['x1']) if 'x1' in prev else float(actual['x0']) - float(prev['x0'])

        if dx > umbral and dx > ancho_med * 0.25:
            resultado += " " + actual['text']
        else:
            resultado += actual['text']
    return resultado


# ---------------- TEXTO: pdfplumber ----------------
def pdf_words_y_lineas_en_rect(page, rect_img, img_w, img_h):
    """Convierte rect px→PDF y extrae words + líneas simples (orden por top,x0)."""
    x, y, w, h = rect_img

    x0_pdf = x / float(img_w) * page.width
    x1_pdf = (x + w) / float(img_w) * page.width
    y0_pdf = (float(y) / img_h) * page.height
    y1_pdf = (float(y + h) / img_h) * page.height

    bbox = (x0_pdf, y0_pdf, x1_pdf, y1_pdf)
    words = page.within_bbox(bbox).extract_words(extra_attrs=['size','x0','x1','top','bottom'])



    from collections import Counter

    alturas = [float(w['bottom']) - float(w['top']) for w in words]
    if(len(alturas)>0):
        altura_maxima = max(alturas)
    else:
        altura_maxima=100
    rango_min = 0.8 * altura_maxima
    rango_max = 1.05 * altura_maxima

    words_filtrados = [
        w for w in words
        if rango_min <= (float(w['bottom']) - float(w['top'])) <= rango_max
    ]

    if not words_filtrados:
        return [], [], []


    lineas_letras = agrupar_por_linea(words_filtrados)
    lineas_texto = [unir_letras_por_distancia(linea) for linea in lineas_letras]

    words_sorted = sorted(words_filtrados, key=lambda w: (float(w['top']), float(w['x0'])))
    lineas = []

    cur = [words_sorted[0]]
    altura_media_linea = float(words_sorted[0]['bottom']) - float(words_sorted[0]['top'])
    centro_y_linea = (float(words_sorted[0]['top']) + float(words_sorted[0]['bottom'])) / 2

    for w in words_sorted[1:]:
        altura_palabra = float(w['bottom']) - float(w['top'])
        centro_y_palabra = (float(w['top']) + float(w['bottom'])) / 2

        tol_vertical = 0.5 * altura_media_linea
        tol_altura = 0.25 * altura_media_linea

        misma_altura = abs(altura_palabra - altura_media_linea) <= tol_altura
        mismo_centro = abs(centro_y_palabra - centro_y_linea) <= tol_vertical

        if misma_altura and mismo_centro:
            cur.append(w)
            alturas = [float(p['bottom']) - float(p['top']) for p in cur]
            centros = [(float(p['top']) + float(p['bottom'])) / 2 for p in cur]
            altura_media_linea = sum(alturas) / len(alturas)
            centro_y_linea = sum(centros) / len(centros)
        else:
            cur.sort(key=lambda z: float(z['x0']))
            lineas.append(unir_letras_por_distancia(cur))
            cur = [w]
            altura_media_linea = altura_palabra
            centro_y_linea = centro_y_palabra

    if cur:
        cur.sort(key=lambda z: float(z['x0']))
        lineas.append(unir_letras_por_distancia(cur))


        

    return words, lineas, bbox

# ---------------- TEXTO: OCR ----------------
def ocr_lineas_en_rect(img_rgb_global, rect_img):
    """OCR en recorte, agrupando por top simple."""
    x, y, w, h = rect_img
    crop = img_rgb_global[y:y+h, x:x+w]
    if crop.size == 0:
        return []

    gray = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY)
    # Upscale y binarización
    up = cv2.resize(gray, (ensure_int(gray.shape[1]*2), ensure_int(gray.shape[0]*2)), interpolation=cv2.INTER_CUBIC)
    _, th = cv2.threshold(up, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    data = pytesseract.image_to_data(th, output_type=pytesseract.Output.DICT,
                                     lang='spa+eng', config='--psm 6')
    words = []
    for i, txt in enumerate(data['text']):
        txt = txt.strip()
        if txt:
            words.append({
                'text': txt,
                'left': data['left'][i],
                'top': data['top'][i],
                'width': data['width'][i],
                'height': data['height'][i],
            })
    if not words:
        return []

    words_sorted = sorted(words, key=lambda w: (w['top'], w['left']))
    lineas = []
    cur_top = words_sorted[0]['top']; cur = [words_sorted[0]['text']]
    for w in words_sorted[1:]:
        if abs(w['top'] - cur_top) <= max(3, 0.2 * cur_top):
            cur.append(w['text'])
        else:
            lineas.append(" ".join(cur))
            cur = [w['text']]
            cur_top = w['top']
    if cur:
        lineas.append(" ".join(cur))

    return lineas

# ---------------- CLASIFICACIÓN ----------------
def clasificar_lineas(lineas):
    """Devuelve (titulo:list|None, codigos:list, revision:str)."""
    textos = [l.strip() for l in lineas if l.strip()]
    text_join = " ".join(textos).upper()

    codigos = [c for c in re.findall(r"\b[\w\-]{5,}\b", text_join) if es_codigo_valido(c)]
    revision = next((t for t in textos if t.strip().upper() in revisiones_validas), "")

    titulo = None
    if len(textos) == 4 and 'MINA' in text_join and not any(p in text_join for p in palabras_prohibidas):
        titulo = textos

    return titulo, codigos, revision

# ---------------- CONSOLIDACIÓN ----------------
def elegir_mejor_texto(lineas_ocr, lineas_pdf):
    """Estrategia simple: prioriza pdfplumber si no está vacío; si no, OCR."""
    if lineas_pdf and any(l.strip() for l in lineas_pdf):
        fuente = "PDF"
        lineas = lineas_pdf
    elif lineas_ocr and any(l.strip() for l in lineas_ocr):
        fuente = "OCR"
        lineas = lineas_ocr
    else:
        fuente = "VACIO"
        lineas = []
    return fuente, lineas

# ---------------- PROCESAR UN PDF ----------------
def procesar_pdf(ruta_pdf):
    base = os.path.basename(ruta_pdf)
    nombre = os.path.splitext(base)[0]
    out_dir = os.path.join(carpeta_dst, nombre)
    os.makedirs(out_dir, exist_ok=True)

    resultados_fila = {
        "PDF": base,
        "TITULO_P1": "", "REVISION_P1": "",
        "CODIGO_1_P1": "", "CODIGO_2_P1": "", "CODIGO_3_P1": "",
        "TITULO_P2": "", "REVISION_P2": "",
        "CODIGO_1_P2": "", "CODIGO_2_P2": "", "CODIGO_3_P2": "",
    }

    with pdfplumber.open(ruta_pdf) as pdf:
        page = pdf.pages[0]
        page_w, page_h = page.width, page.height
        img_w = ensure_int((page_w / 72.0) * dpi)
        img_h = ensure_int((page_h / 72.0) * dpi)

        pil_img = convert_from_path(ruta_pdf, dpi=dpi, first_page=1, last_page=1)[0]
        img_rgb_full = cv2.cvtColor(np.array(pil_img), cv2.COLOR_BGR2RGB)
        

        # Proceso 1: Vectorial
        canvas = dibujar_lineas_pdf_a_canvas(page, img_w, img_h)
        bin_vec = binarizar_para_lineas(canvas)
        recs1 = detectar_recuadros_en_zona(bin_vec, zona_relativa, area_minima_rel, area_max_rel, max_recuadros)


        ann1 = cv2.cvtColor(bin_vec, cv2.COLOR_GRAY2RGB)
        for i,(x,y,w,h) in enumerate(recs1,1):
            cv2.rectangle(ann1,(x,y),(x+w,y+h),(0,255,0),4)
            cv2.putText(ann1,f"{i}",(x+3,y+max(12,int(h*0.5))),cv2.FONT_HERSHEY_SIMPLEX,3,(0,0,255),4)
        cv2.imwrite(os.path.join(out_dir, f"p1_recuadros.png"), ann1)

        # Proceso 2: Raster
        bin_ras = binarizar_para_lineas(img_rgb_full)
        recs2 = detectar_recuadros_en_zona(bin_ras, zona_relativa, area_minima_rel, area_max_rel, max_recuadros)


        # Proceso 3: OCR en Recuadros_1
        ocr_p1 = [ocr_lineas_en_rect(img_rgb_full, rect) for rect in recs1]

        # Proceso 4: pdfplumber en Recuadros_1
        pdf_p1 = [pdf_words_y_lineas_en_rect(page, rect, img_w, img_h)[1] for rect in recs1]

        # Proceso 5: OCR en Recuadros_2
        ocr_p2 = [ocr_lineas_en_rect(img_rgb_full, rect) for rect in recs2]

        # Proceso 6: pdfplumber en Recuadros_2
        pdf_p2 = [pdf_words_y_lineas_en_rect(page, rect, img_w, img_h)[1] for rect in recs2]
        

        # ✅ Validación de longitud
        while len(ocr_p1) < len(recs1):
            ocr_p1.append([])
        while len(pdf_p1) < len(recs1):
            pdf_p1.append([])
        while len(ocr_p2) < len(recs2):
            ocr_p2.append([])
        while len(pdf_p2) < len(recs2):
            pdf_p2.append([])

        # Proceso 7: Comparar OCR vs PDF
        best_p1 = [elegir_mejor_texto(ocr_p1[i], pdf_p1[i]) for i in range(len(recs1))]
        best_p2 = [elegir_mejor_texto(ocr_p2[i], pdf_p2[i]) for i in range(len(recs2))]

        # Proceso 8: Clasificar
        hall_p1 = {}
        for fuente, lineas in best_p1:
            titulo, cods, rev = clasificar_lineas(lineas)
            if titulo and not hall_p1.get("TITULO_P1"):
                hall_p1["TITULO_P1"] = " | ".join(titulo)
            if rev and not hall_p1.get("REVISION_P1"):
                hall_p1["REVISION_P1"] = rev
            for c in cods:
                for idx in range(1, 4):
                    k = f"CODIGO_{idx}_P1"
                    if not hall_p1.get(k):
                        hall_p1[k] = c
                        break

        hall_p2 = {}
        for fuente, lineas in best_p2:
            titulo, cods, rev = clasificar_lineas(lineas)
            if titulo and not hall_p2.get("TITULO_P2"):
                hall_p2["TITULO_P2"] = " | ".join(titulo)
            if rev and not hall_p2.get("REVISION_P2"):
                hall_p2["REVISION_P2"] = rev
            for c in cods:
                for idx in range(1, 4):
                    k = f"CODIGO_{idx}_P2"
                    if not hall_p2.get(k):
                        hall_p2[k] = c
                        break

        for k in hall_p1:
            resultados_fila[k] = hall_p1[k]
        for k in hall_p2:
            resultados_fila[k] = hall_p2[k]

        # Separar título en líneas
        titulo_completo = resultados_fila.get("TITULO_P1", "").strip() or resultados_fila.get("TITULO_P2", "").strip()
        partes = [p.strip() for p in titulo_completo.split('|') if p.strip()]
        while len(partes) < 4:
            partes.append('')
        for i in range(4):
            resultados_fila[f'TITULO_LINEA_{i+1}'] = partes[i]

        # Guardar CSV individual
        #pd.DataFrame([resultados_fila]).to_csv(
        #    os.path.join(out_dir, f"resultado_{nombre}.csv"),
        #    index=False,
        #    sep=';',
        #    encoding='utf-8'
        #)
    #print(nombre+"--------------------------------------------------------")
    #print(resultados_fila)
    
    return resultados_fila

import os
import datetime
import pandas as pd
from tqdm import tqdm

# ---------------- MAIN ----------------
def main():
    pdfs = []
    for root, _, files in os.walk(carpeta_pdf):
        for f in files:
            if f.lower().endswith(".pdf"):
                pdfs.append(os.path.join(root, f))

    print(f"Se encontraron {len(pdfs)} PDFs en: {carpeta_pdf}")

    resultados_global = []
    errores = []

    for ruta in tqdm(pdfs, desc="Procesando PDFs"):
        try:
            res = procesar_pdf(ruta)
            resultados_global.append(res)
        except Exception as e:
            error_msg = f"{os.path.basename(ruta)}: {type(e).__name__} → {e}"
            print(f"⚠️ Error procesando {error_msg}")
            errores.append({"archivo": os.path.basename(ruta), "error": str(e)})

    if resultados_global:
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        xlsx = os.path.join(carpeta_dst, f"resultados_global_{timestamp}.xlsx")
        #csv = os.path.join(carpeta_dst, f"resultados_global_{timestamp}.csv")

        df_resultados = pd.DataFrame(resultados_global)
        df_resultados.to_excel(xlsx, index=False)
        #df_resultados.to_csv(csv, index=False, sep=';', encoding='utf-8')

        print("✅ Resultados guardados en:")
        print("   📄 Excel:", xlsx)
        #print("   📄 CSV:", csv)
    else:
        print("No se extrajeron resultados.")

    #if errores:
        #log_errores = os.path.join(carpeta_dst, f"errores_{timestamp}.csv")
        #pd.DataFrame(errores).to_csv(log_errores, index=False, encoding='utf-8')
        #print("⚠️ Errores guardados en:", log_errores)

if __name__ == "__main__":
    main()
