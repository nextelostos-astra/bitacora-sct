import streamlit as pd_stream
import io
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages

# Configuración de la página web
pd_stream.set_page_config(page_title="Generador de Bitácoras NOM-087", layout="centered")

pd_stream.title("📊 Extractor Automatizado de Horas de Servicio")
pd_stream.subheader("Normativa NOM-087-SCT-2017")
pd_stream.write("Sube tu archivo de telemetría en Excel para generar los reportes unificados por unidad.")

# ==========================================
# FUNCIONES AUXILIARES DE CONVERSIÓN
# ==========================================
def extraer_hora_decimal(datetime_obj):
    if pd.isna(datetime_obj):
        return None
    if isinstance(datetime_obj, str):
        datetime_obj = datetime_obj.split(".")
        datetime_obj = pd.to_datetime(datetime_obj)
    return datetime_obj.hour + (datetime_obj.minute / 60.0) + (getattr(datetime_obj, 'second', 0) / 3600.0)

def calcular_minutos(inicio, fin):
    if fin < inicio: fin += 24.0
    return int(round((fin - inicio) * 60.0))

def minutos_a_hhmm(minutos):
    horas = int(minutos // 60)
    mins = int(minutos % 60)
    return f"{horas:02d}:{mins:02d}"

# ==========================================
# INTERFAZ DE CARGA DE ARCHIVOS
# ==========================================
archivo_cargado = pd_stream.file_uploader("Selecciona el archivo 'reportehos.xlsx'", type=["xlsx"])

if archivo_cargado is not None:
    try:
        # Cargar la hoja 'Data' desde el archivo subido por el usuario
        df_excel = pd.read_excel(archivo_cargado, sheet_name="Data", skiprows=11, header=None, engine="openpyxl")

        idx_unidad = 0
        idx_nombre = 5
        idx_apellido = 6
        idx_inicio = 13
        idx_fin = 15
        idx_distancia = 16
        mapeo_estados = {'M': 0, 'C': 1, 'D': 2, 'F': 3}

        # Limpieza de datos inicial
        df_excel = df_excel.dropna(subset=[idx_inicio, idx_fin, idx_unidad])

        def obtener_solo_fecha(dt_obj):
            if isinstance(dt_obj, str): dt_obj = pd.to_datetime(dt_obj.split("."))
            return dt_obj.strftime("%Y-%m-%d")

        df_excel['Fecha_Solo'] = df_excel[idx_inicio].apply(obtener_solo_fecha)
        grupos_unidad = df_excel.groupby(idx_unidad)

        pd_stream.success(f"¡Archivo leído con éxito! Se detectaron **{len(grupos_unidad)} unidades** diferentes.")
        pd_stream.write("---")

        # Procesar cada unidad detectada
        for unidad_id, df_unidad in grupos_unidad:
            unidad_str = str(unidad_id).strip().upper()

            # Buscar el nombre del operador
            nombre_operador = "OPERADOR NO ESPECIFICADO"
            for _, fila in df_unidad.iterrows():
                nom = str(fila.iloc[idx_nombre]).strip() if pd.notna(fila.iloc[idx_nombre]) else ""
                ape = str(fila.iloc[idx_apellido]).strip() if pd.notna(fila.iloc[idx_apellido]) else ""
                if nom or ape:
                    nombre_operador = f"{nom} {ape}".strip().upper()
                    break

            # Crear un contenedor en memoria RAM para guardar el PDF sin escribir en el disco rígido directamente
            buffer_pdf = io.BytesIO()
            grupos_por_dia = sorted(list(df_unidad.groupby('Fecha_Solo')), key=lambda x: x[0])

            with PdfPages(buffer_pdf) as pdf:
                for fecha_evaluada, df_dia in grupos_por_dia:
                    bloques_raw = []
                    minutos_totales_actividad = {"C": 0, "M": 0, "D": 0, "F": 0}

                    for idx_fila, fila in df_dia.iterrows():
                        try: distancia = float(fila.iloc[idx_distancia])
                        except: distancia = 0.0

                        estado_letra = "C" if distancia > 2 else "M"
                        valor_estado = mapeo_estados[estado_letra]
                        h_inicio_dec = extraer_hora_decimal(fila.iloc[idx_inicio])
                        h_fin_dec = extraer_hora_decimal(fila.iloc[idx_fin])

                        if h_inicio_dec is None or h_fin_dec is None: continue
                        if h_fin_dec == h_inicio_dec: h_fin_dec += 1.0 / 60.0

                        bloques_raw.append((h_inicio_dec, h_fin_dec, valor_estado, estado_letra))

                    bloques_raw.sort(key=lambda t: t[0])

                    # Lógica de Rellenos 24 hrs
                    bloques_completos = []
                    if bloques_raw:
                        primer_inicio = bloques_raw[0][0]
                        if primer_inicio > 0.0:
                            minutos_totales_actividad["M"] += calcular_minutos(0.0, primer_inicio)
                            bloques_completos.append((0.0, primer_inicio, mapeo_estados["M"]))

                        for j, (inicio, fin, valor, letra) in enumerate(bloques_raw):
                            if j > 0:
                                fin_anterior = bloques_raw[j-1][1]
                                if inicio > fin_anterior:
                                    minutos_totales_actividad["M"] += calcular_minutos(fin_anterior, inicio)
                                    bloques_completos.append((fin_anterior, inicio, mapeo_estados["M"]))

                            minutos_totales_actividad[letra] += calcular_minutos(inicio, fin)
                            bloques_completos.append((inicio, fin, valor))

                        ultimo_fin = bloques_completos[-1][1]
                        if ultimo_fin < 24.0:
                            minutos_totales_actividad["M"] += calcular_minutos(ultimo_fin, 24.0)
                            bloques_completos.append((ultimo_fin, 24.0, mapeo_estados["M"]))
                    else:
                        minutos_totales_actividad["M"] = 24 * 60
                        bloques_completos.append((0.0, 24.0, mapeo_estados["M"]))

                    # Diseño de Escalones Rectos
                    x, y = [], []
                    for j, (inicio, fin, estado) in enumerate(bloques_completos):
                        if j == 0:
                            x.append(inicio); y.append(estado)
                        else:
                            estado_anterior = y[-1]
                            if estado != estado_anterior:
                                x.append(inicio); y.append(estado_anterior)
                                x.append(inicio); y.append(estado)
                        x.append(fin); y.append(estado)

                    # Dibujar Lienzo SCT
                    fig, ax = plt.subplots(figsize=(12, 4))
                    ax.set_xticks(np.arange(0, 25, 1))
                    ax.set_yticks([0, 1, 2, 3])
                    ax.set_yticklabels(["Maniobras (M)", "Conduciendo (C)", "Durmiendo (D)", "Fuera de Serv. (F)"], fontweight='bold')
                    ax.grid(which='major', color='#555555', linestyle='-', linewidth=1)

                    if x: ax.plot(x, y, color="#0d47a1", linewidth=2.5)

                    str_conduciendo = minutos_a_hhmm(minutos_totales_actividad["C"])
                    str_maniobras   = minutos_a_hhmm(minutos_totales_actividad["M"])
                    str_durmiendo   = minutos_a_hhmm(minutos_totales_actividad["D"])
                    str_fuera_serv  = minutos_a_hhmm(minutos_totales_actividad["F"])
                    str_totales     = minutos_a_hhmm(sum(minutos_totales_actividad.values()))

                    texto_resumen = (
                        f"RESUMEN DIARIO Total 24:00 h:\n"
                        f"• Conduciendo: {str_conduciendo} h\n"
                        f"• Maniobras: {str_maniobras} h\n"
                        f"• Durmiendo: {str_durmiendo} h\n"
                        f"• Fuera de Serv: {str_fuera_serv} h\n"
                        f"Total Control: {str_totales} h"
                    )
                    ax.text(-2.5, -1.6, texto_resumen, fontsize=9, style='italic',
                            bbox={'facecolor': '#f9f9f9', 'alpha': 0.8, 'pad': 6}, transform=ax.transData)

                    ax.set_xlim(0, 24)
                    ax.set_ylim(-0.5, 3.5)
                    ax.set_title(f"UNIDAD: {unidad_str}   |   OPERADOR: {nombre_operador}   |   FECHA: {fecha_evaluada}", fontweight='bold', fontsize=10)

                    plt.tight_layout()
                    pdf.savefig(fig, dpi=300)
                    plt.close(fig)

            # Preparar descarga del archivo en la interfaz web
            buffer_pdf.seek(0)
            nombre_archivo_sanitizado = nombre_operador.replace(" ", "_")
            nombre_descarga = f"Bitacoras_Unidad_{unidad_str}_{nombre_archivo_sanitizado}.pdf"

            # Tarjetas de descarga interactivas para el usuario
            col1, col2 = pd_stream.columns([3, 1])
            with col1:
                pd_stream.write(f"📂 **Unidad:** {unidad_str} — Operador: *{nombre_operador}* ({len(grupos_por_dia)} páginas)")
            with col2:
                pd_stream.download_button(
                    label="📥 Descargar PDF",
                    data=buffer_pdf,
                    file_name=nombre_descarga,
                    mime="application/pdf",
                    key=f"btn_{unidad_str}"
                )

    except Exception as e:
        pd_stream.error(f"Error al procesar el archivo Excel: {e}")

