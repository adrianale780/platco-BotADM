
import streamlit as st
import io
import requests
import openpyxl
from datetime import datetime, timedelta
import os
import re
import unicodedata



# ==========================================
# 🧠 UTILIDADES GENERALES
# ==========================================
def normalizar_texto(texto):
    if not texto: return ""
    texto = str(texto).upper().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def obtener_hoja_flexible(wb, nombre_buscado):
    objetivo = normalizar_texto(nombre_buscado)
    for sheet_name in wb.sheetnames:
        if normalizar_texto(sheet_name) == objetivo:
            return wb[sheet_name]
    return None

def cargar_diccionario_cuentas():
    reglas = []
    ruta_diccionario = os.path.join(os.path.dirname(__file__), "diccionario.xlsx")
    if not os.path.exists(ruta_diccionario): return []
    try:
        wb = openpyxl.load_workbook(ruta_diccionario, read_only=True, data_only=True)
        ws = obtener_hoja_flexible(wb, "CATEGORIA") or wb.active
        for fila in ws.iter_rows(min_row=2, values_only=True):
            if fila[0] and fila[1]: 
                reglas.append((str(fila[0]).strip().upper(), str(fila[1]).strip()))
        wb.close()
        reglas.sort(key=lambda x: len(x[0]), reverse=True)
        return reglas
    except: return []

def cargar_diccionario_areas():
    reglas = []
    ruta_diccionario = os.path.join(os.path.dirname(__file__), "diccionario.xlsx")
    if not os.path.exists(ruta_diccionario): return []
    try:
        wb = openpyxl.load_workbook(ruta_diccionario, read_only=True, data_only=True)
        ws = obtener_hoja_flexible(wb, "AREA")
        if ws:
            for fila in ws.iter_rows(min_row=2, values_only=True):
                if fila[0] and fila[1]: 
                    reglas.append((str(fila[0]).strip().upper(), str(fila[1]).strip()))
        wb.close()
        reglas.sort(key=lambda x: len(x[0]), reverse=True)
        return reglas
    except: return []

# ==========================================
# 🧠 CÁLCULO DE TASAS HISTÓRICAS (CON CLAVE PREMIUM)
# ==========================================
def cargar_tasas_historicas(callback_log):
    memoria_tasas = {}
    URL_HISTORICO = "https://api.dolarvzla.com/public/exchange-rate/list"
    
    # TU CLAVE DE ACCESO
    MI_CLAVE = "ce57d7613b2a0eb8d11aa2d0781206c2758ce66587dddee61f6bcf1026d942c0"
    
    # Preparamos el pase de entrada (Header)
    headers = {
        'x-dolarvzla-key': MI_CLAVE
    }
    
    callback_log(f"🌍 Conectando al histórico oficial (Con llave)...")

    try:
        # Enviamos la petición CON la clave
        response = requests.get(URL_HISTORICO, headers=headers, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            lista_datos = []

            if isinstance(data, dict) and 'rates' in data:
                lista_datos = data['rates']
            elif isinstance(data, list):
                lista_datos = data
            
            count = 0
            for item in lista_datos:
                fecha = item.get('date')
                precio = item.get('usd')

                if fecha and precio:
                    fecha_limpia = str(fecha)[:10]
                    memoria_tasas[fecha_limpia] = float(precio)
                    count += 1
            
            callback_log(f"   ✅ Acceso concedido: {count} fechas históricas cargadas.")
        else:
            # Si la clave falla o expira, avisamos pero no detenemos el programa
            callback_log(f"⚠️ Error de acceso al histórico ({response.status_code}). Usando solo tasa de hoy.")

    except Exception as e:
        callback_log(f"⚠️ Sin conexión al histórico: {str(e)}")

    return memoria_tasas

def formatear_fecha_para_api(valor_celda):
    if not valor_celda: return None
    try:
        if isinstance(valor_celda, datetime): return valor_celda.strftime("%Y-%m-%d")
        texto = str(valor_celda).strip()
        formatos = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%y", "%d-%m-%Y"]
        for fmt in formatos:
            try: return datetime.strptime(texto, fmt).strftime("%Y-%m-%d")
            except: continue
    except: return None
    return None

def buscar_tasa_inteligente(fecha_str, memoria_tasas):
    if not fecha_str: return 0
    if fecha_str in memoria_tasas: return memoria_tasas[fecha_str]
    try:
        dt_obj = datetime.strptime(fecha_str, "%Y-%m-%d")
        for dias_atras in range(1, 6):
            fecha_b = dt_obj - timedelta(days=dias_atras)
            f_key = fecha_b.strftime("%Y-%m-%d")
            if f_key in memoria_tasas: return memoria_tasas[f_key]
    except: pass
    return 0

def limpiar_numero(valor):
    if valor is None: return 0
    if isinstance(valor, (int, float)): return float(valor)
    texto = str(valor).strip()
    if not texto or texto == "-": return 0
    try:
        return float(texto.replace(".", "").replace(",", "."))
    except: return 0

def obtener_nombre_mes_es(mes_num):
    meses = {1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL", 5: "MAYO", 6: "JUNIO",
             7: "JULIO", 8: "AGOSTO", 9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"}
    return meses.get(mes_num, "")

from openpyxl.utils import get_column_letter
from datetime import datetime
import re


# ==========================================
# 🛡️ CÓDIGO ORIGINAL (ANOCHE): RESUMEN SEMANAL
# ==========================================
# Solo procesa Ingresos. Solo usa DATA BS. No toca Excedentes.

def procesar_resumen_semanal(wb, callback_log):
    # 1. CARGA SIMPLE
    try:
        # Buscamos las hojas clave
        ws_resumen = None
        ws_data = None
        
        for sheet in wb.sheetnames:
            if "RESUMEN DISPONIBILIDAD" in sheet.upper():
                ws_resumen = wb[sheet]
            if "DATA BS" in sheet.upper():
                ws_data = wb[sheet]
                
        if not ws_resumen or not ws_data:
            callback_log("⚠️ Error: Faltan hojas (Resumen o Data BS).")
            return 0

    except Exception as e:
        callback_log(f"❌ Error cargando hojas: {str(e)}")
        return 0

    callback_log("📅 Procesando Resumen Semanal (Solo Ingresos)...")

    # 2. DETECTAR SEMANAS
    fila_encabezado = 0
    # Buscamos la fila que dice "SEMANA DEL..."
    for r in range(1, 15):
        for c in range(1, 20): 
            val = str(ws_resumen.cell(row=r, column=c).value).upper()
            if "SEMANA" in val and "DEL" in val:
                fila_encabezado = r
                break
        if fila_encabezado > 0: break
    
    if fila_encabezado == 0: 
        callback_log("⚠️ No encontré encabezados de semana.")
        return 0

    semanas_config = [] 
    for col in range(3, ws_resumen.max_column + 1):
        texto = str(ws_resumen.cell(row=fila_encabezado, column=col).value).upper()
        match = re.search(r"(\d+)\s+AL\s+(\d+)", texto)
        if match:
            dia_inicio, dia_fin = int(match.group(1)), int(match.group(2))
            sub = str(ws_resumen.cell(row=fila_encabezado+1, column=col).value).upper()
            col_est, col_act = 0, 0
            
            if "ACTUALIZACIÓN" in sub or "ACTUALIZACION" in sub:
                col_act, col_est = col, col - 1
            elif "ESTIMADO" in sub:
                col_est, col_act = col, col + 1
            
            if col_est > 0:
                semanas_config.append({
                    'min': dia_inicio, 'max': dia_fin,
                    'col_est': col_est, 'col_act': col_act
                })

    # 3. SUMAR REAL (DATA BS)
    acumulados_real = {i: {} for i in range(len(semanas_config))}
    
    for r in range(4, ws_data.max_row + 1):
        try:
            # Leer Fecha
            raw_fecha = ws_data.cell(row=r, column=2).value
            dia_dato = 0
            if isinstance(raw_fecha, datetime): dia_dato = raw_fecha.day
            elif raw_fecha:
                txt = str(raw_fecha).strip()
                if len(txt) >= 2 and txt[:2].isdigit(): dia_dato = int(txt[:2])
            
            if dia_dato == 0: continue

            # Leer Cuenta y Monto
            cta = str(ws_data.cell(row=r, column=13).value).strip().upper()
            
            # Limpieza básica de número
            val_monto = ws_data.cell(row=r, column=8).value
            monto = 0
            if val_monto:
                if isinstance(val_monto, (int, float)): monto = float(val_monto)
                else:
                    try: monto = float(str(val_monto).replace(",", "").strip())
                    except: pass

            if monto != 0:
                for i, cfg in enumerate(semanas_config):
                    if cfg['min'] <= dia_dato <= cfg['max']:
                        if cta not in acumulados_real[i]: acumulados_real[i][cta] = 0
                        acumulados_real[i][cta] += monto
                        break 
        except: continue

    # 4. ESCRIBIR EN RESUMEN (SOLO INGRESOS)
    cuentas_ingresos = ["CONTINUIDAD OPERATIVA", "SIMCARD", "ALIADOS COMERCIALES"]
    start_row = fila_encabezado + 2 
    cambios = 0
    
    for r in range(start_row, ws_resumen.max_row + 1):
        val_cta = ws_resumen.cell(row=r, column=2).value 
        if not val_cta: continue
        nombre = str(val_cta).strip().upper().replace("\xa0", "")

        if nombre in cuentas_ingresos:
            for i, cfg in enumerate(semanas_config):
                # Sumamos lo real de DATA BS
                monto_real = 0
                for k, v in acumulados_real[i].items():
                    if nombre in k or k in nombre: monto_real += v
                
                # Vemos si hay Estimado manual escrito
                valor_estimado = ws_resumen.cell(row=r, column=cfg['col_est']).value
                
                # Si hay estimado O hubo movimiento real, ponemos la fórmula
                if valor_estimado or monto_real != 0:
                    letra = get_column_letter(cfg['col_est'])
                    # FÓRMULA: =ESTIMADO - REAL
                    ws_resumen.cell(row=r, column=cfg['col_act']).value = f"={letra}{r}-{monto_real}"
                    ws_resumen.cell(row=r, column=cfg['col_act']).number_format = '#,##0.00'
                    cambios += 1

    return cambios

from datetime import datetime

# =========================================================================
# ⚖️ LÓGICA V22: CONCILIACIÓN FINAL (DATA BS + EXCEDENTES BP/BM COL H)
# =========================================================================

def procesar_conciliacion_compleja(wb, callback_log):
    
    # --- Limpiador de números (Formato Venezuela) ---
    def limpiar_venezuela(valor):
        if not valor: return 0
        if isinstance(valor, (int, float)): return float(valor)
        
        txt = str(valor).strip().upper()
        txt = txt.replace("BS", "").replace("USD", "").replace(" ", "")
        
        es_negativo = False
        if "(" in txt and ")" in txt:
            es_negativo = True
            txt = txt.replace("(", "").replace(")", "")
        
        if "." in txt and "," in txt:
            txt = txt.replace(".", "").replace(",", ".")
        elif "," in txt: 
            txt = txt.replace(",", ".")
            
        try:
            num = float(txt)
            return -num if es_negativo else num
        except: return 0

    # 1. Cargar Hojas
    ws_apartados = obtener_hoja_flexible(wb, "APARTADOS")
    ws_data = obtener_hoja_flexible(wb, "DATA BS")
    ws_excedente = obtener_hoja_flexible(wb, "MANEJO EXCEDENTE") # Busca "MANEJO DE EXCEDENTES"

    if not (ws_apartados and ws_data and ws_excedente):
        callback_log("⚠️ Error: Faltan hojas para conciliación.")
        return 0 

    callback_log("⚖️ Conciliación V22: Data BS + Excedentes (Etiquetas BP/BM -> Col H)...")
    cambios = 0
    
    # 2. Recorremos hoja APARTADOS
    for fila in range(4, ws_apartados.max_row + 1):
        celda_banco = ws_apartados.cell(row=fila, column=2)
        celda_monto = ws_apartados.cell(row=fila, column=3)
        celda_concepto = ws_apartados.cell(row=fila, column=4)
        celda_mes = ws_apartados.cell(row=fila, column=5)
        
        concepto = str(celda_concepto.value).upper().strip() if celda_concepto.value else ""
        
        # FILTRO: Solo procesamos "SERVICIO ESPECIALIZADO"
        if "ESPECIALIZAD" in concepto:
            banco_objetivo = str(celda_banco.value).upper().strip() 
            mes_objetivo = str(celda_mes.value).upper().strip()     
            
            # -------------------------------------------------------------
            # PASO A: Sumar Egresos (DATA BS) -> Esto da los -103M
            # -------------------------------------------------------------
            suma_data_bs = 0
            for r in range(4, ws_data.max_row + 1):
                d_cuenta = str(ws_data.cell(row=r, column=13).value).upper().strip()
                d_banco = str(ws_data.cell(row=r, column=10).value).upper().strip()
                d_fecha = ws_data.cell(row=r, column=2).value
                d_monto = limpiar_venezuela(ws_data.cell(row=r, column=7).value)
                
                if d_monto != 0:
                    mes_fila = obtener_nombre_mes_es(d_fecha.month if isinstance(d_fecha, datetime) else 0)
                    if ("ESPECIALIZAD" in d_cuenta) and (banco_objetivo in d_banco or d_banco in banco_objetivo) and (mes_fila == mes_objetivo):
                        suma_data_bs += d_monto

            # -------------------------------------------------------------
            # PASO B: Sumar Ingresos (MANEJO EXCEDENTE) -> Esto da los +43M
            # -------------------------------------------------------------
            suma_excedente = 0
            
            # ¿Qué etiqueta buscamos según el banco?
            etiqueta_banco = ""
            if "PROVINCIAL" in banco_objetivo:
                etiqueta_banco = "BP"
            elif "MERCANTIL" in banco_objetivo:
                etiqueta_banco = "BM"
            
            if etiqueta_banco != "":
                # Recorremos Excedentes buscando esa etiqueta en la descripción
                max_r = max(ws_excedente.max_row, 500)
                for r in range(4, max_r + 1):
                    # 1. Descripción (Col B)
                    val_desc = ws_excedente.cell(row=r, column=2).value
                    e_desc = str(val_desc).upper().strip() if val_desc else ""
                    
                    # 2. Mes (Col C)
                    val_mes = ws_excedente.cell(row=r, column=3).value
                    e_mes = str(val_mes).upper().strip() if val_mes else ""
                    
                    # 3. FILTRO MAGISTRAL:
                    # - ¿La descripción tiene "BP" o "BM" (según corresponda)?
                    # - ¿Es el mes correcto?
                    # - ¿Es Servicio Especializado?
                    if (etiqueta_banco in e_desc) and ("ESPECIALIZAD" in e_desc) and (mes_objetivo == e_mes):
                        
                        # 4. CAPTURA: Tomamos el dinero de la COLUMNA H (8)
                        val_h = ws_excedente.cell(row=r, column=8).value
                        monto_h = limpiar_venezuela(val_h)
                        
                        if monto_h != 0:
                            suma_excedente += monto_h
                            # Log para verificar
                            # callback_log(f"   ➕ Encontrado {etiqueta_banco} en fila {r}: {monto_h:,.2f} (Col H)")

            # -------------------------------------------------------------
            # PASO C: Cálculo Final (-103 + 43 = -60)
            # -------------------------------------------------------------
            if suma_data_bs > 0 or suma_excedente > 0:
                # Egresos restan (negativo) + Ingresos suman (positivo)
                resultado_final = (suma_data_bs * -1) + suma_excedente
                
                celda_monto.value = resultado_final
                celda_monto.number_format = '#,##0.00'
                cambios += 1
                
                # Reporte en el log
                if suma_excedente > 0:
                    callback_log(f"   ✅ APARTADOS ({banco_objetivo}): {resultado_final:,.2f} (Incluye +{suma_excedente:,.2f} de Excedentes)")
                else:
                    callback_log(f"   📉 APARTADOS ({banco_objetivo}): {resultado_final:,.2f} (Solo Data BS)")

    return cambios

# ==========================================
# ORQUESTADOR PRINCIPAL
# ==========================================
def lógica_negocio(archivo_obj, callback_log, callback_progreso):
    mensajes = []
    wb = None
    try:
        # 1. CARGA INICIAL
        callback_log("🧠 Iniciando sistemas...")
        reglas_cuentas = cargar_diccionario_cuentas() 
        reglas_areas = cargar_diccionario_areas()
        tasas_historicas = cargar_tasas_historicas(callback_log)
        
        # --- AQUÍ ESTÁ EL ARREGLO DE LA API ---
        try:
            # Conexión a la nueva API (DolarApi)
            url_api = "https://ve.dolarapi.com/v1/dolares/oficial"
            callback_log(f"🌍 Consultando Dólar Oficial en: {url_api}")
            
            resp = requests.get(url_api, timeout=10)
            data = resp.json()
            
            # Usamos 'promedio' como vimos en tu captura
            precio_dolar_hoy = float(data['promedio'])
            
            callback_log(f"💰 ¡TASA OBTENIDA!: {precio_dolar_hoy}")
        except Exception as e: 
            callback_log(f"❌ Error consultando tasa: {str(e)}")
            precio_dolar_hoy = 0
        # --------------------------------------

        callback_progreso(0.1)

        # 2. ABRIR EXCEL
        callback_log("📂 Leyendo archivo Excel...")
        wb = openpyxl.load_workbook(archivo_obj)
        callback_progreso(0.3)

        # 3. ACTUALIZAR PORTADA/HISTORICO
        if precio_dolar_hoy > 0:
            ws_portada = obtener_hoja_flexible(wb, "CUENTAS POR COBRAR")
            if ws_portada:
                ws_portada["D3"] = datetime.now().strftime("%d/%m/%Y")
                ws_portada["D4"] = precio_dolar_hoy
            ws_hist = obtener_hoja_flexible(wb, "COMPORTAMIENTO TASA")
            if ws_hist:
                fila = ws_hist.max_row + 1
                ult_fecha = ws_hist.cell(row=fila-1, column=1).value
                if ult_fecha != datetime.now().strftime("%d/%m/%Y"):
                    ws_hist.cell(row=fila, column=1, value=datetime.now().strftime("%d/%m/%Y"))
                    ws_hist.cell(row=fila, column=2, value="USD")
                    ws_hist.cell(row=fila, column=3, value=precio_dolar_hoy)
                    mensajes.append("✅ Tasa Histórica Agregada")

        # 4. CLASIFICACIÓN Y CÁLCULO USD (ESTRICTO V11)
        callback_log("🚀 Clasificando y Calculando Divisas...")
        ws_data = obtener_hoja_flexible(wb, "DATA BS")
        if ws_data:
            c_clasif, c_usd = 0, 0
            for r in range(4, ws_data.max_row + 1):
                prov = str(ws_data.cell(row=r, column=12).value).upper().strip()
                if prov:
                    if not ws_data.cell(row=r, column=13).value:
                        match = next((v for k,v in reglas_cuentas if k in prov), None)
                        if not match:
                             if "COMISION" in prov: match = "GASTOS BANCARIOS"
                             elif "IVA" in prov: match = "IMPUESTOS"
                        if match: ws_data.cell(row=r, column=13).value = match; c_clasif += 1
                    if not ws_data.cell(row=r, column=15).value:
                        match = next((v for k,v in reglas_areas if k in prov), None)
                        if match: ws_data.cell(row=r, column=15).value = match
                
                val_fecha = ws_data.cell(row=r, column=2).value
                bs = limpiar_numero(ws_data.cell(row=r, column=7).value)
                
                if val_fecha and bs != 0 and not ws_data.cell(row=r, column=8).value:
                    f_key = formatear_fecha_para_api(val_fecha)
                    tasa = buscar_tasa_inteligente(f_key, tasas_historicas)
                    if tasa and tasa > 0:
                        ws_data.cell(row=r, column=8).value = bs / tasa
                        ws_data.cell(row=r, column=8).number_format = '#,##0.00'
                        c_usd += 1

            mensajes.append(f"✅ {c_clasif} Filas Clasificadas")
            mensajes.append(f"✅ {c_usd} Conversiones a USD")
        
        callback_progreso(0.6)

        # 5. CONCILIACIÓN
        cambios_conc = procesar_conciliacion_compleja(wb, callback_log)
        if cambios_conc > 0: mensajes.append("✅ Conciliación Exitosa")

        callback_progreso(0.8)
        
        # 6. RESUMEN SEMANAL (CORREGIDO)
        cambios_sem = procesar_resumen_semanal(wb, callback_log)
        if cambios_sem > 0: mensajes.append(f"✅ Resumen Semanal Actualizado ({cambios_sem} celdas)")
        else: mensajes.append("ℹ️ Resumen Semanal: Sin cambios nuevos")
        
        callback_progreso(0.9)

        # 7. GUARDAR
        callback_log("💾 Preparando descarga....")
        output = io.BytesIO()
        wb.save(output)
        wb.close()
        output.seek(0)
        callback_progreso(1.0)
        
        return output, "\n".join(mensajes)

    except PermissionError:
        return False, "⚠️ CIERRA EL EXCEL. Está abierto y bloqueado."
    except Exception as e:
        return False, f"❌ Error técnico: {str(e)}"

# ==========================================
# 🎮 INTERFAZ WEB (STREAMLIT)
# ==========================================
if __name__ == "__main__":
    st.set_page_config(page_title="Platco Bot Financiero", page_icon="💰")

    # Cargar Logo si existe
    if os.path.exists("logo.png"):
        st.image("logo.png", width=200)

    st.title("🤖 Sistema de Gestión Financiera AI")
    st.markdown("---")

    # Botón para subir archivo
    uploaded_file = st.file_uploader("📂 Sube tu archivo Excel Financiero", type=["xlsx"])

    if uploaded_file is not None:
        st.success(f"Archivo cargado: {uploaded_file.name}")
        
        if st.button("🚀 EJECUTAR AUTOMATIZACIÓN", type="primary"):
            
            # Área de logs visual
            log_container = st.empty()
            logs_historial = []

            def web_log(msg):
                logs_historial.append(f"> {msg}")
                # Muestra los últimos 10 mensajes
                log_container.text("\n".join(logs_historial[-10:]))

            bar = st.progress(0)

            try:
                web_log("--- INICIANDO PROCESO ---")
                
                # LLAMAMOS A TU LÓGICA (MODIFICADA EN EL PASO 2)
                # Nota: Pasamos 'uploaded_file' directo, no una ruta.
                archivo_resultado, texto_resultado = lógica_negocio(uploaded_file, web_log, bar.progress)
                
                web_log("--- FINALIZADO ---")
                st.success("¡Proceso Terminado!")

                # MOSTRAR LOGS COMPLETOS
                with st.expander("Ver Reporte Detallado"):
                    st.text(texto_resultado)

                # BOTÓN DE DESCARGA
                now = datetime.now().strftime("%Y%m%d_%H%M")
                st.download_button(
                    label="📥 DESCARGAR ARCHIVO PROCESADO",
                    data=archivo_resultado,
                    file_name=f"Finanzas_Procesado_{now}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Error Crítico: {str(e)}")

    

