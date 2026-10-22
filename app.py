"""Aplicación local para proyección del cargo Compensación PE de EFE Valparaíso."""
from __future__ import annotations

import streamlit as st
from cen_core import (parse_balance, parse_invoice, parse_price_quillota,
                      parse_hourly_csv, csv_bytes)

st.set_page_config(page_title="Compensación PE | EFE Valparaíso", page_icon="⚡", layout="wide")
st.markdown("""<style>
.block-container{max-width:1220px;padding-top:1.7rem}
h1{letter-spacing:-.045em}.stMetric{border:1px solid #dfe6ef;border-radius:12px;padding:16px}
</style>""", unsafe_allow_html=True)
st.title("⚡ Compensación por Precio Estabilizado")
st.caption("EFE Valparaíso · Enel Generación · Cálculo desde archivos oficiales del CEN · Versión 1.0")

with st.sidebar:
    st.header("1 · Archivos de entrada")
    balance_uploads = st.file_uploader("Balances comerciales CEN (2609P, etc.)", type=["xlsx", "xlsm"], accept_multiple_files=True,
                                       help="Archivo 'Balance_2609P.xlsm' u otros meses; se procesa la hoja Cuadro de Energía y Balance Valorizado.")
    invoice_uploads = st.file_uploader("Proformas / facturas históricas de EFE", type=["xlsx", "xlsm"], accept_multiple_files=True,
                                       help="Archivo EFE_VALPO_SEP2026.xlsx, con hoja Cargos_CEN.")
    price_upload = st.file_uploader("Precios estabilizados (opcional)", type=["xlsx", "xlsm"],
                                    help="Solo para referencia documental de Quillota 220 kV. No se utiliza como cargo PE.")
    st.divider()
    st.header("2 · Identificación EFE")
    clave_text = st.text_area("Claves comerciales (una por línea)", "MERVAL_ELSOL\nMERVAL_LIMACHE", height=92,
                              help="Claves de retiro encontradas en el balance del CEN, no la barra física genérica Quillota 220.")
    keys = [k.strip() for k in clave_text.replace(",", "\n").splitlines() if k.strip()]
    st.caption("Archivos procesados localmente; la app no transmite los datos al CEN ni a otros sistemas.")

@st.cache_data(show_spinner=False)
def calc_balance(data, name, keys_tuple):
    return parse_balance(data, name, list(keys_tuple))

@st.cache_data(show_spinner=False)
def calc_invoice(data, name):
    return parse_invoice(data, name)

balances = []
errors = []
for item in balance_uploads:
    try:
        balances.append(calc_balance(item.getvalue(), item.name, tuple(keys)))
    except Exception as e:
        errors.append(f"{item.name}: {e}")
balances.sort(key=lambda v: v.label)

invoices = []
for item in invoice_uploads:
    try:
        invoices.append(calc_invoice(item.getvalue(), item.name))
    except Exception as e:
        errors.append(f"{item.name}: {e}")
invoices.sort(key=lambda v: v["mes_balance"])

for error in errors:
    st.error(f"No se pudo procesar: {error}")

clp = lambda n: "$" + f"{n:,.0f}".replace(",", ".")
numformat = lambda n, dec=2: f"{n:,.{dec}f}".replace(",", "_").replace(".", ",").replace("_", ".")

if not balances and not invoices:
    st.info("Carga el Balance_2609P.xlsm en la barra lateral. También puedes añadir EFE_VALPO_SEP2026.xlsx para comparar contra el cobro anterior. Si aún no existe el balance, puedes proyectar con una tasa ingresada manualmente.")

main, hist, hour, info = st.tabs(["📊 Calcular proyección", "📈 Histórico", "🕒 Cálculo horario", "📖 Metodología y trazabilidad"])

with main:
    st.subheader("Proyección de cargo mensual")
    chosen = None
    if balances:
        opts = [f"{b.label} · {b.tarifa_clp_mwh:,.2f} $/MWh" for b in balances]
        pick = st.selectbox("Seleccionar balance CEN", range(len(balances)), format_func=lambda i:opts[i], index=len(balances)-1)
        chosen = balances[pick]
    candidates = ["Balance CEN · promedio mensual calculado", "Ingresar tarifa manual"] if chosen else ["Ingresar tarifa manual"]
    if invoices:
        candidates.append("Última tasa efectiva facturada a EFE")
    source = st.radio("Origen de tarifa unitaria ($/MWh)", candidates, horizontal=False)
    if source.startswith("Balance CEN") and chosen:
        rate = chosen.tarifa_clp_mwh
    elif source == "Última tasa efectiva facturada a EFE":
        rate = invoices[-1]["tarifa_clp_kwh"] * 1000.
    else:
        rate = st.number_input("Tarifa estimada ($/MWh)", min_value=-100000.0, max_value=100000.0,
                               value=round((invoices[-1]["tarifa_clp_kwh"]*1000. if invoices else 0.),3), step=100.,format="%.3f",
                               help="Si aún no existe balance, ingresar un valor de referencia histórico.")
    st.markdown("**Consumo sobre el cual aplicar la tarifa**")
    manual_default = float(chosen.consumo_efe_kwh) if chosen else (float(invoices[-1]["consumo_kwh"]) if invoices else 0.)
    energy_type = st.radio("Fuente de consumo", ["Consumo detectado en balance", "Consumo proyectado / manual"] if chosen else ["Consumo proyectado / manual"], horizontal=True)
    if energy_type == "Consumo detectado en balance" and chosen:
        total_kwh = chosen.consumo_efe_kwh
    else:
        total_kwh = st.number_input("Retiro total EFE (kWh)", min_value=0.0, value=round(manual_default,2), step=1000.,format="%.2f")
    with st.expander("Ajustes contractuales y sensibilidades", expanded=False):
        transfer = st.number_input("Porcentaje de traslado contractual (%)", 0., 200., 100., 1.,
                                   help="100% por defecto. Solo modificar si el contrato exige un porcentaje diferente.")
        adjust = st.number_input("Ajuste adicional manual por reliquidaciones ($)", value=0.0, step=10000.,format="%.0f",
                                 help="Se aplica después de multiplicar tarifa por energía, si corresponde al cargo PE.")
        sensitivity = st.slider("Escenarios bajo/alto: variación sobre tarifa", 0, 80, 20, 5, format="%d%%")
    base = total_kwh/1000 * rate
    total = base * transfer/100 + adjust
    variation_a = total_kwh/1000 * rate * (1-sensitivity/100)*transfer/100 + adjust
    variation_b = total_kwh/1000 * rate * (1+sensitivity/100)*transfer/100 + adjust
    low, high = sorted([variation_a, variation_b])
    low_rate, high_rate = sorted([rate*(1-sensitivity/100), rate*(1+sensitivity/100)])
    a,b,c = st.columns(3)
    a.metric("Cargo proyectado (sin IVA)", clp(total))
    b.metric("Tarifa aplicada ($/MWh)", numformat(rate,2))
    c.metric("Retiro EFE (MWh)", numformat(total_kwh/1000,3))
    st.caption("Para $/kWh divide el valor mostrado de $/MWh entre 1.000.")
    st.write("**Escenarios de presupuesto (sin IVA)**")
    scenario = [
        {"Escenario":"Bajo", "Tarifa ($/MWh)":round(low_rate,3), "Monto ($)":round(low)},
        {"Escenario":"Medio", "Tarifa ($/MWh)":round(rate,3), "Monto ($)":round(total)},
        {"Escenario":"Alto", "Tarifa ($/MWh)":round(high_rate,3), "Monto ($)":round(high)},
    ]
    st.dataframe(scenario, hide_index=True, use_container_width=True)
    if chosen:
        st.markdown("**Origen del resultado del balance seleccionado**")
        c1,c2 = st.columns(2)
        c1.metric("Compensación agregada SEN ($)",clp(chosen.compensacion_sen_clp))
        c2.metric("Retiros netos SEN (MWh)",numformat(chosen.retiro_sen_mwh,3))
        if chosen.detalle:
            st.dataframe(chosen.detalle,hide_index=True,use_container_width=True)
        for note in chosen.avisos:
            st.warning(note)
    elif source == "Ingresar tarifa manual":
        st.warning("Resultado indicativo con datos manuales; no es un valor oficial calculado por el CEN.")
    st.download_button("⬇ Descargar proyección en CSV", data=csv_bytes([{
        "mes_balance": chosen.label if chosen else "Estimación manual",
        "fuente_tarifa":source, "retiro_efe_kwh":total_kwh,"tarifa_clp_mwh":rate,
        "factor_traslado_pct":transfer,"ajuste_clp":adjust,
        "cargo_estimado_clp":round(total),"escenario_bajo_clp":round(low),"escenario_alto_clp":round(high),
        "calidad":"Estimación no oficial por tasa media; requiere contraste con asignación por retiro",
    }]),file_name="proyeccion_compensacion_pe_efe.csv",mime="text/csv")

with hist:
    st.subheader("Historial y contraste con cobros")
    if invoices:
        table = [{"Mes de balance":i["mes_balance"],"Archivo":i["archivo"],
                  "Retiro EFE (kWh)":round(i["consumo_kwh"],2),
                  "Cargo Enel ($)":round(i["compensacion_clp"]),
                  "Tasa efectiva EFE ($/kWh)":round(i["tarifa_clp_kwh"],6)} for i in invoices]
        st.dataframe(table,hide_index=True,use_container_width=True)
        if len(invoices)>1:
            st.line_chart({i["mes_balance"]:i["tarifa_clp_kwh"] for i in invoices}, y_label="$/kWh")
        st.download_button("Descargar histórico de facturas CSV",csv_bytes(table),"historico_facturas_efe.csv","text/csv")
    else:
        st.info("Carga proformas históricas EFE con hoja Cargos_CEN para calcular la tasa efectiva y los pagos de cada mes.")
    if balances:
        table = [{"Mes CEN":b.label,"Compensación agregada ($)":round(b.compensacion_sen_clp),
                  "Retiros SEN (MWh)":round(b.retiro_sen_mwh,3),
                  "Tarifa media SEN ($/MWh)":round(b.tarifa_clp_mwh,3),
                  "Retiro EFE detectado (kWh)":round(b.consumo_efe_kwh,2),
                  "Cargo indicativo ($)":round(b.proyeccion_clp)} for b in balances]
        st.dataframe(table,hide_index=True,use_container_width=True)
        st.download_button("Descargar balances calculados CSV",csv_bytes(table),"historico_balance_cen.csv","text/csv")
    if invoices and balances:
        compared = []
        actuals = {i["mes_balance"]:i for i in invoices}
        for b in balances:
            if b.label in actuals:
                i = actuals[b.label]
                pct = (b.proyeccion_clp / i["compensacion_clp"] - 1) * 100 if i["compensacion_clp"] else None
                compared.append({"Mes de balance":b.label,"Cargo Enel ($)":round(i["compensacion_clp"]),
                    "Proyección ($)":round(b.proyeccion_clp), "Diferencia ($)":round(b.proyeccion_clp-i["compensacion_clp"]),
                    "Error relativo (%)":round(pct,2) if pct is not None else None})
        if compared:
            st.write("**Conciliación del modelo contra la factura real (meses coincidentes)**")
            st.dataframe(compared,hide_index=True,use_container_width=True)
            st.caption("Comparación informativa: la tasa media SEN no garantiza coincidencia con la asignación individual de EFE.")

with hour:
    st.subheader("Cálculo horario (mayor precisión si dispones de datos completos)")
    st.write("Si tienes un archivo con compensaciones del SEN, retiros totales y retiros EFE para cada hora, puedes calcular la suma horaria en vez de aplicar un promedio mensual.")
    st.download_button("⬇ Descargar plantilla horaria CSV",data=csv_bytes([{
        "fecha":"2026-09-01","hora":"00:00","compensacion_sen_clp":"", "retiro_sen_mwh":"", "retiro_efe_mwh":""
    }]),file_name="plantilla_compensacion_horaria.csv",mime="text/csv")
    hourly_file = st.file_uploader("Subir archivo horario CSV",type=["csv"])
    if hourly_file:
        try:
            x = parse_hourly_csv(hourly_file.getvalue())
            a,b,c=st.columns(3)
            a.metric("Cargo horario calculado ($)",clp(x["clp"]))
            b.metric("Horas procesadas",f"{x['horas']}")
            c.metric("Retiro horario EFE total (MWh)",numformat(x["mwh"],2))
            if x["omitidas"]:st.warning(f"Se omitieron {x['omitidas']} horas con denominador no positivo.")
            st.dataframe(x["detalle"],hide_index=True,use_container_width=True)
            st.download_button("Descargar resultado por hora CSV",csv_bytes(x["detalle"]),"detalle_compensacion_horaria.csv","text/csv")
            st.warning("Requiere confirmar que las compensaciones horarias de entrada son asignables íntegramente según la metodología CEN; el cálculo es indicativo hasta contrastarlo con la asignación oficial por retiro.")
        except Exception as exc:
            st.error(f"Error en CSV horario: {exc}")

with info:
    st.subheader("Metodología y límites del programa")
    st.latex(r"T_{PE,mes}=\frac{|C_{SEN,mes}|}{R_{SEN,mes}}\quad [\mathrm{CLP/MWh}]")
    st.latex(r"Cargo_{EFE,proy}=T_{PE,mes}\times E_{EFE,mes}\quad [\mathrm{CLP}]")
    st.markdown("""
**Balance CEN:** del libro `Balance_2609P.xlsm` toma la fila `Total general` de la hoja `Cuadro de Energía`, columna **M**, cuyo título es `Compensación precio nudo/estabilizado`. Toma la magnitud del valor contable (en el archivo de referencia es negativo). Suma los retiros físicos netos de tipos **L, L_D y R** de `Balance Valorizado`, en kWh (negativos según signo contable), convirtiéndolos a MWh. La tasa es un **indicador agregado mensual**, no un cargo regulatorio único por retiro.

**Consumo EFE:** selecciona en `Balance Valorizado` las claves `MERVAL_ELSOL` y `MERVAL_LIMACHE` (editables). Para la factura anterior, toma el retiro y cargo de la hoja `Cargos_CEN` (columnas G y N). La barra **Quillota 220 kV** no reemplaza esas claves: es necesario revisar la correspondencia comercial antes de contabilizar.

**Precisión:** el programa reproduce una **proyección presupuestaria**. El monto facturado puede diferir porque la asignación CEN es por retiro y puede considerar perfiles horarios y reliquidaciones. No se ha demostrado que todo el agregado de la columna M corresponda exclusivamente a la componente PMGD que factura Enel; podría incluir precio nudo. La tasa agregada no debe presentarse como valor oficial individual de EFE.

**Fecha:** `2609P` significa septiembre de 2026 y balance preliminar; el desfase del contrato descrito por Enel llevaría este mes a la factura de octubre de 2026. Las versiones definitivas pueden modificar los resultados.

**Seguridad:** los archivos se procesan dentro del equipo donde ejecutas esta aplicación. No se ejecutan macros de los `.xlsm`; se leen valores guardados (cacheados) por Excel. No hay API ni sincronización automática con PLABACOM.
""")
    if price_upload:
        try:
            prices = parse_price_quillota(price_upload.getvalue())
            if prices:
                st.markdown("**Referencia documental Quillota 220 kV**")
                st.dataframe(prices,hide_index=True,use_container_width=True)
                st.info("Estos precios estabilizados de generación NO son compensaciones unitarias del SEN; no se suman ni se usan directamente para estimar la factura.")
            else:
                st.warning("No se encontraron los bloques de Quillota 220 kV en el archivo de precios cargado.")
        except Exception as exc:
            st.error(f"No pude leer el archivo de precios: {exc}")
