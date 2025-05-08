import streamlit as st
import pandas as pd
import time
import requests
from bs4 import BeautifulSoup
from io import BytesIO

def obtener_info_direccion(nit):
    url = f"https://empresas.larepublica.co/colombia/antioquia/medellin/t-u-importaciones-sas-{nit}"
    headers = {"User-Agent": "Mozilla/5.0"}
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code != 200:
            return "No encontrado", "No encontrado", "No encontrado"
        
        soup = BeautifulSoup(response.text, "html.parser")
        direccion = ciudad = departamento = "No encontrado"

        for h2 in soup.find_all('h2', class_='key'):
            if "dirección" in h2.text:
                direccion = h2.find_next('div', class_='value').text.strip()
            elif "departamento" in h2.text:
                departamento = h2.find_next('div', class_='value').text.strip()
            elif "ciudad" in h2.text:
                ciudad = h2.find_next('div', class_='value').text.strip()

        return direccion, ciudad, departamento

    except Exception as e:
        return "Error", "Error", "Error"

st.set_page_config(page_title="Consulta de Direcciones por NIT")
st.title("📍 Consulta de Dirección, Ciudad y Departamento por NIT")

archivo = st.file_uploader("Sube tu archivo Excel con una columna llamada 'documento'", type=["xlsx"])

if archivo:
    df = pd.read_excel(archivo)

    if "documento" not in df.columns:
        st.error("El archivo no contiene la columna 'documento'.")
    else:
        resultados = []
        with st.spinner("Consultando NITs... por favor espera."):
            for i, row in df.iterrows():
                nit = str(row["documento"]).strip()
                direccion, ciudad, departamento = obtener_info_direccion(nit)
                resultados.append([nit, direccion, ciudad, departamento])
                time.sleep(4)

        resultados_df = pd.DataFrame(resultados, columns=["NIT", "Dirección", "Ciudad", "Departamento"])

        st.success("✅ Consultas completadas.")
        st.dataframe(resultados_df)

        # Botón para descargar
        output = BytesIO()
        resultados_df.to_excel(output, index=False)
        st.download_button("📥 Descargar resultados en Excel", data=output.getvalue(),
                           file_name="resultado_direcciones.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")