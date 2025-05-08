import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
import time

st.set_page_config(page_title="Consulta de Direcciones DIAN", layout="centered")

st.title("📌 Herramienta para Consultar Dirección, Ciudad y Departamento por NIT")
st.markdown("""
Bienvenida a esta herramienta diseñada para facilitar tu trabajo.  
Solo necesitas subir un archivo de Excel que contenga una columna llamada **`documento`** con los NIT que deseas consultar.

Al final, podrás descargar un archivo con:
- Dirección
- Departamento
- Ciudad
- Código DANE del departamento
- Código DANE del municipio

🕒 La consulta toma algunos segundos por cada NIT, así que ten paciencia.  
❗ Si alguna celda estaba vacía, el resultado será **"Sin información"**.
""")

@st.cache_data
def load_dane_codes():
    df_dane = pd.read_excel("Códigos DANE.xlsx", dtype=str)
    for col in df_dane.columns:
        df_dane[col] = df_dane[col].str.strip().str.upper()
    return df_dane

# Cargar códigos DANE
df_dane = load_dane_codes()

# Subir archivo de NITs
file = st.file_uploader("Sube tu archivo Excel con la columna 'documento'", type=["xlsx"])

if file is not None:
    df = pd.read_excel(file, dtype=str)
    
    if 'documento' not in df.columns:
        st.error("⚠️ El archivo debe tener una columna llamada 'documento'")
    else:
        documentos = df['documento'].fillna("").astype(str).apply(lambda x: x.strip().split('.')[0])
        resultados = []

        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, doc in enumerate(documentos):
            if doc == "":
                resultados.append({
                    'documento': "",
                    'dirección': "Sin información",
                    'departamento': "Sin información",
                    'ciudad': "Sin información",
                    'código departamento': "Sin información",
                    'código municipio': "Sin información"
                })
                continue

            url = f"https://www.rues.org.co/RM/Consultas/DetallePersona?documento={doc}"
            try:
                response = requests.get(url, timeout=10)
                soup = BeautifulSoup(response.text, 'html.parser')

                def extraer_valor(pregunta):
                    etiqueta = soup.find('h2', string=lambda t: t and pregunta in t)
                    if etiqueta:
                        valor = etiqueta.find_next_sibling('div', class_='value')
                        return valor.text.strip() if valor else "Sin información"
                    return "Sin información"

                direccion = extraer_valor("¿Cuál es su dirección?")
                departamento = extraer_valor("¿Cuál es el departamento en el que se registra?").upper().strip()
                ciudad = extraer_valor("¿Cuál es la ciudad en el que se registra?").upper().strip()

                cod_dep = df_dane[df_dane['Nombre departamento'] == departamento]['Código departamento'].values
                cod_mun = df_dane[(df_dane['Nombre departamento'] == departamento) & (df_dane['Nombre municipio'] == ciudad)]['Código municipio'].values

                resultados.append({
                    'documento': doc,
                    'dirección': direccion,
                    'departamento': departamento if departamento else "Sin información",
                    'ciudad': ciudad if ciudad else "Sin información",
                    'código departamento': cod_dep[0] if len(cod_dep) > 0 else "No encontrado",
                    'código municipio': cod_mun[0] if len(cod_mun) > 0 else "No encontrado"
                })

            except Exception as e:
                resultados.append({
                    'documento': doc,
                    'dirección': "Error",
                    'departamento': "Error",
                    'ciudad': "Error",
                    'código departamento': "Error",
                    'código municipio': "Error"
                })

            progress_bar.progress((i + 1) / len(documentos))
            status_text.text(f"Consultando NIT {i + 1} de {len(documentos)}")
            time.sleep(4)

        st.success("✅ Consulta finalizada. Puedes descargar los resultados abajo.")

        df_resultado = pd.DataFrame(resultados)
        st.dataframe(df_resultado)

        output_file = "resultado_consultas.xlsx"
        df_resultado.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button("📥 Descargar archivo con resultados", f, file_name="resultado_consultas.xlsx")