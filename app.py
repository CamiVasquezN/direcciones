import streamlit as st
import pandas as pd
import time
import requests
from bs4 import BeautifulSoup
from io import BytesIO

# Introducción en la app
st.set_page_config(page_title="Consulta de Direcciones por NIT")
st.title("📍 Consulta de Dirección, Ciudad y Departamento por NIT")

# Introducción bonita
st.markdown("""
Bienvenida a la herramienta de **consulta de direcciones**, creada con mucho cariño para ti. 

Esta aplicación facilita la consulta de direcciones, ciudades y departamentos de empresas a partir de su NIT. Solo necesitas:

1. Subir el archivo Excel que contiene los **NITs** de las empresas.
2. La herramienta hará la consulta por ti.
3. Después, podrás descargar un archivo Excel con la información de **dirección**, **ciudad** y **departamento** de cada NIT.

Es muy fácil de usar y no requiere conocimientos técnicos. Si tienes alguna duda, solo pregúntame.

¡Espero que te sea de mucha ayuda! 😊

### Instrucciones:
- **Sube tu archivo Excel** con una columna llamada **`documento`**, que contenga los NITs.
- La aplicación comenzará a consultar los datos y podrás verlos en la pantalla.
- Cuando termine, podrás **descargar** los resultados.

¡Gracias por usar esta herramienta!
""")

# Función para obtener la dirección, ciudad y departamento del NIT
def obtener_info_direccion(nit):
    url = f"https://empresas.larepublica.co/colombia/antioquia/medellin/{nit}"
    response = requests.get(url)
    
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")
        
        # Extraemos los datos
        direccion = soup.find("div", {"class": "company-attribute"}).find_next("div", {"class": "value"}).text.strip()
        departamento = soup.find("h2", string="¿Cuál es el departamento en el que se registra?").find_next("div", {"class": "value"}).text.strip()
        ciudad = soup.find("h2", string="¿Cuál es la ciudad en el que se registra?").find_next("div", {"class": "value"}).text.strip()
        
        return direccion, ciudad, departamento
    else:
        return "No encontrado", "No encontrado", "No encontrado"

# Subir archivo Excel
archivo = st.file_uploader("Sube tu archivo Excel con una columna llamada 'documento'", type=["xlsx"])

if archivo:
    df = pd.read_excel(archivo)

    if "documento" not in df.columns:
        st.error("El archivo no contiene la columna 'documento'.")
    else:
        # Limpiar los NITs para asegurarse de que no tengan decimales, puntos o comas
        df["documento"] = df["documento"].apply(lambda x: str(x).replace(".", "").replace(",", "").split('.')[0].strip())

        # Verificar que todos los NITs sean números enteros válidos (sin puntos ni comas)
        invalid_nits = df[~df["documento"].str.isdigit()]
        if not invalid_nits.empty:
            st.error(f"Hay NITs inválidos que no son números enteros: {invalid_nits['documento'].values}")
        else:
            resultados = []
            with st.spinner("Consultando NITs... por favor espera."):
                for i, row in df.iterrows():
                    nit = row["documento"]
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