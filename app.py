import streamlit as st
import pandas as pd
import io
from openpyxl.utils import column_index_from_string

# Funzione per trasporre taglie da un range di colonne
def trasponi_taglie(file, colonna_inizio, colonna_fine, colonna_identificativa):
    # Leggi il file Excel
    df = pd.read_excel(file, engine="openpyxl")
    
    # Converti i riferimenti delle colonne (lettere) in indici numerici
    col_inizio_idx = column_index_from_string(colonna_inizio) - 1  # Indici 0-based
    col_fine_idx = column_index_from_string(colonna_fine)          # Indici 0-based + 1 per includere la colonna fine
    
    # Seleziona le colonne del range
    colonne_range = df.iloc[:, col_inizio_idx:col_fine_idx]
    colonne_range.columns = df.columns[col_inizio_idx:col_fine_idx]
    
    # Ottieni la colonna identificativa
    colonna_id = df[colonna_identificativa]
    
    # Crea una lista per il dataframe trasposto
    righe = []
    for i, row in df.iterrows():
        for colonna in colonne_range.columns:
            if not pd.isna(row[colonna]):  # Salta celle vuote
                righe.append({
                    colonna_identificativa: row[colonna_identificativa],  # Valore effettivo della colonna identificativa
                    "Colonna di Riferimento": colonna,
                    "Valore": row[colonna]
                })
    
    # Crea un nuovo dataframe
    df_trasposto = pd.DataFrame(righe)
    return df_trasposto

# Interfaccia Streamlit
st.title("Trasposizione Generica di Colonne in Verticale")
st.write("Carica il tuo file Excel, specifica il range di colonne da trasporre e scegli una colonna identificativa.")

# Caricamento del file Excel
file = st.file_uploader("Carica il file Excel", type=["xlsx"])

# Input per specificare il range di colonne
colonna_inizio = st.text_input("Colonna di inizio (es. C)")
colonna_fine = st.text_input("Colonna di fine (es. Y)")

# Input per specificare la colonna identificativa
colonna_identificativa = st.text_input("Nome della colonna identificativa (es. Index)")

if file and colonna_inizio and colonna_fine and colonna_identificativa and st.button("Trasponi"):
    try:
        # Trasforma il file e crea il nuovo dataframe
        nuovo_df = trasponi_taglie(file, colonna_inizio, colonna_fine, colonna_identificativa)
        
        # Salva in un file Excel temporaneo
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            nuovo_df.to_excel(writer, index=False, sheet_name="Trasposizione")
        output.seek(0)
        
        # Link per scaricare il file
        st.success("File trasformato con successo!")
        st.download_button(
            label="Scarica il file Excel trasformato",
            data=output,
            file_name="trasposizione.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Errore durante la trasformazione: {e}")
