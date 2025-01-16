import streamlit as st
import pandas as pd
import io

# Funzione per trasporre taglie da un range di colonne
def trasponi_taglie(file, colonna_inizio, colonna_fine):
    # Leggi il file Excel
    df = pd.read_excel(file, engine="openpyxl")
    
    # Converti i nomi delle colonne in indici
    colonne = df.columns
    col_inizio_idx = colonne.get_loc(colonna_inizio)
    col_fine_idx = colonne.get_loc(colonna_fine) + 1  # Include la colonna finale
    
    # Isola le colonne del range specificato
    taglie_selezionate = colonne[col_inizio_idx:col_fine_idx]
    
    # Crea una lista per il dataframe trasposto
    righe = []
    for _, row in df.iterrows():
        for taglia in taglie_selezionate:
            if not pd.isna(row[taglia]):  # Salta celle vuote
                righe.append({
                    "Index": row["Index"],
                    "Taglia": taglia,
                    "Quantità": row[taglia]
                })
    
    # Crea un nuovo dataframe
    df_trasposto = pd.DataFrame(righe)
    return df_trasposto

# Interfaccia Streamlit
st.title("Trasposizione Taglie in Verticale")
st.write("Carica il tuo file Excel e specifica il range di colonne in cui si trovano le taglie.")

# Caricamento del file Excel
file = st.file_uploader("Carica il file Excel", type=["xlsx"])

# Input per specificare il range di colonne
colonna_inizio = st.text_input("Colonna di inizio (es. C)")
colonna_fine = st.text_input("Colonna di fine (es. Y)")

if file and colonna_inizio and colonna_fine and st.button("Trasponi"):
    try:
        # Trasforma il file e crea il nuovo dataframe
        nuovo_df = trasponi_taglie(file, colonna_inizio, colonna_fine)
        
        # Salva in un file Excel temporaneo
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            nuovo_df.to_excel(writer, index=False, sheet_name="Taglie Trasposte")
        output.seek(0)
        
        # Link per scaricare il file
        st.success("File trasformato con successo!")
        st.download_button(
            label="Scarica il file Excel trasformato",
            data=output,
            file_name="taglie_trasposte.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Errore durante la trasformazione: {e}")
