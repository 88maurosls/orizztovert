import streamlit as st
import pandas as pd
import io

# Funzione per trasporre le taglie selezionate
def trasponi_taglie(file, range_taglie):
    # Legge il file Excel caricato
    df = pd.read_excel(file, engine="openpyxl")
    
    # Trova le colonne corrispondenti al range di taglie
    taglie_selezionate = [col for col in df.columns if col in range_taglie]
    
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
    
    # Converte in dataframe
    df_trasposto = pd.DataFrame(righe)
    return df_trasposto

# Interfaccia Streamlit
st.title("Trasposizione Taglie in Verticale")
st.write("Carica il tuo file Excel e specifica un range di taglie.")

# Caricamento del file Excel
file = st.file_uploader("Carica il file Excel", type=["xlsx"])

# Range di taglie da selezionare
range_taglie = st.text_input("Inserisci le taglie (separate da virgola)", value="5, 6, 7, 8, 9")
range_taglie = [taglia.strip() for taglia in range_taglie.split(",")]

if file and st.button("Trasponi"):
    try:
        # Trasforma il file e crea il nuovo dataframe
        nuovo_df = trasponi_taglie(file, range_taglie)
        
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
