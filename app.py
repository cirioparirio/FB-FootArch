import streamlit as st
import pandas as pd
import io
import time
import numpy as np

# Configurazione della pagina con tema scuro
st.set_page_config(
    page_title="Visualizzatore Excel",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Applica tema scuro personalizzato
st.markdown("""
<style>
    .stApp {
        background-color: #121212;
        color: #FFFFFF;
    }
    .stSidebar {
        background-color: #1E1E1E;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
    }
    .stDownloadButton>button {
        background-color: #3498DB;
        color: white;
    }
    .stDataFrame {
        background-color: #2D2D2D;
    }
    .stMarkdown {
        color: #FFFFFF;
    }
    header {
        background-color: #1E1E1E !important;
    }
    .stTabs [data-baseweb="tab-list"] {
        background-color: #1E1E1E;
    }
    .stTabs [data-baseweb="tab"] {
        color: white;
    }
    .stTabs [aria-selected="true"] {
        background-color: #4CAF50;
    }
    /* Stile per i filtri */
    .filter-container {
        background-color: #2D2D2D;
        padding: 10px;
        border-radius: 5px;
        margin-bottom: 10px;
    }
    .filter-title {
        font-weight: bold;
        margin-bottom: 5px;
    }
</style>
""", unsafe_allow_html=True)

# Funzione per caricare il file Excel con gestione ottimizzata della memoria
@st.cache_data(show_spinner=True)
def load_excel_sheets(uploaded_file):
    """Carica l'elenco dei fogli nel file Excel."""
    try:
        # Ottieni l'elenco dei fogli
        xls = pd.ExcelFile(uploaded_file)
        return xls.sheet_names
    except Exception as e:
        st.error(f"Errore durante il caricamento del file: {e}")
        return None

@st.cache_data(show_spinner=True)
def load_excel_sheet(uploaded_file, sheet_name, header_row=0):
    """Carica un foglio specifico del file Excel."""
    try:
        # Determina la riga di intestazione in base al nome del foglio
        if sheet_name == 'Archivio':
            # Per il foglio Archivio, usa la riga 7 come intestazione
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=6)
        else:
            # Per altri fogli, usa la prima riga come intestazione
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        
        # Converti le colonne con tipi misti in stringhe per evitare errori di confronto
        for col in df.columns:
            if df[col].dtype == 'object':
                # Converti i valori NaN in stringhe vuote per evitare errori
                df[col] = df[col].fillna('').astype(str)
        
        return df
    except Exception as e:
        st.error(f"Errore durante il caricamento del foglio {sheet_name}: {e}")
        return None

# Funzione per creare filtri avanzati per ogni colonna
def create_advanced_filters(df):
    """Crea filtri avanzati per ogni colonna del DataFrame."""
    filters = {}
    
    # Crea un espander per i filtri
    with st.expander("Filtri avanzati (stile Excel)", expanded=False):
        # Organizza i filtri in colonne
        num_cols = 3  # Numero di colonne per i filtri
        col_filters = st.columns(num_cols)
        
        # Distribuisci i filtri tra le colonne
        for i, col_name in enumerate(df.columns):
            with col_filters[i % num_cols]:
                st.markdown(f"**Filtro: {col_name}**")
                
                # Determina il tipo di dati della colonna
                col_type = df[col_name].dtype
                
                # Gestisci diversi tipi di filtri in base al tipo di dati
                if pd.api.types.is_numeric_dtype(col_type):
                    # Filtro numerico
                    min_val = float(df[col_name].min()) if not pd.isna(df[col_name].min()) else 0
                    max_val = float(df[col_name].max()) if not pd.isna(df[col_name].max()) else 100
                    
                    # Evita valori min e max identici
                    if min_val == max_val:
                        max_val = min_val + 1
                    
                    # Slider per intervallo numerico
                    filters[col_name] = st.slider(
                        f"Intervallo per {col_name}",
                        min_value=min_val,
                        max_value=max_val,
                        value=(min_val, max_val),
                        key=f"slider_{col_name}"
                    )
                elif pd.api.types.is_datetime64_dtype(col_type):
                    # Filtro data
                    min_date = df[col_name].min().date()
                    max_date = df[col_name].max().date()
                    
                    # Evita date min e max identiche
                    if min_date == max_date:
                        max_date = min_date + pd.Timedelta(days=1)
                    
                    # Date input per intervallo di date
                    start_date = st.date_input(
                        f"Data inizio per {col_name}",
                        value=min_date,
                        key=f"start_date_{col_name}"
                    )
                    end_date = st.date_input(
                        f"Data fine per {col_name}",
                        value=max_date,
                        key=f"end_date_{col_name}"
                    )
                    filters[col_name] = (start_date, end_date)
                else:
                    # Filtro categorico (multiselect)
                    unique_values = df[col_name].dropna().unique()
                    
                    # Limita il numero di valori visualizzati se sono troppi
                    if len(unique_values) > 100:
                        st.warning(f"Troppi valori unici ({len(unique_values)}) per {col_name}. Usa il campo di ricerca.")
                        # Aggiungi un campo di ricerca per filtrare i valori
                        search_term = st.text_input(
                            f"Cerca in {col_name}",
                            key=f"search_{col_name}"
                        )
                        if search_term:
                            filtered_values = [val for val in unique_values if search_term.lower() in str(val).lower()]
                            if filtered_values:
                                filters[col_name] = st.multiselect(
                                    f"Valori per {col_name}",
                                    options=filtered_values,
                                    default=None,
                                    key=f"multiselect_{col_name}"
                                )
                            else:
                                st.info(f"Nessun valore trovato per '{search_term}' in {col_name}")
                                filters[col_name] = []
                        else:
                            # Se non c'è un termine di ricerca, mostra i primi 100 valori
                            filters[col_name] = st.multiselect(
                                f"Valori per {col_name} (primi 100)",
                                options=unique_values[:100],
                                default=None,
                                key=f"multiselect_{col_name}"
                            )
                    else:
                        # Se ci sono pochi valori, mostra tutti
                        filters[col_name] = st.multiselect(
                            f"Valori per {col_name}",
                            options=unique_values,
                            default=None,
                            key=f"multiselect_{col_name}"
                        )
    
    # Pulsante per resettare tutti i filtri
    if st.button("Resetta tutti i filtri"):
        return {}
    
    return filters

# Funzione avanzata per applicare i filtri
def apply_advanced_filters(df, filters):
    """Applica i filtri avanzati al DataFrame."""
    if not filters:
        return df
    
    # Crea una copia del DataFrame originale
    filtered_df = df.copy()
    
    # Applica ogni filtro
    for col, filter_value in filters.items():
        if col in filtered_df.columns:
            col_type = filtered_df[col].dtype
            
            # Gestisci diversi tipi di filtri in base al tipo di dati
            if pd.api.types.is_numeric_dtype(col_type) and isinstance(filter_value, tuple) and len(filter_value) == 2:
                # Filtro numerico (intervallo)
                min_val, max_val = filter_value
                filtered_df = filtered_df[(filtered_df[col] >= min_val) & (filtered_df[col] <= max_val)]
            elif pd.api.types.is_datetime64_dtype(col_type) and isinstance(filter_value, tuple) and len(filter_value) == 2:
                # Filtro data (intervallo)
                start_date, end_date = filter_value
                filtered_df = filtered_df[(filtered_df[col].dt.date >= start_date) & (filtered_df[col].dt.date <= end_date)]
            elif isinstance(filter_value, list) and filter_value:
                # Filtro categorico (multiselect)
                filtered_df = filtered_df[filtered_df[col].isin(filter_value)]
    
    return filtered_df

# Titolo dell'applicazione
st.title("VISUALIZZATORE EXCEL")
st.markdown("Carica, visualizza e filtra i tuoi dati Excel in modo semplice e veloce.")

# Sidebar per il caricamento del file e i filtri
with st.sidebar:
    st.header("Caricamento File")
    uploaded_file = st.file_uploader("Carica il tuo file Excel", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        # Mostra informazioni sul file
        file_details = {
            "Nome file": uploaded_file.name,
            "Tipo file": uploaded_file.type,
            "Dimensione": f"{uploaded_file.size / (1024*1024):.2f} MB"
        }
        st.write("Dettagli del file:")
        for k, v in file_details.items():
            st.write(f"- {k}: {v}")

# Contenuto principale
if uploaded_file is not None:
    # Mostra un messaggio di caricamento
    with st.spinner("Caricamento del file in corso..."):
        # Ottieni l'elenco dei fogli
        sheet_names = load_excel_sheets(uploaded_file)
        
        if sheet_names:
            # Selettore del foglio
            selected_sheet = st.selectbox(
                "Seleziona il foglio da visualizzare:",
                options=sheet_names,
                index=0  # Default al primo foglio
            )
            
            # Carica il foglio selezionato con la riga di intestazione appropriata
            start_time = time.time()
            df = load_excel_sheet(uploaded_file, selected_sheet)
            load_time = time.time() - start_time
            
            if df is not None:
                st.success(f"Foglio '{selected_sheet}' caricato con successo in {load_time:.2f} secondi!")
                
                # Mostra informazioni sul DataFrame
                st.subheader("Informazioni sui dati")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Righe", f"{df.shape[0]:,}")
                with col2:
                    st.metric("Colonne", f"{df.shape[1]:,}")
                with col3:
                    st.metric("Memoria utilizzata", f"{df.memory_usage(deep=True).sum() / (1024*1024):.2f} MB")
                
                # Crea tabs per organizzare l'interfaccia
                tab1, tab2 = st.tabs(["Visualizzazione Dati", "Statistiche"])
                
                with tab1:
                    # Crea filtri avanzati per ogni colonna
                    filters = create_advanced_filters(df)
                    
                    # Applica i filtri avanzati
                    filtered_df = apply_advanced_filters(df, filters)
                    
                    # Mostra il numero di righe filtrate
                    st.subheader("Dati filtrati")
                    st.write(f"Visualizzazione di {filtered_df.shape[0]:,} righe su {df.shape[0]:,} totali.")
                    
                    # Opzioni di visualizzazione
                    view_options = st.radio(
                        "Opzioni di visualizzazione:",
                        ["Visualizza tutte le colonne", "Seleziona colonne specifiche"],
                        horizontal=True
                    )
                    
                    if view_options == "Seleziona colonne specifiche":
                        selected_columns = st.multiselect(
                            "Seleziona le colonne da visualizzare",
                            options=df.columns,
                            default=df.columns[:10]  # Mostra le prime 10 colonne di default
                        )
                        if selected_columns:
                            display_df = filtered_df[selected_columns]
                        else:
                            display_df = filtered_df
                    else:
                        display_df = filtered_df
                    
                    # Opzioni di paginazione
                    page_size = st.select_slider(
                        "Righe per pagina:",
                        options=[10, 25, 50, 100, 500, 1000],
                        value=50
                    )
                    
                    # Calcola il numero totale di pagine
                    total_pages = max(1, (len(display_df) + page_size - 1) // page_size)
                    
                    # Selettore di pagina
                    if total_pages > 1:
                        page_number = st.number_input(
                            f"Pagina (1-{total_pages}):",
                            min_value=1,
                            max_value=total_pages,
                            value=1
                        )
                    else:
                        page_number = 1
                    
                    # Calcola l'indice di inizio e fine per la pagina corrente
                    start_idx = (page_number - 1) * page_size
                    end_idx = min(start_idx + page_size, len(display_df))
                    
                    # Mostra i dati della pagina corrente
                    st.dataframe(
                        display_df.iloc[start_idx:end_idx],
                        use_container_width=True,
                        height=600
                    )
                    
                    # Mostra informazioni sulla paginazione
                    st.write(f"Visualizzazione righe {start_idx+1}-{end_idx} di {len(display_df)}")
                    
                    # Opzione per scaricare i dati filtrati
                    col1, col2 = st.columns(2)
                    with col1:
                        # Scarica come CSV
                        csv_data = filtered_df.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label="Scarica dati filtrati come CSV",
                            data=csv_data,
                            file_name=f"dati_filtrati_{selected_sheet}.csv",
                            mime="text/csv"
                        )
                    
                    with col2:
                        # Scarica come Excel
                        try:
                            # Crea un buffer di memoria per il file Excel
                            output = io.BytesIO()
                            # Scrivi il DataFrame nel buffer
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                filtered_df.to_excel(writer, index=False, sheet_name=selected_sheet)
                            # Sposta il cursore all'inizio del buffer
                            output.seek(0)
                            # Ottieni i dati binari
                            excel_data = output.getvalue()
                            
                            st.download_button(
                                label="Scarica dati filtrati come Excel",
                                data=excel_data,
                                file_name=f"dati_filtrati_{selected_sheet}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        except Exception as e:
                            st.error(f"Errore durante la creazione del file Excel: {e}")
                
                with tab2:
                    st.subheader("Statistiche descrittive")
                    
                    # Seleziona colonne numeriche per le statistiche
                    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
                    
                    if numeric_cols:
                        selected_stat_columns = st.multiselect(
                            "Seleziona colonne per le statistiche",
                            options=numeric_cols,
                            default=numeric_cols[:5] if len(numeric_cols) > 5 else numeric_cols
                        )
                        
                        if selected_stat_columns:
                            st.write("Statistiche descrittive:")
                            st.dataframe(
                                filtered_df[selected_stat_columns].describe(),
                                use_container_width=True
                            )
                        else:
                            st.info("Seleziona almeno una colonna numerica per visualizzare le statistiche.")
                    else:
                        st.info("Non sono state trovate colonne numeriche nel dataset.")
                    
                    # Conteggio valori per colonne categoriche
                    categorical_cols = [col for col in df.columns if col not in numeric_cols and df[col].nunique() < 50]
                    
                    if categorical_cols:
                        selected_cat_column = st.selectbox(
                            "Seleziona una colonna categorica per visualizzare il conteggio dei valori",
                            options=categorical_cols
                        )
                        
                        if selected_cat_column:
                            value_counts = filtered_df[selected_cat_column].value_counts().reset_index()
                            value_counts.columns = [selected_cat_column, 'Conteggio']
                            
                            st.write(f"Conteggio valori per '{selected_cat_column}':")
                            st.dataframe(
                                value_counts,
                                use_container_width=True
                            )
                    else:
                        st.info("Non sono state trovate colonne categoriche adatte nel dataset.")
else:
    # Messaggio quando nessun file è caricato
    st.info("👈 Carica un file Excel dalla barra laterale per iniziare.")
    
    # Istruzioni per l'uso
    st.subheader("Come utilizzare questa applicazione:")
    st.markdown("""
    1. **Carica il tuo file Excel** utilizzando il pulsante nella barra laterale
    2. **Seleziona il foglio** da visualizzare se il file contiene più fogli
    3. **Utilizza i filtri avanzati** per filtrare i dati in ogni colonna
    4. **Visualizza i dati** nella tabella interattiva
    5. **Naviga tra le pagine** per esplorare grandi dataset
    6. **Scarica i dati filtrati** in formato CSV o Excel per ulteriori analisi
    
    Questa applicazione è progettata per gestire file Excel di grandi dimensioni e offre funzionalità di filtraggio simili a Excel.
    """)

# Footer
st.markdown("---")
st.markdown("Visualizzatore Excel - Creato con Streamlit")
