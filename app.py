import streamlit as st
import pandas as pd
import io
import time

# Configurazione della pagina con tema scuro
st.set_page_config(
    page_title="Visualizzatore Excel Forebet",
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
</style>
""", unsafe_allow_html=True)

# Funzione per caricare il file Excel con gestione ottimizzata della memoria
@st.cache_data(show_spinner=True)
def load_excel(uploaded_file):
    """Carica il file Excel e restituisce un DataFrame pandas."""
    try:
        # Utilizzo di chunk per gestire file di grandi dimensioni
        return pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Errore durante il caricamento del file: {e}")
        return None

# Funzione avanzata per applicare i filtri
def apply_filters(df, filters):
    """Applica i filtri selezionati al DataFrame in modo efficiente."""
    if not filters:
        return df
    
    # Crea una maschera iniziale di True per tutte le righe
    mask = pd.Series(True, index=df.index)
    
    # Applica ogni filtro in sequenza
    for col, values in filters.items():
        if values and col in df.columns:
            # Aggiorna la maschera con l'AND logico
            mask = mask & df[col].isin(values)
    
    # Applica la maschera finale
    return df[mask]

# Titolo dell'applicazione
st.title("Visualizzatore Excel Forebet")
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
        # Carica il file Excel
        start_time = time.time()
        df = load_excel(uploaded_file)
        load_time = time.time() - start_time
        
        if df is not None:
            st.success(f"File caricato con successo in {load_time:.2f} secondi!")
            
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
                # Sidebar per i filtri
                with st.sidebar:
                    st.header("Filtri")
                    st.markdown("Seleziona i filtri da applicare ai dati.")
                    
                    # Inizializza il dizionario dei filtri
                    filters = {}
                    
                    # Aggiungi filtri per le colonne con meno di 50 valori unici
                    filter_columns = [col for col in df.columns if df[col].nunique() < 50 and df[col].nunique() > 1]
                    
                    # Limita il numero di colonne di filtro a massimo 15 per non sovraccaricare l'interfaccia
                    if len(filter_columns) > 15:
                        filter_columns = filter_columns[:15]
                    
                    # Opzione per selezionare quali filtri visualizzare
                    selected_filter_columns = st.multiselect(
                        "Seleziona le colonne da filtrare",
                        options=filter_columns,
                        default=filter_columns[:5] if len(filter_columns) > 5 else filter_columns
                    )
                    
                    # Crea filtri per le colonne selezionate
                    for col in selected_filter_columns:
                        unique_values = sorted(df[col].dropna().unique())
                        if len(unique_values) > 0:
                            filters[col] = st.multiselect(
                                f"Filtra per {col}",
                                options=unique_values,
                                default=None
                            )
                    
                    # Pulsante per resettare i filtri
                    if st.button("Resetta filtri"):
                        filters = {col: [] for col in selected_filter_columns}
                
                # Applica i filtri
                filtered_df = apply_filters(df, filters)
                
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
                total_pages = (len(display_df) + page_size - 1) // page_size
                
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
                    if st.download_button(
                        label="Scarica dati filtrati come CSV",
                        data=filtered_df.to_csv(index=False).encode('utf-8'),
                        file_name=f"dati_filtrati_{uploaded_file.name.split('.')[0]}.csv",
                        mime="text/csv"
                    ):
                        st.success("Download CSV completato!")
                
                with col2:
                    if st.download_button(
                        label="Scarica dati filtrati come Excel",
                        data=io.BytesIO(filtered_df.to_excel(index=False, engine='openpyxl')).getvalue(),
                        file_name=f"dati_filtrati_{uploaded_file.name.split('.')[0]}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ):
                        st.success("Download Excel completato!")
            
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
    2. **Visualizza i dati** nella tabella interattiva
    3. **Applica filtri** selezionando i valori desiderati nella barra laterale
    4. **Naviga tra le pagine** per esplorare grandi dataset
    5. **Scarica i dati filtrati** in formato CSV o Excel per ulteriori analisi
    
    Questa applicazione è progettata per gestire file Excel di grandi dimensioni e offre funzionalità di filtraggio simili a Excel.
    """)

# Footer
st.markdown("---")
st.markdown("Visualizzatore Excel Forebet - Creato con Streamlit")
