import streamlit as st
import pandas as pd
import io
import time
import numpy as np
from itertools import combinations
import base64
import random
from concurrent.futures import ThreadPoolExecutor, as_completed

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
def load_excel_sheet(uploaded_file, sheet_name):
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
        # ma mantieni i tipi numerici e date
        for col in df.columns:
            if df[col].dtype == 'object':
                # Converti i valori NaN in stringhe vuote per evitare errori
                df[col] = df[col].fillna('').astype(str)
        
        return df
    except Exception as e:
        st.error(f"Errore durante il caricamento del foglio {sheet_name}: {e}")
        return None

# Funzione semplificata per creare filtri avanzati
def create_simple_filters(df):
    """Crea filtri semplici per ogni colonna del DataFrame."""
    filters = {}
    
    # Crea un espander per i filtri
    with st.expander("Filtri avanzati (stile Excel)", expanded=False):
        # Limita il numero di colonne per i filtri per evitare sovraccarichi
        max_filter_columns = min(15, len(df.columns))
        filter_columns = list(df.columns)[:max_filter_columns]
        
        # Crea un selettore per scegliere la colonna da filtrare
        selected_column = st.selectbox(
            "Seleziona una colonna da filtrare:",
            options=df.columns,
            index=0
        )
        
        # Determina il tipo di dati della colonna
        col_type = df[selected_column].dtype
        
        # Gestisci diversi tipi di filtri in base al tipo di dati
        if pd.api.types.is_numeric_dtype(col_type):
            # Filtro numerico
            min_val = float(df[selected_column].min()) if not pd.isna(df[selected_column].min()) else 0
            max_val = float(df[selected_column].max()) if not pd.isna(df[selected_column].max()) else 100
            
            # Evita valori min e max identici
            if min_val == max_val:
                max_val = min_val + 1
            
            # Slider per intervallo numerico
            filters[selected_column] = st.slider(
                f"Intervallo per {selected_column}",
                min_value=min_val,
                max_value=max_val,
                value=(min_val, max_val)
            )
        elif pd.api.types.is_datetime64_dtype(col_type):
            # Filtro data
            min_date = df[selected_column].min().date()
            max_date = df[selected_column].max().date()
            
            # Evita date min e max identiche
            if min_date == max_date:
                max_date = min_date + pd.Timedelta(days=1)
            
            # Date input per intervallo di date
            start_date = st.date_input(
                f"Data inizio per {selected_column}",
                value=min_date
            )
            end_date = st.date_input(
                f"Data fine per {selected_column}",
                value=max_date
            )
            filters[selected_column] = (start_date, end_date)
        else:
            # Filtro categorico (multiselect)
            unique_values = sorted(df[selected_column].dropna().unique())
            
            # Aggiungi un campo di ricerca per filtrare i valori
            search_term = st.text_input(
                f"Cerca in {selected_column}"
            )
            
            if search_term:
                filtered_values = [val for val in unique_values if search_term.lower() in str(val).lower()]
                if filtered_values:
                    filters[selected_column] = st.multiselect(
                        f"Valori per {selected_column}",
                        options=filtered_values,
                        default=None
                    )
                else:
                    st.info(f"Nessun valore trovato per '{search_term}' in {selected_column}")
                    filters[selected_column] = []
            else:
                # Limita il numero di valori mostrati
                max_values = min(100, len(unique_values))
                filters[selected_column] = st.multiselect(
                    f"Valori per {selected_column} (primi {max_values})",
                    options=unique_values[:max_values],
                    default=None
                )
    
    # Pulsante per resettare i filtri
    if st.button("Resetta filtri"):
        return {}
    
    return filters

# Funzione avanzata per applicare i filtri
def apply_filters(df, filters):
    """Applica i filtri al DataFrame."""
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

# Funzione per calcolare la percentuale di successo
def calculate_success_percentage(filtered_data, target_column):
    """Calcola la percentuale di successo per una colonna target."""
    if len(filtered_data) == 0:
        return 0, 0
    
    # Calcola la percentuale di successo
    # Assumiamo che la colonna target contenga valori numerici o booleani
    # Se contiene stringhe, convertiamo 'V'/'X' in 1/0
    if pd.api.types.is_numeric_dtype(filtered_data[target_column].dtype):
        success_count = filtered_data[target_column].sum()
    else:
        # Converti stringhe in valori numerici
        success_count = filtered_data[target_column].apply(
            lambda x: 1 if str(x).upper() in ['V', 'TRUE', '1', 'YES', 'Y'] else 0
        ).sum()
    
    total_count = len(filtered_data)
    if total_count > 0:
        return (success_count / total_count) * 100, total_count
    return 0, 0

# Funzione ottimizzata per il filtraggio inverso
def optimized_inverse_filtering(data, target_column, target_percentage, filter_cols, min_matches=50, 
                               min_combinations=1, max_combinations=3, max_results=100, 
                               max_single_results=5, excluded_columns=None, sample_size=None,
                               progress_bar=None):
    """
    Versione ottimizzata della funzione di filtraggio inverso.
    
    Args:
        data: DataFrame con i dati
        target_column: Colonna di risultato target
        target_percentage: Percentuale target da raggiungere
        filter_cols: Lista di colonne di filtro da considerare
        min_matches: Numero minimo di partite per considerare una combinazione valida
        min_combinations: Numero minimo di filtri da combinare
        max_combinations: Numero massimo di filtri da combinare
        max_results: Numero massimo di risultati da restituire
        max_single_results: Numero massimo di risultati singoli per opzione
        excluded_columns: Colonne da escludere dall'analisi
        sample_size: Dimensione del campione da utilizzare (None per utilizzare tutti i dati)
        progress_bar: Barra di progresso Streamlit
    
    Returns:
        Lista di dizionari con le combinazioni di filtri e le relative percentuali
    """
    results = []
    
    # Escludi le colonne specificate
    if excluded_columns:
        filter_cols = [col for col in filter_cols if col not in excluded_columns]
    
    # Campiona i dati se specificato
    if sample_size is not None and sample_size < len(data):
        data_sample = data.sample(sample_size, random_state=42)
    else:
        data_sample = data
    
    # Dizionario per tenere traccia dei risultati per ogni colonna singola
    single_results_count = {}
    
    # Calcola il numero totale di iterazioni per la barra di progresso
    total_iterations = 0
    for n_filters in range(min_combinations, max_combinations + 1):
        # Numero di combinazioni di colonne
        n_col_combinations = len(list(combinations(filter_cols, n_filters)))
        # Numero di tentativi per combinazione
        n_attempts = min(20, 5**n_filters)
        total_iterations += n_col_combinations * n_attempts
    
    # Inizializza la barra di progresso
    if progress_bar is not None:
        progress_bar.progress(0, text="Inizializzazione...")
    
    # Contatore per la barra di progresso
    progress_counter = 0
    
    # Funzione per testare una singola combinazione di filtri
    def test_filter_combination(cols_combo, combo_type, n_filters):
        nonlocal progress_counter
        
        # Crea una griglia di valori ottimizzata per ogni colonna
        grid_values = {}
        for col in cols_combo:
            # Usa percentili per creare una griglia di valori più efficiente
            if pd.api.types.is_numeric_dtype(data_sample[col].dtype):
                # Usa percentili invece di valori unici per ridurre il numero di valori da testare
                percentiles = [10, 20, 30, 40, 50, 60, 70, 80, 90]
                grid_values[col] = [data_sample[col].quantile(p/100) for p in percentiles]
                # Rimuovi duplicati e NaN
                grid_values[col] = [v for v in grid_values[col] if not pd.isna(v)]
                grid_values[col] = list(set(grid_values[col]))
            else:
                # Per colonne non numeriche, usa un campione dei valori unici
                unique_vals = data_sample[col].dropna().unique()
                if len(unique_vals) > 10:
                    grid_values[col] = random.sample(list(unique_vals), min(10, len(unique_vals)))
                else:
                    grid_values[col] = unique_vals
        
        # Numero di tentativi adattivo in base al numero di filtri
        n_attempts = min(20, 5**n_filters)
        
        combination_results = []
        
        # Genera combinazioni di valori per le colonne
        for i in range(n_attempts):
            filter_conditions = []
            combined_condition = pd.Series([True] * len(data_sample), index=data_sample.index)
            
            for col in cols_combo:
                # Scegli casualmente un valore dalla colonna
                if len(grid_values[col]) > 0:
                    val = np.random.choice(grid_values[col])
                    
                    # Crea la condizione
                    if combo_type == '>':
                        condition = data_sample[col] > val
                    else:  # combo_type == '<'
                        condition = data_sample[col] < val
                    
                    filter_conditions.append({
                        'column': col,
                        'operator': combo_type,
                        'value': val
                    })
                    
                    combined_condition = combined_condition & condition
            
            # Applica i filtri e calcola la percentuale
            filtered_data = data_sample[combined_condition]
            if len(filtered_data) >= min_matches:
                success_percentage, count = calculate_success_percentage(filtered_data, target_column)
                if success_percentage >= target_percentage:
                    combination_results.append({
                        'filters': filter_conditions,
                        'percentage': success_percentage,
                        'count': count
                    })
            
            # Aggiorna il contatore di progresso
            progress_counter += 1
            if progress_bar is not None and progress_counter % 10 == 0:
                progress_percentage = min(progress_counter / total_iterations, 1.0)
                progress_bar.progress(progress_percentage, text=f"Analisi in corso... {progress_counter}/{total_iterations} combinazioni testate")
        
        return combination_results
    
    # Analizza combinazioni di filtri da min_combinations a max_combinations
    for n_filters in range(min_combinations, max_combinations + 1):
        # Genera tutte le possibili combinazioni di colonne di filtro
        col_combinations = list(combinations(filter_cols, n_filters))
        
        # Limita il numero di combinazioni di colonne per migliorare le prestazioni
        if len(col_combinations) > 50:
            col_combinations = random.sample(col_combinations, 50)
        
        # Per ogni combinazione di colonne
        for combo_type in ['>', '<']:
            batch_results = []
            
            # Elabora le combinazioni in batch
            for cols_combo in col_combinations:
                batch_results.extend(test_filter_combination(cols_combo, combo_type, n_filters))
            
            # Filtra i risultati per filtri singoli
            if n_filters == 1:
                # Raggruppa per colonna
                column_results = {}
                for result in batch_results:
                    col_name = result['filters'][0]['column']
                    if col_name not in column_results:
                        column_results[col_name] = []
                    column_results[col_name].append(result)
                
                # Prendi i migliori risultati per ogni colonna
                for col_name, col_results in column_results.items():
                    # Ordina per percentuale e numero di campioni
                    col_results.sort(key=lambda x: (-x['percentage'], -x['count']))
                    # Prendi i migliori risultati
                    results.extend(col_results[:max_single_results])
            else:
                # Per combinazioni multiple, aggiungi tutti i risultati
                results.extend(batch_results)
    
    # Completa la barra di progresso
    if progress_bar is not None:
        progress_bar.progress(1.0, text="Analisi completata!")
    
    # Ordina i risultati per percentuale decrescente e poi per numero di partite decrescente
    results.sort(key=lambda x: (-x['percentage'], -x['count']))
    
    # Limita il numero di risultati se specificato
    if max_results is not None:
        return results[:max_results]
    else:
        return results

# Funzione per generare un link di download per un DataFrame
def get_table_download_link(df, filename, text):
    """Genera un link per scaricare il dataframe come file CSV"""
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">📥 {text}</a>'
    return href

# Funzione per generare un link di download per un Excel con più fogli
def get_excel_download_link(dfs_dict, filename, text):
    """Genera un link per scaricare più DataFrame come file Excel con più fogli"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dfs_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    output.seek(0)
    excel_data = output.getvalue()
    b64 = base64.b64encode(excel_data).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">📥 {text}</a>'
    return href

# Funzione per convertire i risultati del filtraggio inverso in DataFrame per l'esportazione
def results_to_export_dfs(results):
    """Converte i risultati del filtraggio inverso in DataFrame per l'esportazione"""
    # Raggruppa i risultati per numero di filtri
    grouped_results = {}
    for result in results:
        n_filters = len(result['filters'])
        if n_filters not in grouped_results:
            grouped_results[n_filters] = []
        grouped_results[n_filters].append(result)
    
    # Crea un DataFrame per ogni gruppo
    export_dfs = {}
    for n_filters, group_results in grouped_results.items():
        if n_filters == 1:
            # Per filtri singoli
            data = []
            for result in group_results:
                filter_info = result['filters'][0]
                data.append({
                    'Opzione': filter_info['column'],
                    'Operatore': filter_info['operator'],
                    'Valore': filter_info['value'],
                    'Risultato %': f"{result['percentage']:.2f}%",
                    'Campioni': result['count']
                })
            export_dfs[f"Combo_{n_filters}"] = pd.DataFrame(data)
        else:
            # Per combinazioni multiple
            data = []
            for result in group_results:
                row = {}
                for i, filter_info in enumerate(result['filters'], 1):
                    row[f'Opzione_{i}'] = filter_info['column']
                    row[f'Operatore_{i}'] = filter_info['operator']
                    row[f'Valore_{i}'] = filter_info['value']
                    if i < len(result['filters']):
                        row[f'AND_{i}'] = '&'
                row['Risultato %'] = f"{result['percentage']:.2f}%"
                row['Campioni'] = result['count']
                data.append(row)
            export_dfs[f"Combo_{n_filters}"] = pd.DataFrame(data)
    
    return export_dfs

# Titolo dell'applicazione
st.title("VISUALIZZATORE EXCEL")

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
    # Crea tabs per separare le funzionalità
    tab1, tab2 = st.tabs(["Visualizzazione e Filtraggio", "Filtraggio Inverso"])
    
    # Mostra un messaggio di caricamento
    with st.spinner("Caricamento del file in corso..."):
        # Ottieni l'elenco dei fogli
        sheet_names = load_excel_sheets(uploaded_file)
    
    if sheet_names:
        # Selettore del foglio (comune a entrambe le tabs)
        selected_sheet = st.selectbox(
            "Seleziona il foglio da visualizzare:",
            options=sheet_names,
            index=0  # Default al primo foglio
        )
        
        # Carica il foglio selezionato con la riga di intestazione appropriata
        start_time = time.time()
        df = load_excel_sheet(uploaded_file, selected_sheet)
        load_time = time.time() - start_time
        
        if df is not None and not df.empty:
            st.success(f"Foglio '{selected_sheet}' caricato con successo in {load_time:.2f} secondi!")
            
            # Tab 1: Visualizzazione e Filtraggio
            with tab1:
                st.markdown("Carica, visualizza e filtra i tuoi dati Excel in modo semplice e veloce.")
                
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
                subtab1, subtab2 = st.tabs(["Visualizzazione Dati", "Statistiche"])
                
                with subtab1:
                    # Inizializza con visualizzazione di tutti i dati (nessun filtro)
                    filtered_df = df.copy()
                    
                    # Crea filtri semplici per ogni colonna
                    filters = create_simple_filters(df)
                    
                    # Applica i filtri solo se sono stati creati
                    if filters:
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
                
                with subtab2:
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
            
            # Tab 2: Filtraggio Inverso
            with tab2:
                st.markdown("### Filtraggio Inverso")
                st.markdown("""
                Questa sezione ti permette di identificare quali combinazioni di filtri portano a percentuali elevate nei mercati selezionati.
                A differenza dell'approccio tradizionale, il filtraggio inverso parte da un target di percentuale desiderato e identifica quali filtri applicare per raggiungere quel target.
                """)
                
                # Identifica le colonne di mercato (da CB a DV) e le colonne di filtro (da AF a BZ)
                market_columns = []
                filter_columns = []
                
                # Identifica le colonne in base alla posizione nel DataFrame
                for i, col in enumerate(df.columns):
                    # Colonne da AF a BZ per i filtri (indici approssimativi 30-75)
                    if i >= 30 and i <= 75:
                        filter_columns.append(col)
                    # Colonne da CB a DV per i mercati (indici approssimativi 80-100)
                    elif i >= 80 and i <= 100:
                        market_columns.append(col)
                
                # Se non sono state trovate colonne, usa un approccio alternativo
                if not market_columns:
                    st.warning("Non sono state trovate colonne di mercato (da CB a DV). Utilizzando le ultime 20 colonne come mercati.")
                    market_columns = df.columns[-20:]
                
                if not filter_columns:
                    st.warning("Non sono state trovate colonne di filtro (da AF a BZ). Utilizzando le colonne da 30 a 75 come filtri.")
                    filter_columns = df.columns[30:76] if len(df.columns) > 75 else df.columns[30:]
                
                # Parametri per il filtraggio inverso
                col1, col2 = st.columns(2)
                
                with col1:
                    # Selezione del mercato
                    target_column = st.selectbox(
                        "Seleziona il mercato",
                        options=market_columns,
                        index=0
                    )
                    
                    # Selezione della percentuale target
                    target_percentage = st.number_input(
                        "Valore target del mercato (%)",
                        min_value=50,
                        max_value=100,
                        value=85,
                        step=1
                    )
                    
                    # Numero minimo di partite
                    min_matches = st.number_input(
                        "Numero minimo di partite",
                        min_value=10,
                        max_value=10000,
                        value=1000,
                        step=100
                    )
                
                with col2:
                    # Numero minimo di combinazioni
                    min_combinations = st.number_input(
                        "Numero minimo di combinazioni",
                        min_value=1,
                        max_value=10,
                        value=1,
                        step=1
                    )
                    
                    # Numero massimo di combinazioni
                    max_combinations = st.number_input(
                        "Numero massimo di combinazioni",
                        min_value=1,
                        max_value=10,
                        value=3,
                        step=1
                    )
                    
                    # Numero massimo di risultati
                    max_results = st.number_input(
                        "Numero massimo di risultati",
                        min_value=1,
                        max_value=1000,
                        value=100,
                        step=10
                    )
                
                # Opzioni avanzate
                with st.expander("Opzioni avanzate"):
                    # Numero massimo di risultati singoli per opzione
                    max_single_results = st.number_input(
                        "Numero massimo di risultati singoli per opzione",
                        min_value=1,
                        max_value=100,
                        value=5,
                        step=1,
                        help="Limita il numero di risultati per ogni singola opzione di filtro per evitare risultati ripetitivi"
                    )
                    
                    # Selezione delle colonne da escludere
                    excluded_columns = st.multiselect(
                        "Colonne da escludere dall'analisi",
                        options=filter_columns,
                        default=None,
                        help="Seleziona le colonne che vuoi escludere dall'analisi"
                    )
                    
                    # Opzione per utilizzare un campione dei dati
                    use_sample = st.checkbox(
                        "Utilizza un campione dei dati per migliorare le prestazioni",
                        value=True,
                        help="Analizza solo un sottoinsieme casuale dei dati per ottenere risultati più velocemente"
                    )
                    
                    # Dimensione del campione
                    if use_sample:
                        sample_size = st.slider(
                            "Dimensione del campione",
                            min_value=1000,
                            max_value=min(20000, len(df)),
                            value=min(5000, len(df)),
                            step=1000,
                            help="Numero di righe da utilizzare per l'analisi (valori più bassi migliorano le prestazioni)"
                        )
                    else:
                        sample_size = None
                
                # Pulsante per eseguire il filtraggio inverso
                if st.button("Esegui Filtraggio Inverso", type="primary", use_container_width=True):
                    # Crea una barra di progresso
                    progress_bar = st.progress(0, text="Inizializzazione...")
                    
                    # Mostra un messaggio di caricamento
                    with st.spinner(f"Ricerca delle combinazioni di filtri per {target_column} > {target_percentage}% con almeno {min_matches} partite..."):
                        # Prepara i dati per il filtraggio inverso
                        # Converti la colonna target in numerica se necessario
                        if pd.api.types.is_numeric_dtype(df[target_column].dtype):
                            # La colonna è già numerica
                            pass
                        else:
                            # Converti la colonna in numerica (assumendo che contenga valori 0/1 o True/False)
                            try:
                                df[target_column] = pd.to_numeric(df[target_column], errors='coerce')
                            except:
                                # Se la conversione fallisce, prova a convertire da stringhe come 'V'/'X' a 1/0
                                if df[target_column].dtype == 'object':
                                    df[target_column] = df[target_column].apply(lambda x: 1 if str(x).upper() in ['V', 'TRUE', '1', 'YES', 'Y'] else 0)
                        
                        # Esegui il filtraggio inverso ottimizzato
                        results = optimized_inverse_filtering(
                            df,
                            target_column,
                            target_percentage,
                            filter_columns,
                            min_matches=min_matches,
                            min_combinations=min_combinations,
                            max_combinations=max_combinations,
                            max_results=max_results,
                            max_single_results=max_single_results,
                            excluded_columns=excluded_columns,
                            sample_size=sample_size,
                            progress_bar=progress_bar
                        )
                        
                        # Mostra i risultati
                        st.header(f"Risultati per {target_column} > {target_percentage}% (min. {min_matches} partite)")
                        
                        if not results:
                            st.warning("Nessuna combinazione di filtri trovata che soddisfi i criteri specificati.")
                        else:
                            st.success(f"Trovate {len(results)} combinazioni di filtri che soddisfano i criteri.")
                            
                            # Raggruppa i risultati per numero di filtri
                            grouped_results = {}
                            for result in results:
                                n_filters = len(result['filters'])
                                if n_filters not in grouped_results:
                                    grouped_results[n_filters] = []
                                grouped_results[n_filters].append(result)
                            
                            # Mostra i risultati raggruppati per numero di filtri
                            for n_filters, group_results in sorted(grouped_results.items()):
                                with st.expander(f"Combinazioni con {n_filters} filtri ({len(group_results)} risultati)", expanded=True):
                                    # Crea una tabella per i risultati
                                    if n_filters == 1:
                                        # Per filtri singoli
                                        data = []
                                        for result in group_results:
                                            filter_info = result['filters'][0]
                                            data.append({
                                                'Opzione': filter_info['column'],
                                                'Operatore': filter_info['operator'],
                                                'Valore': filter_info['value'],
                                                'Risultato %': f"{result['percentage']:.2f}%",
                                                'Campioni': result['count']
                                            })
                                        st.dataframe(pd.DataFrame(data), use_container_width=True)
                                    else:
                                        # Per combinazioni multiple
                                        data = []
                                        for result in group_results:
                                            row = {}
                                            for i, filter_info in enumerate(result['filters'], 1):
                                                row[f'Opzione {i}'] = filter_info['column']
                                                row[f'Op. {i}'] = filter_info['operator']
                                                row[f'Valore {i}'] = filter_info['value']
                                            row['Risultato %'] = f"{result['percentage']:.2f}%"
                                            row['Campioni'] = result['count']
                                            data.append(row)
                                        st.dataframe(pd.DataFrame(data), use_container_width=True)
                            
                            # Prepara i dati per l'esportazione
                            export_dfs = results_to_export_dfs(results)
                            
                            # Opzioni di download
                            st.subheader("Scarica i risultati")
                            col1, col2 = st.columns(2)
                            
                            with col1:
                                # Scarica come CSV (solo il primo foglio)
                                first_sheet = list(export_dfs.keys())[0]
                                csv_data = export_dfs[first_sheet].to_csv(index=False).encode('utf-8')
                                st.download_button(
                                    label="Scarica risultati come CSV",
                                    data=csv_data,
                                    file_name=f"filtraggio_inverso_{target_column}_{target_percentage}.csv",
                                    mime="text/csv"
                                )
                            
                            with col2:
                                # Scarica come Excel (tutti i fogli)
                                try:
                                    # Crea un buffer di memoria per il file Excel
                                    output = io.BytesIO()
                                    # Scrivi i DataFrame nel buffer
                                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                        for sheet_name, df in export_dfs.items():
                                            df.to_excel(writer, sheet_name=sheet_name, index=False)
                                    # Sposta il cursore all'inizio del buffer
                                    output.seek(0)
                                    # Ottieni i dati binari
                                    excel_data = output.getvalue()
                                    
                                    st.download_button(
                                        label="Scarica risultati come Excel",
                                        data=excel_data,
                                        file_name=f"filtraggio_inverso_{target_column}_{target_percentage}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                                except Exception as e:
                                    st.error(f"Errore durante la creazione del file Excel: {e}")
                else:
                    # Istruzioni per l'uso
                    st.info("""
                    ### Come utilizzare il Filtraggio Inverso
                    
                    1. **Seleziona il mercato** che ti interessa analizzare
                    2. **Imposta il valore target** in percentuale che desideri raggiungere
                    3. **Specifica il numero minimo di partite** per considerare valida una combinazione
                    4. **Imposta il numero minimo e massimo di combinazioni** di filtri da considerare
                    5. **Utilizza le opzioni avanzate** per migliorare le prestazioni (consigliato per file grandi)
                    6. **Clicca su "Esegui Filtraggio Inverso"** per avviare l'analisi
                    
                    L'applicazione cercherà combinazioni di filtri che producono una percentuale di successo superiore al target specificato.
                    
                    **Suggerimento per prestazioni migliori**: Utilizza un campione dei dati (opzione avanzata) per ottenere risultati più velocemente.
                    """)
        else:
            st.error(f"Errore: Il foglio '{selected_sheet}' è vuoto o non è stato caricato correttamente.")
else:
    # Messaggio quando nessun file è caricato
    st.info("👈 Carica un file Excel dalla barra laterale per iniziare.")
    
    # Istruzioni per l'uso
    st.subheader("Come utilizzare questa applicazione:")
    st.markdown("""
    1. **Carica il tuo file Excel** utilizzando il pulsante nella barra laterale
    2. **Seleziona il foglio** da visualizzare se il file contiene più fogli
    3. **Scegli la modalità** che desideri utilizzare:
       - **Visualizzazione e Filtraggio**: Per visualizzare e filtrare i dati in modo tradizionale
       - **Filtraggio Inverso**: Per trovare combinazioni di filtri che producono percentuali elevate
    
    ### Visualizzazione e Filtraggio
    - Utilizza i filtri avanzati per filtrare i dati in ogni colonna
    - Visualizza i dati nella tabella interattiva
    - Naviga tra le pagine per esplorare grandi dataset
    - Scarica i dati filtrati in formato CSV o Excel
    
    ### Filtraggio Inverso
    - Seleziona il mercato e il valore target desiderato
    - Specifica il numero minimo di partite e di combinazioni
    - L'applicazione troverà combinazioni di filtri che producono percentuali elevate
    - Scarica i risultati in formato CSV o Excel
    
    **Nota sulle prestazioni**: Per file di grandi dimensioni, utilizza le opzioni avanzate nel Filtraggio Inverso per migliorare le prestazioni.
    """)

# Footer
st.markdown("---")
st.markdown("Visualizzatore Excel - Creato con Streamlit")
