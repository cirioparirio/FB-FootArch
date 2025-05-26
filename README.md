# Visualizzatore Excel Forebet

Un'applicazione web Streamlit per caricare, visualizzare e filtrare file Excel di grandi dimensioni.

## Caratteristiche

- **Caricamento file**: Supporta il caricamento di file Excel (.xlsx, .xls) aggiornati quotidianamente
- **Visualizzazione dati**: Interfaccia tabellare reattiva con paginazione per gestire file di grandi dimensioni
- **Filtri avanzati**: Funzionalità di filtraggio simili a Excel per colonne selezionate
- **Tema scuro**: Interfaccia con tema scuro per una migliore esperienza visiva
- **Download dati**: Possibilità di scaricare i dati filtrati in formato CSV o Excel
- **Statistiche**: Visualizzazione di statistiche descrittive sui dati numerici
- **Ottimizzato per dispositivi mobili**: Interfaccia responsive per l'utilizzo da smartphone

## Requisiti

- Python 3.7 o superiore
- Streamlit
- Pandas
- Openpyxl

## Installazione

1. Clona questa repository:
```bash
git clone https://github.com/tuousername/visualizzatore-excel-forebet.git
cd visualizzatore-excel-forebet
```

2. Installa le dipendenze:
```bash
pip install -r requirements.txt
```

## Utilizzo

1. Avvia l'applicazione:
```bash
streamlit run app.py
```

2. Apri il browser all'indirizzo indicato (di solito http://localhost:8501)

3. Carica il tuo file Excel utilizzando il pulsante nella barra laterale

4. Utilizza i filtri nella barra laterale per filtrare i dati

5. Naviga tra le pagine per esplorare i dati

6. Scarica i dati filtrati in formato CSV o Excel

## Deployment su Streamlit Cloud

Per rendere l'applicazione accessibile online:

1. Crea un account su [Streamlit Cloud](https://streamlit.io/cloud)

2. Collega il tuo repository GitHub

3. Seleziona il file `app.py` come punto di ingresso

4. Avvia il deployment

5. L'applicazione sarà accessibile tramite un URL pubblico

## Utilizzo da smartphone

1. Accedi all'URL dell'applicazione deployata dal browser del tuo smartphone

2. Carica il file Excel aggiornato

3. Utilizza i filtri e la visualizzazione ottimizzata per dispositivi mobili

4. Scarica i dati filtrati se necessario

## Struttura del progetto

- `app.py`: Applicazione Streamlit principale
- `requirements.txt`: Dipendenze del progetto
- `README.md`: Documentazione del progetto

## Personalizzazione

Se la struttura delle colonne del file Excel dovesse cambiare in futuro, sarà necessario modificare il codice in `app.py` per adattarlo alla nuova struttura.

## Contributi

Contributi, segnalazioni di bug e richieste di funzionalità sono benvenuti. Sentiti libero di aprire una issue o una pull request.

## Licenza

Questo progetto è distribuito con licenza MIT. Vedi il file `LICENSE` per maggiori dettagli.
