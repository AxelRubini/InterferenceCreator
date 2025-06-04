# InterferenceCreator

---

## Italiano 🇮🇹

### Descrizione
**InterferenceCreator** è un’applicazione desktop che genera due file di configurazione a partire da un file Excel contenente le variabili di interferenza tra motori.  
- Il file **chart_config.txt** contiene righe tab-separate necessarie per popolare una sezione di “ChartConfig” nel documento Word.  
- Il file **interferences_summary.txt** elenca, per ciascuna pagina di interferenza, la stringa “Interferences : {MotoreA}/{MotoreB}”.  

### Perché i commenti e i nomi dei metodi sono in italiano
- **Chiarezza per i colleghi**: tutti i membri del team parlano italiano. Un’API che usa nomi come `raccogli_zone_no_interf()` o `estrai_motori_da_root()` è immediatamente comprensibile, senza dover tradurre mentalmente.  
- **Manutenibilità**: commenti e nomi di metodi rispecchiano il vocabolario tecnico che usiamo quotidianamente nel reparto Automazione/PLC.  
- **Allineamento con le specifiche**: gli stakeholder (Team Meccanica, Team Grafica, ecc.) hanno documentazione in italiano. Mantenere la stessa terminologia riduce gli errori di interpretazione.

### Funzionalità chiave
1. **Lettura Excel**  
   - Supporta sia file `.xlsx` (Excel 2007+) sia `.xls` (Excel 97-2003).  
   - Verifica la presenza delle colonne obbligatorie:  
     - `DescrizioneRadice`, `DescrizioneEstensione`, `DataType`, `ObjectType`, `Index`, `New Page`.  

2. **Filtraggio “DynamicInterference”**  
   - Individua tutte le righe di tipo `BOOL` la cui `DescrizioneEstensione` contiene la stringa “DynamicInterference”.  

3. **Estrazione motori e zone di no-interference**  
   - Il metodo `estrai_motori_da_root(root)` divide `root = "MC4_MotoreA_MC4_MotoreB"` in `prefix="MC4"`, `motoreA` e `motoreB`.  
   - `raccogli_zone_no_interf(motX, motY, prefix)` cerca, in “Variabili”, tutti i record “StartNoInterference_…” e “EndNoInterference_…” (incluse versioni “2nd_”) per ciascuna coppia motori.  

4. **Ordinamento e generazione file**  
   - `parse_zone_and_index(page_name)` restituisce `(zona_index, indice_numerico)` in base alla lista di priorità definita dall’utente (es. “Infeed”, “Wheel1”, ecc.).  
   - I record vengono ordinati per `(zona_index, indice_numerico)` e poi esportati in `chart_config.txt` (con header e due righe per ogni coppia motori: `ChartLeft` e `ChartRight`) e in `interferences_summary.txt` (con “pagina” e “Interferences : MotA/MotB”).  

### Struttura del progetto

InterferenceCreator/
├── processor.py # Logica core: lettura Excel, estrazione dati, export file
├── gui.py # Interfaccia Tkinter: definiamo input utente (file, sheet, zone)
├── main.py # Entry point: avvia la GUI
└── README.md # Documentazione bilingue (IT/EN)


#### Riepilogo dei metodi principali (Italiano)
- `load_data()`  
  - Controlla l’esistenza del file Excel.  
  - In base all’estensione seleziona l’engine: `openpyxl` per `.xlsx`, `xlrd` per `.xls`.  
- `filter_dynamic_interference()`  
  - Filtra il DataFrame su `DataType == "BOOL"` e `DescrizioneEstensione.contains("DynamicInterference")`.  
- `estrai_motori_da_root(root)`  
  - Separa il “root” in `prefix`, `motoreA`, `motoreB`.  
- `raccogli_zone_no_interf(motX, motY, prefix)`  
  - Cerca zone di inizio/fine interferenza (no-interf) e restituisce coppie `(tag_start, tag_end)` per zona primaria e secondaria.  
- `parse_zone_and_index(page_name)`  
  - Identifica l’indice di zona e l’indice numerico a partire dal nome pagina (es. “Wheel3_02”→ zona Wheel3, indice 2).  
- `process()`  
  - Componi le liste ordinate `inter_grouped` e `summary_grouped`.  
- `write_chart_config()` / `write_summary()`  
  - Scrivono rispettivamente `chart_config.txt` e `interferences_summary.txt`.

---

## English 🇬🇧

### Description
**InterferenceCreator** is a desktop application that generates two configuration files from an Excel sheet containing interference variables between motors.  
- The **chart_config.txt** file lists tab-separated lines needed to populate a “ChartConfig” section in a Word document.  
- The **interferences_summary.txt** file lists, for each interference page, the string “Interferences : {MotorA}/{MotorB}”.  

### Why comments and method names are in Italian
- **Clarity for colleagues**: the entire team speaks Italian. An API using names like `raccogli_zone_no_interf()` or `estrai_motori_da_root()` is immediately understandable without mental translation.  
- **Maintainability**: comments and method names reflect the technical vocabulary we use daily in the Automation/PLC department.  
- **Alignment with specs**: stakeholders (Mechanical, Graphics teams, etc.) have documentation in Italian. Keeping the same terminology reduces interpretation errors.

### Key Features
1. **Excel Reading**  
   - Supports both `.xlsx` (Excel 2007+) and `.xls` (Excel 97-2003).  
   - Checks the presence of required columns:  
     - `DescrizioneRadice`, `DescrizioneEstensione`, `DataType`, `ObjectType`, `Index`, `New Page`.  

2. **“DynamicInterference” Filtering**  
   - Identifies all rows of type `BOOL` whose `DescrizioneEstensione` contains “DynamicInterference”.  

3. **Motor Extraction and No-Interference Zones**  
   - The method `estrai_motori_da_root(root)` splits `root = "MC4_MotoreA_MC4_MotoreB"` into `prefix="MC4"`, `motoreA` and `motoreB`.  
   - `raccogli_zone_no_interf(motX, motY, prefix)` searches in “Variabili” for all “StartNoInterference_…” and “EndNoInterference_…” records (including “2nd_” versions) for each motor pair.  

4. **Sorting and File Generation**  
   - `parse_zone_and_index(page_name)` returns `(zone_index, numeric_index)` based on a user-defined priority list (e.g. “Infeed”, “Wheel1”, etc.).  
   - Records are sorted by `(zone_index, numeric_index)` and then exported to `chart_config.txt` (with a header and two lines per motor pair: `ChartLeft` and `ChartRight`) and to `interferences_summary.txt` (with “page” and “Interferences : MotA/MotB”).  

### Project Structure
InterferenceCreator/
├── processor.py # Core logic: read Excel, extract data, export files
├── gui.py # Tkinter GUI: define user inputs (file, sheet, zones)
├── main.py # Entry point: launches the GUI
└── README.md # Bilingual documentation (IT/EN)

#### Summary of Main Methods (Italian)
- `load_data()`  
  - Checks for the existence of the Excel file.  
  - Chooses the engine based on extension: `openpyxl` for `.xlsx`, `xlrd` for `.xls`.  
- `filter_dynamic_interference()`  
  - Filters the DataFrame for `DataType == "BOOL"` and `DescrizioneEstensione.contains("DynamicInterference")`.  
- `estrai_motori_da_root(root)`  
  - Splits the “root” into `prefix`, `motoreA`, `motoreB`.  
- `raccogli_zone_no_interf(motX, motY, prefix)`  
  - Finds start/end interference zones (no-interf) and returns `(tag_start, tag_end)` pairs for primary and secondary zones.  
- `parse_zone_and_index(page_name)`  
  - Identifies the zone index and numeric index from the page name (e.g. “Wheel3_02” → zone Wheel3, index 2).  
- `process()`  
  - Builds and sorts the `inter_grouped` and `summary_grouped` lists.  
- `write_chart_config()` / `write_summary()`  
  - Write `chart_config.txt` and `interferences_summary.txt` respectively.

---

### Come utilizzare (IT) / How to Use (EN)

1. **Download repository**
   git clone https://github.com/YourName/InterferenceCreator.git
   cd InterferenceCreator
   
2. **SetUp Venv**
python -m venv .venv
.venv\Scripts\activate          # Windows
source .venv/bin/activate       # Linux/macOS

3. **Install Dependencies**
    pip install pandas openpyxl xlrd==1.2.0

4. **Run** python main.py

Licenza / License
Questo progetto è rilasciato con licenza MIT.
(English: This project is released under the MIT License.)

Nota: La scelta di commenti e nomi dei metodi in italiano è stata fatta per mantenere coerenza con il glossario tecnico interno e migliorare la leggibilità per il team italiano.

Note: Using Italian method names and comments ensures alignment with our internal technical glossary and improves readability for the Italian-speaking development team.












