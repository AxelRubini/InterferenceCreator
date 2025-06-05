# processor.py

import pandas as pd
import os
import re

# =============================================================================
# CONFIGURAZIONE DI DEFAULT
# =============================================================================
DEFAULT_ZONE_ORDER = [
    "Infeed",
    "Wheel1",
    "Wheel2",
    "Wheel3",
    "Exit",
    "Stamp",
    "InnerLiner",
    "OuterLiner",
]


class InterferenceProcessor:
    """
    Classe responsabile di:
    - leggere il file Excel (.xlsx o .xls)
    - filtrare le righe di DynamicInterference
    - raccogliere le zone di no-interferenza (qualunque prefisso ‘StartNo…’/‘EndNo…’)
      fino a due zone (ordinal 1 e 2)
    - generare due file di output:
        * chart_config.txt (NoInterf1 contiene tag 1st+2nd)
        * interferences_summary.txt
      fino a tre grafici per pagina: ChartLeft, ChartRight, ChartCenter.
    """

    def __init__(
        self,
        excel_path: str,
        sheet_name: str,
        output_chart: str,
        output_summary: str,
        zone_order=None,
    ):
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.output_chart = output_chart
        self.output_summary = output_summary
        self.zone_order = zone_order or DEFAULT_ZONE_ORDER

        self.df_vars = None
        self.df_dyn = None
        self.inter_grouped = []   # lista di tuple (zone_idx, idx_num, riga_chart)
        self.summary_grouped = [] # lista di tuple (zone_idx, idx_num, pagina, "Interferences : A/B")

    def load_data(self):
        """
        Legge il file Excel (supporta .xlsx e .xls) e verifica la presenza delle colonne obbligatorie.
        """
        if not os.path.isfile(self.excel_path):
            raise FileNotFoundError(f"File non trovato: '{self.excel_path}'")

        ext = os.path.splitext(self.excel_path)[1].lower()
        if ext == ".xlsx":
            engine = "openpyxl"
        elif ext == ".xls":
            engine = "xlrd"
        else:
            raise ValueError(f"Formato non supportato: '{ext}'. Usa .xlsx o .xls")

        try:
            self.df_vars = pd.read_excel(
                self.excel_path, sheet_name=self.sheet_name, engine=engine
            )
        except Exception as e:
            raise ValueError(f"Errore apertura foglio '{self.sheet_name}' ({ext}): {e}")

        required_cols = [
            "DescrizioneRadice",
            "DescrizioneEstensione",
            "DataType",
            "ObjectType",
            "Index",
            "New Page",
        ]
        missing = [c for c in required_cols if c not in self.df_vars.columns]
        if missing:
            raise ValueError(f"Mancano colonne nel foglio '{self.sheet_name}': {missing}")

    def filter_dynamic_interference(self):
        """
        Filtra le righe con DataType == "BOOL" e DescrizioneEstensione contenente 'DynamicInterference'.
        """
        self.df_dyn = self.df_vars[
            (self.df_vars["DataType"] == "BOOL")
            & (self.df_vars["DescrizioneEstensione"].str.contains("DynamicInterference", na=False))
        ].copy()

        if self.df_dyn.empty:
            raise RuntimeError("Nessuna riga con DynamicInterference trovata.")

    @staticmethod
    def estrai_motori_da_root(root: str):
        """
        Da root come "MC4_MOTOREA_MC4_MOTOREB" restituisce (prefix, motoreA, motoreB).
        Se il formato non è corretto solleva ValueError.
        """
        tokens = root.split("_")
        if len(tokens) < 4:
            raise ValueError(f"Root non conforme: '{root}'")
        prefix = tokens[0]
        idx = next((i for i in range(1, len(tokens)) if tokens[i] == prefix), None)
        if idx is None or idx == len(tokens) - 1:
            raise ValueError(f"Formato root invalido: '{root}'")
        motA = "_".join(tokens[1:idx])
        motB = "_".join(tokens[idx + 1 :])
        return prefix, motA, motB

    @staticmethod
    def genera_tag_plc(obj_type: str, descr_ext: str, idx_val):
        """
        Genera tag PLC nel formato: <ObjectType>_<DescrizioneEstensione>_<IndexInt>.
        Se idx_val è NaN o non convertibile, ritorna stringa vuota.
        """
        primo_tipo = obj_type.split(";")[0].strip() if isinstance(obj_type, str) else ""
        if pd.isna(idx_val):
            return ""
        try:
            idx_int = int(idx_val)
        except Exception:
            return ""
        return f"{primo_tipo}_{descr_ext}_{idx_int}"

    def raccogli_zone_no_interf(self, motX: str, motY: str, prefix: str):
        """
        Estrae due liste di coppie (tag_start, tag_end) per le prime due “zone”:
          - include qualunque DescrizioneEstensione che inizia con "StartNo" / "EndNo"
            e contiene il nome del motore motY.
          - Riconosce ordinali nel nome: “1st”, “2nd”, “3rd”…
            * Se non trova “1st”/“2nd”/“3rd”, assume ordinal=1.
          - Restituisce due liste: (zone1_list, zone2_list), ognuna come [(tag_s, tag_e), …],
            con ordinal == 1 e ordinal == 2. Eventuali ordinal >= 3 vengono ignorati.
        """
        motX_no_us = motX.replace("_", "")
        root_with = f"{prefix}_{motX}"
        root_without = f"{prefix}_{motX_no_us}"

        # Filtra le righe che appartenengono al root (con o senza underscore)
        mask_base = (
            ((self.df_vars["DescrizioneRadice"] == root_with)
             | (self.df_vars["DescrizioneRadice"] == root_without))
            & (self.df_vars["DescrizioneEstensione"].str.contains(motY, na=False))
        )
        df_base = self.df_vars[mask_base].copy()

        # Seleziona tutte le righe "StartNo..." e "EndNo..."
        df_start = df_base[df_base["DescrizioneEstensione"].str.startswith("StartNo", na=False)].copy()
        df_end   = df_base[df_base["DescrizioneEstensione"].str.startswith("EndNo", na=False)].copy()

        # Funzione per estrarre ordinal: cerca "1st", "2nd", "3rd" (case-insensitive),
        # altrimenti ritorna 1.
        def estrai_ordinal(descr_ext: str) -> int:
            m = re.search(r"(?i)(\d+)(?:st|nd|rd|th)", descr_ext)
            if m:
                try:
                    return int(m.group(1))
                except:
                    return 1
            # Se non trova suffisso numerico, consideralo prima zona (ordinal=1)
            return 1

        # Costruisci lista di tuple (ordinal, index, tag) per start ed end
        starts = []
        for _, r in df_start.iterrows():
            ord_val = estrai_ordinal(r["DescrizioneEstensione"])
            starts.append((ord_val, r["Index"], r))

        ends = []
        for _, r in df_end.iterrows():
            ord_val = estrai_ordinal(r["DescrizioneEstensione"])
            ends.append((ord_val, r["Index"], r))

        # Ordina per ordinal ASC, poi per Index ASC
        starts.sort(key=lambda x: (x[0], x[1]))
        ends.sort(key=lambda x: (x[0], x[1]))

        # Appiattisci i record raggruppando per ordinal 1st e 2nd
        zone1 = []
        zone2 = []

        # Per ogni ordinal 1 o 2, prendi la prima coppia start/end se esistono
        for ordinal_target, zone_list in [(1, zone1), (2, zone2)]:
            # Filtra tutti i start con quell'ordinal_target
            starts_ord = [row for (ordv, _, row) in starts if ordv == ordinal_target]
            ends_ord   = [row for (ordv, _, row) in ends   if ordv == ordinal_target]

            # Prendi il minimo tra len(starts_ord) e len(ends_ord)
            n_coppie = min(len(starts_ord), len(ends_ord))
            for i in range(n_coppie):
                r_s = starts_ord[i]
                r_e = ends_ord[i]
                tag_s = self.genera_tag_plc(r_s["ObjectType"], r_s["DescrizioneEstensione"], r_s["Index"])
                tag_e = self.genera_tag_plc(r_e["ObjectType"], r_e["DescrizioneEstensione"], r_e["Index"])
                zone_list.append((tag_s, tag_e))

        return zone1, zone2

    def parse_zone_and_index(self, page_name: str):
        """
        Restituisce (zone_idx, index_num) basato su self.zone_order e numero finale.
        - Se es. 'Wheel1' senza indice => (idx di Wheel1, 1)
        - Se manca completamente => (len(zone_order), 1)
        """
        p_lower = page_name.lower()
        for idx, zone in enumerate(self.zone_order):
            zone_lower = zone.lower()
            if zone_lower in p_lower:
                match = re.search(rf"{zone_lower}[_\-]?(\d+)", p_lower, re.IGNORECASE)
                if match:
                    try:
                        return idx, int(match.group(1))
                    except:
                        return idx, 1
                return idx, 1
        return len(self.zone_order), 1

    def process(self):
        """
        Processo principale:
        - Filtra DynamicInterference
        - Per ogni coppia motori genera fino a 3 grafici (ChartLeft/Right/Center)
        - Monta NoInterf1 come concatenazione di tutti i tag 1st e 2nd
        - Ordina in base a (zone_idx, idx_num)
        """
        self.filter_dynamic_interference()
        self.inter_grouped.clear()
        self.summary_grouped.clear()

        page_chart_counter = {}

        for _, row in self.df_dyn.iterrows():
            root = row["DescrizioneRadice"]
            pagina = str(row["New Page"]).strip()
            if not pagina:
                continue

            try:
                prefix, motA, motB = self.estrai_motori_da_root(root)
            except ValueError:
                continue

            # Itera sui due motori: per ognuno assegna il grafico in base a quante volte compare pagina
            for motore, function_type in [(motA, "Axe1_RefPosition"), (motB, "Axe2_RefPosition")]:
                if pagina not in page_chart_counter:
                    page_chart_counter[pagina] = 0

                idx_counter = page_chart_counter[pagina]
                if idx_counter == 0:
                    chart_name = "ChartLeft"
                elif idx_counter == 1:
                    chart_name = "ChartRight"
                elif idx_counter == 2:
                    chart_name = "ChartCenter"
                else:
                    # Ignora motori oltre il terzo per la stessa pagina
                    print(
                        f"[WARNING] Pagina '{pagina}' già ha 3 grafici; "
                        f"skipping motore '{motore}'."
                    )
                    continue

                # Raccogli tag 1st e 2nd (nomi arbitrari "No…") e concatena in NoInterf1
                motY = motB if motore == motA else motA
                zone1, zone2 = self.raccogli_zone_no_interf(motore, motY, prefix)

                all_zones = zone1 + zone2

                def monta(lista_coppie):
                    if not lista_coppie:
                        return ""
                    flat = []
                    for s, e in lista_coppie:
                        if s:
                            flat.append(s)
                        if e:
                            flat.append(e)
                    return ",".join(flat)

                no_interf1 = monta(all_zones)
                no_interf2 = ""  # rimane vuoto

                riga_chart = [
                    pagina,        # pagina
                    chart_name,    # nome: ChartLeft/ChartRight/ChartCenter
                    "",            # visiblePlc (vuoto)
                    "Doughnut",    # Type
                    "0",           # Rotation
                    "360",         # Period
                    motore,        # Title
                    function_type, # FunctionType
                    no_interf1,    # NoInterf1 (1st+2nd tag)
                    no_interf2,    # NoInterf2 (vuoto)
                ]

                zone_idx, idx_num = self.parse_zone_and_index(pagina)
                self.inter_grouped.append((zone_idx, idx_num, riga_chart))

                # Aggiungi summary **solo** per motA (per evitare duplicati)
                if motore == motA:
                    summary_str = f"Interferences : {motA}/{motB}"
                    self.summary_grouped.append((zone_idx, idx_num, pagina, summary_str))

                page_chart_counter[pagina] += 1

        # Ordina i risultati
        self.inter_grouped.sort(key=lambda x: (x[0], x[1]))
        self.summary_grouped.sort(key=lambda x: (x[0], x[1]))

    def write_chart_config(self):
        """
        Scrive chart_config.txt con header e righe tab-separated.
        """
        header = [
            "pagina",
            "nome",
            "visiblePlc",
            "Type",
            "Rotation",
            "Period",
            "Title",
            "FunctionType",
            "NoInterf1",
            "NoInterf2",
        ]
        with open(self.output_chart, "w", encoding="utf-8") as f:
            f.write("\t".join(header) + "\n")
            for _, _, riga in self.inter_grouped:
                f.write("\t".join(riga) + "\n")

    def write_summary(self):
        """
        Scrive interferences_summary.txt con header e righe tab-separated.
        """
        with open(self.output_summary, "w", encoding="utf-8") as f2:
            f2.write("pagina\tInterferences\n")
            for _, _, pagina, testo in self.summary_grouped:
                f2.write(f"{pagina}\t{testo}\n")

    def run(self):
        """
        Flusso completo:
        1. load_data()
        2. process()
        3. write_chart_config()
        4. write_summary()
        """
        self.load_data()
        self.process()
        self.write_chart_config()
        self.write_summary()
