import pandas as pd
import os
import re

# =============================================================================
# CONFIGURAZIONE DI DEFAULT (ora qui, in processor)
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
    - raccogliere le zone di no-interference
    - generare i due file di output (chart_config.txt e interferences_summary.txt)
    """

    def __init__(self, excel_path: str, sheet_name: str,
                 output_chart: str, output_summary: str,
                 zone_order=None):
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.output_chart = output_chart
        self.output_summary = output_summary
        self.zone_order = zone_order or DEFAULT_ZONE_ORDER

        self.df_vars = None
        self.df_dyn = None
        self.inter_grouped = []
        self.summary_grouped = []

    def load_data(self):
        """
        Legge il file Excel (supporta sia .xlsx che .xls) e verifica la presenza delle colonne obbligatorie.
        Se il file è .xlsx, usa engine="openpyxl"; se è .xls, usa engine="xlrd".
        """
        if not os.path.isfile(self.excel_path):
            raise FileNotFoundError(f"File non trovato: '{self.excel_path}'")

        # Determino l'estensione del file e scelgo l'engine corretto
        ext = os.path.splitext(self.excel_path)[1].lower()
        if ext == ".xlsx":
            engine = "openpyxl"
        elif ext == ".xls":
            engine = "xlrd"
        else:
            raise ValueError(f"Formato non supportato: '{ext}'. Usa .xlsx o .xls")

        try:
            self.df_vars = pd.read_excel(
                self.excel_path,
                sheet_name=self.sheet_name,
                engine=engine
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
        """Filtra le righe con DataType BOOL e DescrizioneEstensione contenente 'DynamicInterference'."""
        self.df_dyn = self.df_vars[
            (self.df_vars["DataType"] == "BOOL")
            & (self.df_vars["DescrizioneEstensione"].str.contains("DynamicInterference", na=False))
        ].copy()

        if self.df_dyn.empty:
            raise RuntimeError("Nessuna riga con DynamicInterference trovata.")

    @staticmethod
    def estrai_motori_da_root(root: str):
        """
        Da root del tipo "MC4_MotoreA_MC4_MotoreB" restituisce (prefix, motoreA, motoreB).
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
        Genera tag PLC nel formato: <ObjectType>_<DescrizioneEstensione>_<IndexInt>
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
        Estrae due liste di coppie (tag_start, tag_end) per:
          - zone1: "StartNoInterference_…", "EndNoInterference_…"
          - zone2: "StartNoInterference2nd_…", "EndNoInterference2nd_…"
        Considera DescrizioneRadice sia con underscore sia senza.
        """
        motX_no_us = motX.replace("_", "")
        root_with = f"{prefix}_{motX}"
        root_without = f"{prefix}_{motX_no_us}"

        base_mask = (
            ((self.df_vars["DescrizioneRadice"] == root_with)
             | (self.df_vars["DescrizioneRadice"] == root_without))
            & (self.df_vars["DescrizioneEstensione"].str.contains(motY, na=False))
        )
        df_base = self.df_vars[base_mask].copy()

        p1_s = df_base["DescrizioneEstensione"].str.startswith("StartNoInterference_", na=False)
        p1_e = df_base["DescrizioneEstensione"].str.startswith("EndNoInterference_", na=False)
        p2_s = df_base["DescrizioneEstensione"].str.startswith("StartNoInterference2nd_", na=False)
        p2_e = df_base["DescrizioneEstensione"].str.startswith("EndNoInterference2nd_", na=False)

        df_z1 = df_base[p1_s | p1_e]
        df_z2 = df_base[p2_s | p2_e]

        def ordina_accoppia(dfz):
            if dfz.empty:
                return []
            dfz_start = dfz[dfz["DescrizioneEstensione"].str.startswith("StartNoInterference")].sort_values(by="Index")
            dfz_end = dfz[dfz["DescrizioneEstensione"].str.startswith("EndNoInterference")].sort_values(by="Index")
            n = min(len(dfz_start), len(dfz_end))
            coppie = []
            for i in range(n):
                r_s = dfz_start.iloc[i]
                r_e = dfz_end.iloc[i]
                tag_s = self.genera_tag_plc(r_s["ObjectType"], r_s["DescrizioneEstensione"], r_s["Index"])
                tag_e = self.genera_tag_plc(r_e["ObjectType"], r_e["DescrizioneEstensione"], r_e["Index"])
                coppie.append((tag_s, tag_e))
            return coppie

        return ordina_accoppia(df_z1), ordina_accoppia(df_z2)

    def parse_zone_and_index(self, page_name: str):
        """
        Estrae (zone_idx, index_num) dal nome pagina basato su self.zone_order.
        Se non trova numero, assume index_num = 1. Se non trova zona, assume zone_idx = len(zone_order).
        """
        p_lower = page_name.lower()
        for idx, zone in enumerate(self.zone_order):
            zone_lower = zone.lower()
            if zone_lower in p_lower:
                match = re.search(rf"{zone_lower}[_\-]?(\d+)", p_lower, re.IGNORECASE)
                if match:
                    try:
                        num = int(match.group(1))
                    except:
                        num = 1
                    return idx, num
                return idx, 1
        return len(self.zone_order), 1

    def process(self):
        """
        Processo principale:
        - Filtra DynamicInterference
        - Costruisce inter_grouped e summary_grouped
        - Ordina in base a (zone_idx, idx_num)
        """
        self.filter_dynamic_interference()
        self.inter_grouped.clear()
        self.summary_grouped.clear()

        for _, row in self.df_dyn.iterrows():
            root = row["DescrizioneRadice"]
            pagina = str(row["New Page"]).strip()
            if not pagina:
                continue

            try:
                prefix, motA, motB = self.estrai_motori_da_root(root)
            except ValueError:
                continue

            zone1_A, zone2_A = self.raccogli_zone_no_interf(motA, motB, prefix)
            zone1_B, zone2_B = self.raccogli_zone_no_interf(motB, motA, prefix)

            def monta(lista_coppie):
                if not lista_coppie:
                    return ""
                flat = []
                for s, e in lista_coppie:
                    if s: flat.append(s)
                    if e: flat.append(e)
                return ",".join(flat)

            no_int1_A = monta(zone1_A)
            no_int2_A = monta(zone2_A)
            no_int1_B = monta(zone1_B)
            no_int2_B = monta(zone2_B)

            rA = [
                pagina,
                "ChartLeft",
                "",
                "Doughnut",
                "0",
                "360",
                motA,
                "Axe1_RefPosition",
                no_int1_A,
                no_int2_A,
            ]
            rB = [
                pagina,
                "ChartRight",
                "",
                "Doughnut",
                "0",
                "360",
                motB,
                "Axe2_RefPosition",
                no_int1_B,
                no_int2_B,
            ]

            zone_idx, idx_num = self.parse_zone_and_index(pagina)
            self.inter_grouped.append((zone_idx, idx_num, rA, rB))
            self.summary_grouped.append((zone_idx, idx_num, pagina, f"Interferences : {motA}/{motB}"))

        self.inter_grouped.sort(key=lambda x: (x[0], x[1]))
        self.summary_grouped.sort(key=lambda x: (x[0], x[1]))

    def write_chart_config(self):
        """
        Scrive il file chart_config.txt con header + righe tab-separated.
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
            for _, _, rA, rB in self.inter_grouped:
                f.write("\t".join(rA) + "\n")
                f.write("\t".join(rB) + "\n")

    def write_summary(self):
        """
        Scrive il file interferences_summary.txt con header + righe tab-separated.
        """
        with open(self.output_summary, "w", encoding="utf-8") as f2:
            f2.write("pagina\tInterferences\n")
            for _, _, pagina, testo in self.summary_grouped:
                f2.write(f"{pagina}\t{testo}\n")

    def run(self):
        """
        Esegue l'intero flusso:
        - load_data
        - process
        - write_chart_config
        - write_summary
        """
        self.load_data()
        self.process()
        self.write_chart_config()
        self.write_summary()
