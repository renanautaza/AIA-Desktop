# -*- coding: utf-8 -*-
"""
AUTAZA INSPECTION ANALYTICS — V1.1.6
Industrial Inspection Analytics Platform

V0.9.0
- Modo Expectativa / Critério de Entrega
- Metas de Accuracy, Precision, Recall, F1, FP/100 e FN/100
- Status META ATINGIDA / EM CALIBRAÇÃO / FORA DA EXPECTATIVA
- Rastreabilidade da origem da meta
- Tooltips técnicos com fórmulas ao passar o mouse sobre métricas
- Comparação multicoleta integrada ao critério de expectativa

V0.3.3
- Coleta Rápida sem classificação de defeito
- Coleta detalhada opcional com múltiplos tipos de defeito por inspeção
- Separação entre:
    1) classificação da inspeção (TP/TN/FP/FN por carroceria/peça)
    2) detecção de ocorrências de defeitos (real x detectado)
- Importação de Excel compatível com planilhas históricas
- Exportação Excel completo
- Exportação CSV estruturado
- Salvamento de projeto JSON
- Normalização de tipos de defeito
- Compatibilidade retroativa com projetos V0.3.2
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime
from collections import defaultdict
import json, math, csv, re, unicodedata, os, sys, hashlib

try:
    import openpyxl
except Exception:
    openpyxl = None

APP = "Autaza Inspection Analytics"
VERSION = "1.1.6"


def application_dir():
    """Return the directory containing bundled resources or the source file."""
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent

BG = "#07131C"
PANEL = "#0E202D"
PANEL2 = "#123B52"
PANEL3 = "#091A25"
BORDER = "#21485F"
TEXT = "#F3F7FA"
MUTED = "#8FA9B8"
ACCENT = "#00A9E8"
ACCENT2 = "#67D5FF"
GOOD = "#45D39A"
WARN = "#F4C15D"
BAD = "#FF6B78"

REFS = ["Defeito confirmado", "OK confirmado", "Não determinado"]
GABARITOS = ["Produção", "Técnico", "Produção + Técnico", "Outro", "Não determinado"]

GLOSSARY = {
    "N — Número de inspeções":
        "Quantidade total de inspeções classificadas na amostra: N = TP + TN/VN + FP + FN. Não representa quantidade de defeitos; representa a unidade estatística de inspeções.",
    "TP — Verdadeiro Positivo":
        "Na classificação da inspeção: a peça/carroceria realmente possui defeito e o Vision AI classifica como NOK.",
    "TN/VN — Verdadeiro Negativo":
        "Na classificação da inspeção: a peça/carroceria realmente está OK e o Vision AI classifica como OK.",
    "FP — Falso Positivo (Classificação)":
        "Na classificação da inspeção: a peça/carroceria está OK, mas o Vision AI classifica como NOK. Mede erro no resultado OK/NOK da inspeção.",
    "FN — Falso Negativo (Classificação)":
        "Na classificação da inspeção: existe defeito real, mas o Vision AI classifica como OK. Mede escape no resultado OK/NOK da inspeção.",
    "FP — Falso Positivo (Detecção)":
        "Na análise de ocorrências: uma detecção registrada não encontra defeito real correspondente. Nesta versão, quando não existe correspondência individual, a estimativa por contagem usa max(Detectado−Real, 0).",
    "FN — Falso Negativo (Detecção)":
        "Na análise de ocorrências: existe defeito real que não aparece nas detecções. Quando não existe correspondência individual, a estimativa por contagem usa max(Real−Detectado, 0).",
    "Accuracy — Acurácia":
        "(TP+TN)/(TP+TN+FP+FN). Percentual de inspeções classificadas corretamente.",
    "Precision — Precisão":
        "TP/(TP+FP). Entre as inspeções classificadas como NOK, proporção que realmente possuía defeito.",
    "Recall / Sensibilidade":
        "TP/(TP+FN). Entre as inspeções realmente NOK, proporção classificada como NOK.",
    "Specificity — Especificidade":
        "TN/(TN+FP). Entre as inspeções realmente OK, proporção classificada como OK.",
    "FPR — False Positive Rate":
        "FP/(FP+TN). Taxa clássica de falso positivo na classificação das inspeções.",
    "FNR — False Negative Rate":
        "FN/(FN+TP). Taxa clássica de falso negativo na classificação das inspeções.",
    "F1-Score":
        "Média harmônica entre Precision e Recall.",
    "MCC":
        "Matthews Correlation Coefficient. Indicador de correlação da matriz de confusão.",
    "Threshold / Grau de Confiança":
        "Limite usado pelo sistema de visão para aceitar uma detecção.",
    "Método de referência":
        "Método usado para determinar a verdade do caso: técnico, produção, ambos etc.",
    "Ocorrência de defeito":
        "Quantidade de defeitos reais ou detectados dentro de uma inspeção. Pode ser registrada sem informar o tipo.",
    "TP-D / FP-D / FN-D":
        "Indicadores de detecção de ocorrências. São calculados separadamente da classificação OK/NOK; quando não há correspondência individual, usam uma estimativa por contagem.",
    "TP-C / VN-C / FP-C / FN-C":
        "Indicadores da classificação da inspeção, isto é, o resultado OK/NOK da carroceria, peça ou inspeção.",
    "Prevalência":
        "Proporção de inspeções realmente NOK dentro da amostra avaliada. É um indicador de composição da amostra, não de qualidade do sistema.",
    "Taxa de alarme NOK":
        "Proporção de inspeções classificadas pelo sistema como NOK. Indicador operacional complementar.",
    "FPR / FNR":
        "FPR mede a fração de casos realmente OK classificados como NOK; FNR mede a fração de casos realmente NOK classificados como OK.",
    "Expectativa / Meta de Entrega":
        "Critério de desempenho definido para orientar a calibração e a entrega. Accuracy, Precision, Recall e F1 usam meta mínima; FP/100 e FN/100 usam limite máximo.",
    "FP / 100 inspeções":
        "Fórmula: (FP / N) × 100. Normaliza a carga de falsos positivos de classificação para 100 inspeções.",
    "FN / 100 inspeções":
        "Fórmula: (FN / N) × 100. Normaliza a carga de falsos negativos de classificação para 100 inspeções.",
    "Precision de detecção":
        "Fórmula: TP-D / (TP-D + FP-D). Mede a confiabilidade das ocorrências detectadas.",
    "Recall de detecção":
        "Fórmula: TP-D / (TP-D + FN-D). Mede a capacidade de encontrar ocorrências reais.",
}

def safe(a, b):
    return None if b == 0 else a / b

def pct(v):
    return "N/D" if v is None else f"{v*100:.2f}%"

def classification_metrics(tp, tn, fp, fn):
    n = tp + tn + fp + fn
    precision = safe(tp, tp + fp)
    recall = safe(tp, tp + fn)
    specificity = safe(tn, tn + fp)
    fpr = safe(fp, fp + tn)
    fnr = safe(fn, fn + tp)
    npv = safe(tn, tn + fn)
    accuracy = safe(tp + tn, n)
    bal = None if recall is None or specificity is None else (recall + specificity) / 2
    f1 = None if precision is None or recall is None or precision + recall == 0 else 2 * precision * recall / (precision + recall)
    beta = 2.0
    f2 = None if precision is None or recall is None or (beta**2 * precision + recall) == 0 else (1+beta**2)*precision*recall/(beta**2*precision+recall)
    den = (tp+fp)*(tp+fn)*(tn+fp)*(tn+fn)
    mcc = None if den <= 0 else (tp*tn-fp*fn)/math.sqrt(den)
    return {
        "N": n, "TP": tp, "VN/TN": tn, "FP": fp, "FN": fn,
        "Accuracy": accuracy, "Precision": precision,
        "Recall / Sensibilidade": recall, "Specificity": specificity,
        "FPR": fpr, "FNR": fnr, "NPV": npv,
        "Balanced Accuracy": bal, "F1": f1, "F-beta (β=2)": f2, "MCC": mcc
    }

def normalize_text(value):
    if value is None:
        return ""
    s = str(value).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s)
    return s.upper().strip()

def normalize_defect(value):
    s = normalize_text(value)
    if not s or s in {"0", "-", "N/A", "NA", "N/D", "SEM DEFEITO", "SEM DEFEITO DETECTADO"}:
        return ""
    aliases = {
        "AMASSADO": "AMASSADO",
        "AMASS.": "AMASSADO",
        "AMASSADO/DEF.": "AMASSADO / DEFORMACAO",
        "AMASSADO / DEF.": "AMASSADO / DEFORMACAO",
        "AMASSADO-DEFORMACAO": "AMASSADO / DEFORMACAO",
        "DEFORMAÇÃO/AMASSADO": "DEFORMACAO / AMASSADO",
        "DEFORMACAO/AMASSADO": "DEFORMACAO / AMASSADO",
    }
    return aliases.get(s, s)

def blank_row():
    return {
        "id": "", "vision": "OK", "ref": "Não determinado",
        "real": 0, "det": 0,
        "tp": 0, "vn": 0, "fp": 0, "fn": 0,
        "obs": "", "calc_status": "N/D",
        "defects": [], "source": "Manual"
    }

def normalize_row(r):
    x = blank_row()
    x.update(r or {})
    x["id"] = str(x.get("id") or "")
    x["vision"] = str(x.get("vision") or "OK").upper()
    x["ref"] = str(x.get("ref") or "Não determinado")
    for k in ("real","det","tp","vn","fp","fn"):
        try:
            x[k] = max(0, int(float(x.get(k) or 0)))
        except Exception:
            x[k] = 0
    x["obs"] = str(x.get("obs") or "")
    x["calc_status"] = str(x.get("calc_status") or "N/D")
    x["source"] = str(x.get("source") or "Manual")
    x["detection_available"] = bool(x.get("detection_available", True))
    raw = x.get("defects") or []
    if not isinstance(raw, list):
        raw = []
    defects = []
    for d in raw:
        if not isinstance(d, dict):
            continue
        name = normalize_defect(d.get("type") or d.get("defeito"))
        if not name:
            continue
        try: rr = max(0, int(float(d.get("real") or 0)))
        except Exception: rr = 0
        try: dd = max(0, int(float(d.get("det") or d.get("detectado") or 0)))
        except Exception: dd = 0
        defects.append({"type": name, "real": rr, "det": dd})
    x["defects"] = defects
    return x

TOOLTIP_DEFS = {
    "N • Inspeções": ("N — Número de inspeções", "N = TP + TN/VN + FP + FN. Unidade estatística: inspeção/carroceria/peça."),
    "TP • Inspeções": ("TP — Verdadeiro Positivo", "TP = inspeções realmente NOK e classificadas como NOK."),
    "TN/VN • Inspeções": ("TN/VN — Verdadeiro Negativo", "TN/VN = inspeções realmente OK e classificadas como OK."),
    "FP • Inspeções": ("FP — Falso Positivo de classificação", "FP = inspeções realmente OK classificadas como NOK."),
    "FN • Inspeções": ("FN — Falso Negativo de classificação", "FN = inspeções realmente NOK classificadas como OK."),
    "Accuracy": ("Accuracy — Acurácia", "Fórmula: (TP + TN) / (TP + TN + FP + FN). Percentual de classificações corretas."),
    "Precision": ("Precision — Precisão", "Fórmula: TP / (TP + FP). Entre os NOK classificados, quantos realmente são NOK."),
    "Recall": ("Recall — Sensibilidade", "Fórmula: TP / (TP + FN). Entre os NOK reais, quantos foram encontrados."),
    "Recall / Sensibilidade": ("Recall — Sensibilidade", "Fórmula: TP / (TP + FN). Entre os NOK reais, quantos foram encontrados."),
    "Specificity": ("Specificity — Especificidade", "Fórmula: TN / (TN + FP). Entre os OK reais, quantos foram mantidos como OK."),
    "FPR": ("FPR — False Positive Rate", "Fórmula: FP / (FP + TN). Taxa de falsos positivos entre os OK reais."),
    "FNR": ("FNR — False Negative Rate", "Fórmula: FN / (FN + TP). Taxa de escapes entre os NOK reais."),
    "F1": ("F1-Score", "Fórmula: 2 × Precision × Recall / (Precision + Recall). Média harmônica entre Precision e Recall."),
    "MCC": ("MCC — Matthews Correlation Coefficient", "Coeficiente de correlação da matriz de confusão. Varia de -1 a +1."),
    "Prevalência": ("Prevalência", "Fórmula: (TP + FN) / N. Proporção de inspeções realmente NOK."),
    "Taxa de alarme NOK": ("Taxa de alarme NOK", "Fórmula: (TP + FP) / N. Proporção de inspeções classificadas como NOK."),
    "FP / 100 inspeções": ("Falsos positivos normalizados", "Fórmula: (FP / N) × 100. Carga de falso alarme normalizada para 100 inspeções."),
    "FN / 100 inspeções": ("Falsos negativos normalizados", "Fórmula: (FN / N) × 100. Carga de escape normalizada para 100 inspeções."),
    "Precision detecção": ("Precision de detecção", "Fórmula: TP-D / (TP-D + FP-D). Mede a confiabilidade das ocorrências detectadas."),
    "Recall detecção": ("Recall de detecção", "Fórmula: TP-D / (TP-D + FN-D). Mede a capacidade de encontrar ocorrências reais."),
}

class Tooltip:
    def __init__(self, widget, title, text):
        self.widget=widget; self.title=title; self.text=text; self.tip=None
        widget.bind("<Enter>", self.show, add="+")
        widget.bind("<Leave>", self.hide, add="+")
        widget.bind("<ButtonPress>", self.hide, add="+")
    def show(self, event=None):
        if self.tip is not None: return
        x=self.widget.winfo_rootx()+18; y=self.widget.winfo_rooty()+self.widget.winfo_height()+6
        self.tip=tk.Toplevel(self.widget); self.tip.wm_overrideredirect(True); self.tip.wm_geometry(f"+{x}+{y}")
        self.tip.configure(bg="#07131C")
        frame=tk.Frame(self.tip,bg="#0E202D",highlightbackground="#2B718F",highlightthickness=1); frame.pack()
        tk.Label(frame,text=self.title,bg="#0E202D",fg=ACCENT2,font=("Segoe UI",9,"bold"),anchor="w",justify="left").pack(fill="x",padx=10,pady=(8,2))
        tk.Label(frame,text=self.text,bg="#0E202D",fg=TEXT,font=("Segoe UI",8),anchor="w",justify="left",wraplength=390).pack(fill="x",padx=10,pady=(0,9))
    def hide(self,event=None):
        if self.tip is not None:
            try:self.tip.destroy()
            except Exception:pass
            self.tip=None

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP} | V{VERSION}")
        self.after(120, lambda: self.status.set("AIA V1.0 • PROJETO DE CALIBRAÇÃO • HISTÓRICO • RECOMENDAÇÃO • RELATÓRIO EXECUTIVO"))
        self.geometry("1520x930")
        self.minsize(1200, 760)
        self.configure(bg=BG)
        self.rows = []
        self.history = []
        self.comparison_archive = []
        self.calibration_history = []
        self.calibration_project = {"id":"", "name":"", "objective":"", "responsible":"", "created_at":""}
        self.asset_dir = application_dir()
        self.logo_img = None
        self.icon_img = None
        self.source_file = ""
        self.current_view = "dashboard"
        self.meta = {k: tk.StringVar() for k in
                     ["cliente","local","projeto","linha","tecnico","data","threshold","gabarito"]}
        self.expectation = {
            "accuracy": tk.StringVar(value=""), "precision": tk.StringVar(value=""),
            "recall": tk.StringVar(value=""), "f1": tk.StringVar(value=""),
            "fp100": tk.StringVar(value=""), "fn100": tk.StringVar(value=""),
            "source": tk.StringVar(value="Acordo / critério de entrega"),
            "responsible": tk.StringVar(value=""), "date": tk.StringVar(value=""),
            "notes": tk.StringVar(value=""),
        }
        self.setup_style()
        try:
            icon_ico = self.asset_dir / "autaza_icon.ico"
            if os.name == "nt" and icon_ico.exists():
                self.iconbitmap(str(icon_ico))
            icon = self.asset_dir / "autaza_icon.png"
            self.icon_img = tk.PhotoImage(file=str(icon))
            self.iconphoto(True, self.icon_img)
        except Exception:
            pass
        self.build_shell()
        self.new_test(False)

    def setup_style(self):
        s = ttk.Style(self)
        try: s.theme_use("clam")
        except Exception: pass
        s.configure(".", font=("Segoe UI",10), background=BG, foreground=TEXT)
        s.configure("TFrame", background=BG)
        s.configure("TLabel", background=BG, foreground=TEXT)
        s.configure("Accent.TButton", background=ACCENT, foreground="white",
                    padding=(14,8), font=("Segoe UI",10,"bold"))
        s.map("Accent.TButton", background=[("active",ACCENT2)])
        s.configure("TButton", background=PANEL2, foreground=TEXT,
                    bordercolor=BORDER, padding=(11,7))
        s.map("TButton", background=[("active","#203247")])
        s.configure("TEntry", fieldbackground="#0d151e", foreground=TEXT, insertcolor=TEXT)
        s.configure("TCombobox", fieldbackground="#F0F3F5", foreground="#17212B")
        s.configure("Treeview", background="#14202B", fieldbackground="#14202B",
                    foreground=TEXT, rowheight=31)
        s.configure("Treeview.Heading", background="#21465D", foreground=TEXT,
                    font=("Segoe UI",9,"bold"))
        s.map("Treeview", background=[("selected","#245D8F")],
              foreground=[("selected","#FFFFFF")])

    def build_shell(self):
        # ==============================================================
        # AUTAZA INDUSTRIAL CONTROL ROOM — V0.7
        # Uma mudança real de arquitetura visual:
        # - dock vertical compacto em vez da sidebar tradicional;
        # - barra superior de estado/execução;
        # - identidade de "sala de controle", não de dashboard genérico;
        # - todas as funções existentes permanecem nas mesmas rotas.
        # ==============================================================

        # ---------- COMMAND HEADER ----------
        header = tk.Frame(self, bg="#061019", height=78)
        header.pack(fill="x", padx=14, pady=(10, 0))
        header.pack_propagate(False)

        # Brand block
        brand_wrap = tk.Frame(header, bg="#061019")
        brand_wrap.pack(side="left", fill="y")

        logo_path = self.asset_dir / "autaza_header.png"
        try:
            self.logo_img = tk.PhotoImage(file=str(logo_path))
            tk.Label(
                brand_wrap, image=self.logo_img, bg="#061019", bd=0
            ).pack(side="left", padx=(4, 14), pady=8)
        except Exception:
            tk.Label(
                brand_wrap, text="AUTAZA", bg="#061019", fg="white",
                font=("Segoe UI", 22, "bold")
            ).pack(side="left", padx=(4, 14))

        title_wrap = tk.Frame(brand_wrap, bg="#061019")
        title_wrap.pack(side="left", fill="y", pady=11)
        tk.Label(
            title_wrap, text="INSPECTION ANALYTICS",
            bg="#061019", fg=TEXT, font=("Segoe UI", 16, "bold")
        ).pack(anchor="w")
        tk.Label(
            title_wrap, text=f"INDUSTRIAL CONTROL ROOM   /   V{VERSION}",
            bg="#061019", fg=ACCENT2, font=("Segoe UI", 8, "bold")
        ).pack(anchor="w", pady=(2, 0))

        # Right command area
        cmd = tk.Frame(header, bg="#061019")
        cmd.pack(side="right", fill="y", pady=10)

        status_chip = tk.Frame(
            cmd, bg="#0B2531",
            highlightbackground="#1B718E", highlightthickness=1
        )
        status_chip.pack(side="left", padx=(0, 8), fill="y")
        tk.Label(
            status_chip, text="●  SISTEMA ATIVO",
            bg="#0B2531", fg=GOOD,
            font=("Segoe UI", 8, "bold")
        ).pack(padx=10, pady=5)
        tk.Label(
            status_chip, text="APLICAÇÃO LOCAL",
            bg="#0B2531", fg=MUTED,
            font=("Segoe UI", 7)
        ).pack(padx=10, pady=(0, 5))

        for label, command in [
            ("ABRIR EXCEL", self.import_excel),
            ("PROJETO", self.edit_calibration_project),
            ("SALVAR", self.save_project),
            ("EXPORTAR", self.export_excel),
            ("RELATÓRIO", self.export_executive_report),
            ("CSV", self.export_csv),
            ("GLOSSÁRIO", self.glossary),
        ]:
            b = tk.Button(
                cmd, text=label, command=command,
                bg="#0D2432", fg=TEXT, activebackground="#15506B",
                activeforeground="white", bd=0, relief="flat",
                highlightbackground="#24536A", highlightthickness=1,
                padx=10, pady=7, font=("Segoe UI", 8, "bold"),
                cursor="hand2"
            )
            b.pack(side="left", padx=2)

        # ---------- MAIN AREA ----------
        self.main = tk.Frame(self, bg=BG)
        self.main.pack(fill="both", expand=True, padx=14, pady=7)

        # ---------- VERTICAL NAV DOCK ----------
        self.sidebar = tk.Frame(
            self.main, bg="#081923", width=92,
            highlightbackground="#1A4054", highlightthickness=1
        )
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # Compact AIA mark
        mark = tk.Frame(self.sidebar, bg="#0A2532", height=76)
        mark.pack(fill="x", padx=7, pady=7)
        mark.pack_propagate(False)
        tk.Label(
            mark, text="A", bg=ACCENT, fg="white",
            font=("Segoe UI", 25, "bold")
        ).pack(pady=(7, 0))
        tk.Label(
            mark, text="AIA", bg="#0A2532", fg=ACCENT2,
            font=("Segoe UI", 7, "bold")
        ).pack()

        tk.Label(
            self.sidebar, text="MODULES", bg="#081923", fg="#5E8799",
            font=("Segoe UI", 7, "bold")
        ).pack(pady=(4, 5))

        self.nav_buttons = {}
        nav = [
            ("dashboard", "⌂", "HOME"),
            ("graficos", "◒", "GRÁFICOS"),
            ("comparacao", "⇄", "COMPARA"),
            ("calibracao", "◉", "CALIBRA"),
            ("operacao", "▣", "COLETA"),
            ("auditoria", "✓", "AUDITORIA"),
            ("tecnica", "◆", "TÉCNICA"),
            ("engenharia", "⌁", "ENG."),
            ("defeitos", "≡", "DEFEITOS"),
            ("expectativa", "🎯", "META"),
        ]

        for key, icon, label in nav:
            b = tk.Button(
                self.sidebar,
                text=f"{icon}\n{label}",
                command=lambda k=key: self.show(k),
                bg="#081923", fg="#B8CBD5",
                activebackground="#124765", activeforeground="white",
                bd=0, relief="flat", highlightthickness=0,
                anchor="center", justify="center",
                pady=6, font=("Segoe UI", 8, "bold"),
                cursor="hand2"
            )
            b.pack(fill="x", padx=6, pady=2)
            self.nav_buttons[key] = b

        # Compact status area instead of the old large sidebar block
        tk.Frame(self.sidebar, bg="#1B4053", height=1).pack(
            fill="x", padx=12, pady=7
        )
        tk.Label(
            self.sidebar, text="STATUS", bg="#081923", fg="#5E8799",
            font=("Segoe UI", 7, "bold")
        ).pack(pady=(0, 4))
        self.side_status = tk.Label(
            self.sidebar, bg="#081923", fg=MUTED, justify="center",
            anchor="center", wraplength=76, font=("Segoe UI", 7)
        )
        self.side_status.pack(fill="x", padx=6)

        tk.Label(
            self.sidebar, text=f"V{VERSION}", bg="#081923", fg="#3E6D80",
            font=("Segoe UI", 7, "bold")
        ).pack(side="bottom", pady=10)

        # ---------- CONTENT ----------
        # Fixed instrument strip + scrollable work area.
        # This prevents lower cards/text from being clipped on smaller screens.
        self.content_host = tk.Frame(self.main, bg=BG)
        self.content_host.pack(side="left", fill="both", expand=True, padx=(9, 0))

        self.instrument_bar = tk.Frame(
            self.content_host, bg="#091C27", height=34,
            highlightbackground="#173D50", highlightthickness=1
        )
        self.instrument_bar.pack(fill="x", pady=(0, 7))
        self.instrument_bar.pack_propagate(False)

        self.instrument_left = tk.Label(
            self.instrument_bar,
            text="AIA / CONTROL ROOM",
            bg="#091C27", fg=ACCENT2,
            font=("Segoe UI", 8, "bold")
        )
        self.instrument_left.pack(side="left", padx=12)

        self.instrument_right = tk.Label(
            self.instrument_bar,
            text="DATA STREAM: MANUAL     •     ENGINE: READY",
            bg="#091C27", fg="#6E8E9E",
            font=("Segoe UI", 7, "bold")
        )
        self.instrument_right.pack(side="right", padx=12)

        scroll_wrap = tk.Frame(self.content_host, bg=BG)
        scroll_wrap.pack(fill="both", expand=True)

        self.content_canvas = tk.Canvas(
            scroll_wrap, bg=BG, bd=0, highlightthickness=0
        )
        self.content_canvas.pack(side="left", fill="both", expand=True)

        self.content_scrollbar = tk.Scrollbar(
            scroll_wrap, orient="vertical",
            command=self.content_canvas.yview,
            bg="#0B2634", troughcolor="#061019",
            activebackground="#1689B5", width=12
        )
        self.content_scrollbar.pack(side="right", fill="y")

        self.content = tk.Frame(self.content_canvas, bg=BG)
        self.content_window = self.content_canvas.create_window(
            (0, 0), window=self.content, anchor="nw"
        )
        self.content.configure(
            highlightthickness=0, bd=0
        )

        def _sync_scroll_region(event=None):
            self.content_canvas.configure(
                scrollregion=self.content_canvas.bbox("all")
            )

        def _fit_content_width(event):
            self.content_canvas.itemconfigure(
                self.content_window, width=event.width
            )

        self.content.bind("<Configure>", _sync_scroll_region)
        self.content_canvas.bind("<Configure>", _fit_content_width)

        self.content_canvas.bind_all(
            "<MouseWheel>",
            lambda e: self.content_canvas.yview_scroll(
                int(-1 * (e.delta / 120)), "units"
            )
        )
        self.content_canvas.configure(
            yscrollcommand=self.content_scrollbar.set
        )

        # Footer status
        bar = tk.Frame(self, bg="#040C12", height=25)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        self.status = tk.StringVar(value="Pronto.")
        tk.Label(
            bar, textvariable=self.status, bg="#040C12", fg=MUTED,
            font=("Segoe UI", 8)
        ).pack(side="left", padx=12)
        tk.Label(
            bar, text=f"AIA  /  CONTROL ROOM  /  V{VERSION}",
            bg="#040C12", fg="#416777", font=("Segoe UI", 8, "bold")
        ).pack(side="right", padx=12)

        self.show("dashboard")

    def clear_content(self):
        for w in self.content.winfo_children():
            if w is not getattr(self, "instrument_bar", None):
                w.destroy()

    def set_instrument(self, module_name):
        if hasattr(self, "instrument_left"):
            self.instrument_left.configure(text=f"AIA  /  {module_name.upper()}")
        if hasattr(self, "instrument_right"):
            self.instrument_right.configure(
                text=f"DATA STREAM: {'EXCEL' if self.source_file else 'MANUAL'}"
                      f"     •     ENGINE: {'ACTIVE' if self.rows else 'READY'}"
            )

    def show(self,key):
        self.current_view=key
        self.clear_content()
        module_names = {
            "dashboard": "EXECUTIVE / HOME",
            "graficos": "PERFORMANCE / GRÁFICOS",
            "comparacao": "COMPARAÇÃO / EVOLUÇÃO",
            "calibracao": "PROJETO / CALIBRAÇÃO",
            "operacao": "COLETA / OPERAÇÃO",
            "auditoria": "AUDITORIA / HISTÓRICO",
            "tecnica": "ANÁLISE TÉCNICA / DIAGNÓSTICO",
            "engenharia": "ENGENHARIA / DECISÃO",
            "defeitos": "ANÁLISE POR DEFEITO",
            "expectativa": "EXPECTATIVA / CRITÉRIO DE ENTREGA",
        }
        self.set_instrument(module_names.get(key, key))
        for k,b in self.nav_buttons.items():
            b.configure(bg="#0D5676" if k==key else "#081923", fg="white" if k==key else "#B8CBD5")
        fn={"dashboard":self.view_dashboard,"graficos":self.view_graficos,
            "comparacao":self.view_comparacao,
            "calibracao":self.view_calibracao,
            "operacao":self.view_operacao,"auditoria":self.view_auditoria,
            "tecnica":self.view_tecnica,"engenharia":self.view_engenharia,
            "defeitos":self.view_defeitos,"expectativa":self.view_expectativa}[key]
        fn()

    def card(self,parent,title,subtitle=None):
        f=tk.Frame(parent,bg=PANEL,highlightbackground=BORDER,highlightthickness=1)
        tk.Label(f,text=title,bg=PANEL,fg=TEXT,font=("Segoe UI",11,"bold")).pack(
            anchor="w",padx=14,pady=(12,3))
        if subtitle:
            tk.Label(f,text=subtitle,bg=PANEL,fg=MUTED,font=("Segoe UI",8),
                     wraplength=1000,justify="left").pack(anchor="w",padx=14,pady=(0,8))
        return f

    # ---------- DATA / MOTOR ----------
    def classify(self,r):
        r=normalize_row(r)
        r.update(tp=0,vn=0,fp=0,fn=0,calc_status="N/D")
        ref=r["ref"]; vision=r["vision"]
        if ref=="Defeito confirmado":
            r["tp"]=1 if vision=="NOK" else 0
            r["fn"]=1 if vision!="NOK" else 0
            r["calc_status"]="Classificação calculada"
        elif ref=="OK confirmado":
            r["vn"]=1 if vision=="OK" else 0
            r["fp"]=1 if vision!="OK" else 0
            r["calc_status"]="Classificação calculada"
        return r

    def detection_for_row(self, r0):
        """Calcula erros no nível de ocorrência, separado da classificação OK/NOK."""
        r=normalize_row(r0)
        if not r.get("detection_available", True):
            return {"real":0,"det":0,"tp":0,"fp":0,"fn":0,"available":False}
        ds=r.get("defects",[])
        if ds:
            real=det=tp=fp=fn=0
            for d in ds:
                rr,dd=int(d.get("real",0)),int(d.get("det",0))
                real += rr; det += dd
                tp += min(rr,dd); fp += max(dd-rr,0); fn += max(rr-dd,0)
            return {"real":real,"det":det,"tp":tp,"fp":fp,"fn":fn,"available":True}
        rr,dd=int(r.get("real",0)),int(r.get("det",0))
        return {"real":rr,"det":dd,"tp":min(rr,dd),"fp":max(dd-rr,0),"fn":max(rr-dd,0),"available":True}

    def detection_totals(self):
        total={"real":0,"det":0,"tp":0,"fp":0,"fn":0,"available_rows":0,"unavailable_rows":0}
        for r in self.rows:
            d=self.detection_for_row(r)
            if not d["available"]:
                total["unavailable_rows"] += 1
                continue
            total["available_rows"] += 1
            for k in ("real","det","tp","fp","fn"):
                total[k] += d[k]
        total["precision"]=safe(total["tp"], total["tp"]+total["fp"])
        total["recall"]=safe(total["tp"], total["tp"]+total["fn"])
        return total

    def defect_counts(self):
        real=det=tp=fp=fn=0
        detailed=False
        by=defaultdict(lambda:{"real":0,"det":0,"tp":0,"fp":0,"fn":0,"inspecoes":0})
        for r0 in self.rows:
            r=normalize_row(r0)
            ds=r.get("defects",[])
            if ds:
                detailed=True
                for d in ds:
                    rr,dd=d["real"],d["det"]
                    tt=min(rr,dd); ff=max(dd-rr,0); nn=max(rr-dd,0)
                    real+=rr; det+=dd; tp+=tt; fp+=ff; fn+=nn
                    a=by[d["type"]]
                    a["real"]+=rr; a["det"]+=dd; a["tp"]+=tt; a["fp"]+=ff; a["fn"]+=nn; a["inspecoes"]+=1
            else:
                d=self.detection_for_row(r)
                if not d["available"]:
                    continue
                real+=d["real"]; det+=d["det"]; tp+=d["tp"]; fp+=d["fp"]; fn+=d["fn"]
        return {"real":real,"det":det,"tp":tp,"fp":fp,"fn":fn,
                "by":by,"detailed":detailed}

    def classification_totals(self):
        # O motor novo só contabiliza inspeções com referência explícita.
        # Linhas importadas do Excel permanecem como HISTÓRICO e não entram
        # automaticamente na matriz de confusão do motor novo.
        vals={"tp":0,"vn":0,"fp":0,"fn":0}
        for i,r0 in enumerate(self.rows):
            r=normalize_row(r0)
            if r.get("calc_status") == "Histórico" or str(r.get("ref","")).startswith("Importado"):
                self.rows[i]=r
                continue
            r=self.classify(r)
            self.rows[i]=r
            for k in vals:
                vals[k]+=r[k]
        return vals

    def metrics(self):
        t=self.classification_totals()
        return classification_metrics(t["tp"],t["vn"],t["fp"],t["fn"])

    def historical_metrics(self):
        """Métricas calculadas diretamente dos totais importados da planilha."""
        if not self.history:
            return classification_metrics(0,0,0,0)
        h=self.history[-1]
        return classification_metrics(h.get("TP",0), h.get("VN/TN",0),
                                      h.get("FP",0), h.get("FN",0))

    def display_metrics(self):
        """Dashboard/gráficos usam o histórico importado quando ele é a fonte ativa."""
        if self.source_file and self.history:
            return self.historical_metrics()
        return self.metrics()

    # ---------- EXPECTATIVA / META ----------
    def _parse_target(self,key):
        raw=(self.expectation[key].get() or "").strip().replace(",",".")
        if not raw:return None
        try:
            v=float(raw)
            if key in ("accuracy","precision","recall","f1"):
                if v>1:v/=100.0
                return max(0.0,min(1.0,v))
            return max(0.0,v)
        except Exception:return None

    def expectation_snapshot(self):
        return {k:self._parse_target(k) for k in ("accuracy","precision","recall","f1","fp100","fn100")} | {
            "source":self.expectation["source"].get().strip(),
            "responsible":self.expectation["responsible"].get().strip(),
            "date":self.expectation["date"].get().strip(),
            "notes":self.expectation["notes"].get().strip(),
        }

    def _expectation_actuals(self):
        m=self.display_metrics(); n=m.get("N",0)
        return {"accuracy":m.get("Accuracy"),"precision":m.get("Precision"),"recall":m.get("Recall / Sensibilidade"),
                "f1":m.get("F1"),"fp100":None if not n else m.get("FP",0)/n*100.0,
                "fn100":None if not n else m.get("FN",0)/n*100.0}

    def _expectation_rows(self):
        actual=self._expectation_actuals(); target=self.expectation_snapshot()
        specs=[("Accuracy","accuracy",">="),("Precision","precision",">="),("Recall","recall",">="),
               ("F1","f1",">="),("FP / 100 inspeções","fp100","<="),("FN / 100 inspeções","fn100","<=")]
        rows=[]
        for label,key,op in specs:
            t=target[key]; a=actual[key]
            if t is None: status="SEM META"
            elif a is None: status="N/D"
            elif (a>=t if op==">=" else a<=t): status="ATENDIDA"
            else: status="ABAIXO DA META" if op==">=" else "ACIMA DA META"
            rows.append((label,key,op,a,t,status))
        return rows

    def expectation_summary(self):
        rows=self._expectation_rows(); configured=[r for r in rows if r[4] is not None]
        if not configured:return {"configured":0,"met":0,"status":"SEM META"}
        met=sum(1 for r in configured if r[5]=="ATENDIDA")
        status="META ATINGIDA" if met==len(configured) else ("FORA DA EXPECTATIVA" if met==0 else "EM CALIBRAÇÃO")
        return {"configured":len(configured),"met":met,"status":status}

    def _expectation_status_color(self,status):
        return {"META ATINGIDA":GOOD,"ATENDIDA":GOOD,"EM CALIBRAÇÃO":WARN,"FORA DA EXPECTATIVA":BAD,
                "ABAIXO DA META":WARN,"ACIMA DA META":WARN,"SEM META":MUTED,"N/D":MUTED}.get(status,TEXT)

    def view_expectativa(self):
        self.section_title("EXPECTATIVA  /  CRITÉRIO DE ENTREGA",
            "Defina o norte da calibração e acompanhe resultado × meta. Campo vazio = indicador não monitorado.")
        summary=self.expectation_summary()
        hero=self.card(self.content,"🎯 MODO EXPECTATIVA","Critérios de aceitação definidos para orientar a calibração e a entrega.")
        hero.pack(fill="x",pady=(0,7))
        row=tk.Frame(hero,bg=PANEL);row.pack(fill="x",padx=16,pady=(2,12))
        tk.Label(row,text=summary["status"],bg=PANEL,fg=self._expectation_status_color(summary["status"]),font=("Segoe UI",16,"bold")).pack(side="left")
        tk.Label(row,text=f"   {summary['met']}/{summary['configured']} metas configuradas atendidas",bg=PANEL,fg=MUTED,font=("Segoe UI",9)).pack(side="left")

        meta=self.card(self.content,"METAS DE DESEMPENHO","Accuracy, Precision, Recall e F1 = mínimo. FP/100 e FN/100 = máximo.")
        meta.pack(fill="x",pady=5)
        grid=tk.Frame(meta,bg=PANEL);grid.pack(fill="x",padx=16,pady=8)
        for i,(lab,key) in enumerate([("Accuracy (%)","accuracy"),("Precision (%)","precision"),("Recall (%)","recall"),("F1 (%)","f1"),("FP / 100","fp100"),("FN / 100","fn100")]):
            r,c=divmod(i,3);box=tk.Frame(grid,bg=PANEL);box.grid(row=r,column=c,sticky="ew",padx=6,pady=6);grid.columnconfigure(c,weight=1)
            tk.Label(box,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
            ttk.Entry(box,textvariable=self.expectation[key]).pack(fill="x")
        table=ttk.Treeview(meta,columns=("ind","meta","atual","regra","status"),show="headings",height=6)
        for c,h,w in [("ind","Indicador",210),("meta","Meta",100),("atual","Resultado atual",120),("regra","Critério",90),("status","Status",180)]:table.heading(c,text=h);table.column(c,width=w,anchor="center" if c!="ind" else "w")
        table.pack(fill="x",padx=16,pady=(2,14))
        for label,key,op,a,t,status in self._expectation_rows():
            ms="N/D" if t is None else (f"{t*100:.2f}%" if key in ("accuracy","precision","recall","f1") else f"{t:.2f}")
            av="N/D" if a is None else (pct(a) if key in ("accuracy","precision","recall","f1") else f"{a:.2f}")
            iid=table.insert("","end",values=(label,ms,av,op,status))
            for item in table.get_children():
                if item==iid and status not in ("SEM META","N/D"):
                    # no per-row colors in ttk; tooltip carries the technical explanation
                    break

        ctx=self.card(self.content,"RASTREABILIDADE DA META","Registre de onde veio o critério para que o número continue defensável depois da calibração.")
        ctx.pack(fill="x",pady=5)
        form=tk.Frame(ctx,bg=PANEL);form.pack(fill="x",padx=16,pady=8)
        for i,(lab,key) in enumerate([("Origem do critério","source"),("Responsável","responsible"),("Data do critério","date")]):
            box=tk.Frame(form,bg=PANEL);box.grid(row=0,column=i,sticky="ew",padx=6);form.columnconfigure(i,weight=1)
            tk.Label(box,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w");ttk.Entry(box,textvariable=self.expectation[key]).pack(fill="x")
        box=tk.Frame(form,bg=PANEL);box.grid(row=1,column=0,columnspan=3,sticky="ew",padx=6,pady=6)
        tk.Label(box,text="Observações / acordo com o cliente",bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w");ttk.Entry(box,textvariable=self.expectation["notes"]).pack(fill="x")

        read=self.card(self.content,"LEITURA DE CALIBRAÇÃO","Use a meta como norte; ela não substitui a análise de risco nem define sozinha a condição de entrega.")
        read.pack(fill="both",expand=True,pady=5)
        for line in ["• META ATINGIDA: todos os critérios configurados estão dentro das metas.","• EM CALIBRAÇÃO: parte das metas foi atendida; continue ajustando e comparando.","• FORA DA EXPECTATIVA: nenhum critério configurado foi atendido na amostra atual.","• COMPARA continua mostrando o efeito de diferentes thresholds; EXPECTATIVA acrescenta o critério de aceitação."]:
            tk.Label(read,text=line,bg=PANEL,fg=TEXT,wraplength=1100,justify="left",font=("Segoe UI",9)).pack(anchor="w",padx=16,pady=5)

    # ---------- VIEWS ----------
    def metric_card(self,parent,label,value,accent=None):
        f=tk.Frame(parent,bg=PANEL,highlightbackground=BORDER,highlightthickness=1,cursor="hand2")
        title=tk.Label(f,text=label,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold"),cursor="hand2")
        title.pack(anchor="w",padx=12,pady=(10,1))
        val=tk.Label(f,text=value,bg=PANEL,fg=accent or TEXT,font=("Segoe UI",18,"bold"),cursor="hand2")
        val.pack(anchor="w",padx=12,pady=(0,10))
        tip=TOOLTIP_DEFS.get(label)
        if tip:
            Tooltip(f,*tip); Tooltip(title,*tip); Tooltip(val,*tip)
        return f

    def section_title(self, title, subtitle=""):
        tk.Label(self.content, text=title, bg=BG, fg=TEXT,
                 font=("Segoe UI",18,"bold")).pack(anchor="w")
        if subtitle:
            tk.Label(self.content, text=subtitle, bg=BG, fg=MUTED,
                     font=("Segoe UI",9)).pack(anchor="w", pady=(2,10))

    def chart_card(self, parent, title, subtitle=""):
        f = tk.Frame(parent, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        tk.Label(f, text=title, bg=PANEL, fg=TEXT,
                 font=("Segoe UI",10,"bold")).pack(anchor="w", padx=14, pady=(10,2))
        if subtitle:
            tk.Label(f, text=subtitle, bg=PANEL, fg=MUTED,
                     font=("Segoe UI",8)).pack(anchor="w", padx=14, pady=(0,4))
        return f

    def draw_bar_chart(self, parent, items, max_value=None, suffix=""):
        """Gráfico de barras simples, sem dependências externas."""
        canvas = tk.Canvas(parent, bg=PANEL, highlightthickness=0, height=220)
        canvas.pack(fill="both", expand=True, padx=12, pady=(4,12))
        canvas.update_idletasks()

        w = max(canvas.winfo_width(), 520)
        h = 220
        left, right, top, bottom = 72, 18, 18, 38
        plot_w, plot_h = w-left-right, h-top-bottom
        vals = [float(v or 0) for _,v,_ in items]
        maxv = max_value if max_value is not None else (max(vals) if vals else 1)
        if maxv <= 0: maxv = 1

        # grade horizontal
        for i in range(5):
            y = top + plot_h - (plot_h*i/4)
            canvas.create_line(left, y, w-right, y, fill="#29404F")
            val = maxv*i/4
            canvas.create_text(left-8, y, text=f"{val:.0f}{suffix}",
                               fill=MUTED, anchor="e", font=("Segoe UI",7))

        n = max(len(items),1)
        slot = plot_w/n
        bw = min(58, slot*0.58)
        for i,(label,val,accent) in enumerate(items):
            x = left + slot*i + slot/2
            v = float(val or 0)
            bh = plot_h*(v/maxv)
            y1 = top+plot_h-bh
            canvas.create_rectangle(x-bw/2, y1, x+bw/2, top+plot_h,
                                    fill=accent, outline="")
            canvas.create_text(x, y1-8, text=f"{v:.0f}{suffix}",
                               fill=TEXT, font=("Segoe UI",8,"bold"))
            canvas.create_text(x, top+plot_h+12, text=label,
                               fill=MUTED, font=("Segoe UI",7))

        return canvas

    def draw_percent_chart(self, parent, items):
        canvas = tk.Canvas(parent, bg=PANEL, highlightthickness=0, height=220)
        canvas.pack(fill="both", expand=True, padx=12, pady=(4,12))
        canvas.update_idletasks()
        w, h = max(canvas.winfo_width(), 520), 220
        left, right, top, bottom = 120, 48, 18, 30
        plot_w = w-left-right
        for i in range(0, 101, 25):
            x = left + plot_w*i/100
            canvas.create_line(x, top, x, h-bottom, fill="#29404F")
            canvas.create_text(x, h-bottom+12, text=f"{i}%",
                               fill=MUTED, font=("Segoe UI",7))
        row_h = (h-top-bottom)/max(len(items),1)
        for i,(label,val,accent) in enumerate(items):
            y = top + row_h*i + row_h/2
            p = 0 if val is None else max(0,min(1,float(val)))
            canvas.create_text(left-8,y,text=label,fill=MUTED,anchor="e",
                               font=("Segoe UI",8))
            canvas.create_rectangle(left,y-7,left+plot_w,y+7,
                                    fill="#1B2B38",outline="")
            canvas.create_rectangle(left,y-7,left+plot_w*p,y+7,
                                    fill=accent,outline="")
            canvas.create_text(left+plot_w*p+8,y,text="N/D" if val is None else f"{p*100:.1f}%",
                               fill=TEXT,anchor="w",font=("Segoe UI",8,"bold"))
        return canvas

    def view_dashboard(self):
        fonte = "HISTÓRICO DA PLANILHA" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
        self.section_title(
            "DASHBOARD  /  VISÃO EXECUTIVA",
            f"Autaza Inspection Analytics • {fonte} • desempenho consolidado da inspeção e detecção")

        m=self.display_metrics(); d=self.defect_counts()
        cards=[("TP • Inspeções",m["TP"],GOOD),("TN/VN • Inspeções",m["VN/TN"],GOOD),
               ("FP • Inspeções",m["FP"],BAD),("FN • Inspeções",m["FN"],WARN),
               ("Accuracy",pct(m["Accuracy"]),ACCENT2),("Precision",pct(m["Precision"]),ACCENT2),
               ("Recall",pct(m["Recall / Sensibilidade"]),ACCENT2),("F1",pct(m["F1"]),ACCENT2)]
        grid=tk.Frame(self.content,bg=BG);grid.pack(fill="x")
        for i,(lab,val,col) in enumerate(cards):
            r,c=divmod(i,4)
            self.metric_card(grid,lab,str(val),col).grid(row=r,column=c,sticky="ew",padx=4,pady=4)
        for c in range(4):grid.columnconfigure(c,weight=1)

        es=self.expectation_summary()
        if es["configured"]:
            ec=self.card(self.content,"🎯 EXPECTATIVA DE ENTREGA",f"{es['met']}/{es['configured']} critérios atendidos")
            ec.pack(fill="x",pady=(7,3))
            er=tk.Frame(ec,bg=PANEL);er.pack(fill="x",padx=16,pady=(0,10))
            tk.Label(er,text=es["status"],bg=PANEL,fg=self._expectation_status_color(es["status"]),font=("Segoe UI",12,"bold")).pack(side="left")
            tk.Button(er,text="ABRIR EXPECTATIVA",command=lambda:self.show("expectativa"),bg="#0E4560",fg=TEXT,activebackground="#176B8F",bd=0,padx=12,pady=6,font=("Segoe UI",8,"bold")).pack(side="right")

        quick=tk.Frame(self.content,bg=BG)
        quick.pack(fill="x", pady=(8,4))
        tk.Button(quick,text="◒  ABRIR GRÁFICOS",command=lambda:self.show("graficos"),
                  bg="#0E4560",fg=TEXT,activebackground="#176B8F",
                  activeforeground="white",bd=0,padx=14,pady=8,
                  font=("Segoe UI",9,"bold")).pack(side="right")

        lower=tk.Frame(self.content,bg=BG);lower.pack(fill="both",expand=True,pady=6)
        left=self.card(lower,"CLASSIFICAÇÃO DAS INSPEÇÕES",
                        "Uma linha representa uma carroceria/peça. As métricas estatísticas usam esta unidade.")
        left.pack(side="left",fill="both",expand=True,padx=(0,4))
        txt=(f"TP = {m['TP']}\nTN/VN = {m['VN/TN']}\nFP = {m['FP']}\nFN = {m['FN']}\n\n"
             f"N de inspeções classificadas = {m['N']}")
        tk.Label(left,text=txt,bg=PANEL,fg=TEXT,font=("Consolas",11),
                 justify="left").pack(anchor="w",padx=16,pady=12)

        right=self.card(lower,"DETECÇÃO DE OCORRÊNCIAS",
                         "Real × detectado. O tipo do defeito é opcional.")
        right.pack(side="left",fill="both",expand=True,padx=(4,0))
        txt2=(f"Defeitos reais registrados = {d['real']}\n"
              f"Detecções registradas = {d['det']}\n"
              f"TP de detecção (por contagem) = {d['tp']}\n"
              f"FP de detecção (extras) = {d['fp']}\n"
              f"FN de detecção (não detectados) = {d['fn']}\n"
              f"Precision de detecção = {pct(safe(d['tp'], d['tp'] + d['fp']))}\n"
              f"Recall de detecção = {pct(safe(d['tp'], d['tp'] + d['fn']))}\n\n"
              + ("Tipos detalhados informados." if d["detailed"] else
                 "Coleta sem classificação de tipo."))
        tk.Label(right,text=txt2,bg=PANEL,fg=TEXT,font=("Consolas",11),
                 justify="left").pack(anchor="w",padx=16,pady=12)

    def view_graficos(self):
        fonte = "HISTÓRICO DA PLANILHA" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
        self.section_title(
            "GRÁFICOS  /  PERFORMANCE",
            f"Autaza Inspection Analytics • {fonte} • leitura visual para acompanhamento técnico e executivo")

        m=self.display_metrics(); d=self.defect_counts()
        top=tk.Frame(self.content,bg=BG)
        top.pack(fill="both",expand=True)

        c1=self.chart_card(top,"MATRIZ DE CLASSIFICAÇÃO","Quantidade de inspeções por resultado")
        c1.grid(row=0,column=0,sticky="nsew",padx=(0,4),pady=(0,4))
        self.draw_bar_chart(c1,[("TP",m["TP"],GOOD),("TN",m["VN/TN"],GOOD),
                                ("FP",m["FP"],BAD),("FN",m["FN"],WARN)],
                            max_value=max([m["TP"],m["VN/TN"],m["FP"],m["FN"],1]))

        c2=self.chart_card(top,"INDICADORES DE PERFORMANCE","Percentual das principais métricas")
        c2.grid(row=0,column=1,sticky="nsew",padx=(4,0),pady=(0,4))
        self.draw_percent_chart(c2,[
            ("Accuracy",m["Accuracy"],ACCENT2),
            ("Precision",m["Precision"],ACCENT),
            ("Recall",m["Recall / Sensibilidade"],GOOD),
            ("F1",m["F1"],WARN),
        ])

        c3=self.chart_card(top,"OCORRÊNCIAS","Defeitos reais × detecções × divergências")
        c3.grid(row=1,column=0,sticky="nsew",padx=(0,4),pady=(4,0))
        self.draw_bar_chart(c3,[("Real",d["real"],GOOD),("Detectado",d["det"],ACCENT),
                                ("TP Det.",d["tp"],GOOD),("FP Det.",d["fp"],BAD),("FN Det.",d["fn"],WARN)],
                            max_value=max([d["real"],d["det"],d["tp"],d["fp"],d["fn"],1]))

        c4=self.chart_card(top,"LEITURA EXECUTIVA","Resumo visual para apresentação")
        c4.grid(row=1,column=1,sticky="nsew",padx=(4,0),pady=(4,0))
        summary = [
            ("Inspeções", m["N"]),
            ("Recall", pct(m["Recall / Sensibilidade"])),
            ("Precision", pct(m["Precision"])),
            ("Falsos positivos", m["FP"]),
            ("Falsos negativos", m["FN"]),
        ]
        for i,(lab,val) in enumerate(summary):
            row=tk.Frame(c4,bg=PANEL)
            row.pack(fill="x",padx=16,pady=6)
            tk.Label(row,text=lab,bg=PANEL,fg=MUTED,
                     font=("Segoe UI",9)).pack(side="left")
            tk.Label(row,text=str(val),bg=PANEL,fg=TEXT,
                     font=("Segoe UI",10,"bold")).pack(side="right")
        tk.Label(c4,text="Use esta tela para apresentar tendência e desempenho sem abrir a planilha.",
                 bg=PANEL,fg=ACCENT2,wraplength=460,justify="left",
                 font=("Segoe UI",8)).pack(anchor="w",padx=16,pady=(10,16))

        top.columnconfigure(0,weight=1); top.columnconfigure(1,weight=1)
        top.rowconfigure(0,weight=1); top.rowconfigure(1,weight=1)

    def view_operacao(self):
        tk.Label(self.content,text="COLETA / OPERAÇÃO",bg=BG,fg=TEXT,
                 font=("Segoe UI",18,"bold")).pack(anchor="w")
        tk.Label(self.content,text="Coleta rápida para campo + detalhamento opcional por inspeção",
                 bg=BG,fg=MUTED).pack(anchor="w",pady=(2,8))
        info=self.card(self.content,"IDENTIFICAÇÃO DO TESTE")
        info.pack(fill="x",pady=4)
        fields=[("Cliente","cliente"),("Local / Planta","local"),("Projeto","projeto"),
                ("Linha / Estação","linha"),("Técnico","tecnico"),("Data","data"),
                ("Threshold","threshold"),("Método de referência","gabarito")]
        body=tk.Frame(info,bg=PANEL);body.pack(fill="x",padx=6,pady=(0,10))
        for i,(lab,key) in enumerate(fields):
            r,c=divmod(i,4);box=tk.Frame(body,bg=PANEL)
            box.grid(row=r,column=c,sticky="ew",padx=8,pady=4)
            tk.Label(box,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
            if key=="gabarito":
                w=ttk.Combobox(box,textvariable=self.meta[key],values=GABARITOS,state="readonly")
            else:
                w=ttk.Entry(box,textvariable=self.meta[key])
            w.pack(fill="x");body.columnconfigure(c,weight=1)

        source_label = "✓ Metadados preenchidos automaticamente pelo Excel" if self.source_file else "Entrada manual — preencha o contexto da inspeção"
        tk.Label(info,text=source_label,bg=PANEL,fg=GOOD if self.source_file else MUTED,
                 font=("Segoe UI",8,"bold")).pack(anchor="w",padx=14,pady=(0,8))

        tb=tk.Frame(self.content,bg=BG);tb.pack(fill="x",pady=5)
        ttk.Button(tb,text="+ Nova inspeção",style="Accent.TButton",
                   command=self.add_row).pack(side="left",padx=3)
        ttk.Button(tb,text="↻ Nova coleta",command=self.new_collection).pack(side="left",padx=3)
        ttk.Button(tb,text="⚡ Coleta rápida",command=self.quick_entry).pack(side="left",padx=3)
        ttk.Button(tb,text="Detalhar defeitos",command=self.edit_row).pack(side="left",padx=3)
        ttk.Button(tb,text="Excluir",command=self.delete_row).pack(side="left",padx=3)
        ttk.Button(tb,text="Recalcular",command=self.refresh_all).pack(side="left",padx=3)

        legend=tk.Label(self.content, text="TP-C / VN-C / FP-C / FN-C = classificação da inspeção (OK/NOK)  •  TP-D / FP-D / FN-D = detecção de ocorrências por contagem", bg=BG, fg=ACCENT2, font=("Segoe UI",8,"bold"))
        legend.pack(anchor="w", pady=(0,3))

        table=self.card(self.content,"INSPEÇÕES")
        table.pack(fill="both",expand=True,pady=5)
        cols=("n","id","vision","ref","real","det","tp","vn","fp","fn","dtp","dfp","dfn","def","status")
        heads=("Nº","Identificação","Vision","Referência","Real","Detectado","TP-C","VN-C","FP-C","FN-C","TP-D","FP-D","FN-D","Detalhes","Motor")
        self.tree=ttk.Treeview(table,columns=cols,show="headings")
        widths={"n":45,"id":90,"vision":65,"ref":115,"real":55,"det":65,
                "tp":50,"vn":50,"fp":50,"fn":50,"dtp":50,"dfp":50,"dfn":50,"def":75,"status":125}
        for c,h in zip(cols,heads):
            self.tree.heading(c,text=h);self.tree.column(c,width=widths.get(c,80),anchor="center")
        self.tree.pack(fill="both",expand=True,padx=10,pady=10)
        self.tree.bind("<Double-1>",lambda e:self.edit_row())
        self.refresh_table()

    def refresh_table(self):
        if not hasattr(self,"tree"): return
        for x in self.tree.get_children(): self.tree.delete(x)
        for i,r0 in enumerate(self.rows,1):
            r=normalize_row(r0); self.rows[i-1]=r
            detailed=f"{len(r['defects'])} tipo(s)" if r["defects"] else "Opcional"
            dd=self.detection_for_row(r)
            detvals=(dd["tp"],dd["fp"],dd["fn"]) if dd["available"] else ("N/D","N/D","N/D")
            self.tree.insert("", "end", iid=str(i-1),
                values=(i,r["id"],r["vision"],r["ref"],r["real"],r["det"],
                        r["tp"],r["vn"],r["fp"],r["fn"],detvals[0],detvals[1],detvals[2],detailed,r["calc_status"]))

    def add_row(self):
        self.rows.append(blank_row())
        self.refresh_all()
        self.edit_row(len(self.rows)-1)

    def quick_entry(self):
        win=tk.Toplevel(self);win.title("Coleta rápida — nova inspeção")
        win.geometry("610x570");win.configure(bg=BG);win.transient(self);win.grab_set()
        frame=tk.Frame(win,bg=PANEL,highlightbackground=BORDER,highlightthickness=1)
        frame.pack(fill="both",expand=True,padx=15,pady=15)
        tk.Label(frame,text="COLETA RÁPIDA",bg=PANEL,fg=TEXT,
                 font=("Segoe UI",16,"bold")).pack(anchor="w",padx=18,pady=(18,3))
        tk.Label(frame,text="Não é necessário informar o tipo do defeito. Registre somente o que foi observado.",
                 bg=PANEL,fg=MUTED,wraplength=520,justify="left").pack(anchor="w",padx=18,pady=(0,14))
        vars={}
        specs=[("Identificação","id"),("Vision AI","vision"),("Referência / gabarito","ref"),
               ("Quantidade real de defeitos","real"),("Quantidade detectada","det"),
               ("Observação","obs")]
        for lab,key in specs:
            row=tk.Frame(frame,bg=PANEL);row.pack(fill="x",padx=18,pady=5)
            tk.Label(row,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
            default={"vision":"OK","ref":"Não determinado","real":"0","det":"0"}.get(key,"")
            v=tk.StringVar(value=default);vars[key]=v
            if key=="vision": w=ttk.Combobox(row,textvariable=v,values=["OK","NOK"],state="readonly")
            elif key=="ref": w=ttk.Combobox(row,textvariable=v,values=REFS,state="readonly")
            else: w=ttk.Entry(row,textvariable=v)
            w.pack(fill="x")
        btns=tk.Frame(frame,bg=PANEL);btns.pack(fill="x",padx=18,pady=18)
        def collect(next_one=False):
            try:
                real=max(0,int(float(vars["real"].get() or 0)))
                det=max(0,int(float(vars["det"].get() or 0)))
            except ValueError:
                messagebox.showerror("Dados","As quantidades devem ser numéricas.",parent=win);return
            r=blank_row()
            r.update({"id":vars["id"].get().strip(),"vision":vars["vision"].get() or "OK",
                      "ref":vars["ref"].get() or "Não determinado","real":real,"det":det,
                      "obs":vars["obs"].get().strip(),"source":"Manual — Coleta rápida"})
            self.rows.append(self.classify(r))
            self.source_file=""
            self.refresh_all()
            self.status.set(f"Inspeção {len(self.rows)} registrada.")
            if next_one:
                for k,v in vars.items():
                    if k not in ("vision","ref"): v.set("")
                vars["vision"].set("OK");vars["ref"].set("Não determinado")
                vars["real"].set("0");vars["det"].set("0")
            else: win.destroy()
        ttk.Button(btns,text="Registrar e fechar",command=lambda:collect(False)).pack(side="right",padx=4)
        ttk.Button(btns,text="Registrar e próxima",style="Accent.TButton",
                   command=lambda:collect(True)).pack(side="right",padx=4)

    def edit_row(self,index=None):
        if not self.rows:return
        if index is None:
            if not hasattr(self,"tree"):return
            sel=self.tree.selection()
            if not sel:return
            index=int(sel[0])
        # O formulário trabalha sobre uma cópia normalizada. Assim, cancelar a
        # janela não altera a inspeção original e salvar não reaproveita dados
        # de outra linha.
        r=normalize_row(self.rows[index])
        win=tk.Toplevel(self);win.title(f"Inspeção {index+1} — detalhamento")
        win.geometry("820x760");win.minsize(700,560);win.configure(bg=BG);win.transient(self);win.grab_set()

        # Área rolável: o conteúdo pode crescer sem esconder o botão Salvar.
        outer=tk.Frame(win,bg=BG);outer.pack(fill="both",expand=True,padx=12,pady=12)
        canvas=tk.Canvas(outer,bg=BG,highlightthickness=0,bd=0)
        scroll=ttk.Scrollbar(outer,orient="vertical",command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left",fill="both",expand=True)
        scroll.pack(side="right",fill="y")
        frame=tk.Frame(canvas,bg=PANEL,highlightbackground=BORDER,highlightthickness=1)
        window_id=canvas.create_window((0,0),window=frame,anchor="nw")
        def _sync(_=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _width(e):
            canvas.itemconfigure(window_id,width=e.width)
        frame.bind("<Configure>",_sync)
        canvas.bind("<Configure>",_width)
        canvas.bind_all("<MouseWheel>",lambda e: canvas.yview_scroll(int(-e.delta/120),"units"))

        tk.Label(frame,text=f"INSPEÇÃO {index+1}",bg=PANEL,fg=TEXT,
                 font=("Segoe UI",15,"bold")).pack(anchor="w",padx=18,pady=(18,10))
        tk.Label(frame,text="Role a janela para acessar todos os campos. O botão Salvar permanece sempre visível.",
                 bg=PANEL,fg=ACCENT2,font=("Segoe UI",8)).pack(anchor="w",padx=18,pady=(0,8))
        vars={}
        for lab,key,values in [
            ("Identificação","id",None),("Vision AI","vision",["OK","NOK"]),
            ("Referência / gabarito","ref",REFS)]:
            row=tk.Frame(frame,bg=PANEL);row.pack(fill="x",padx=18,pady=4)
            tk.Label(row,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
            v=tk.StringVar(value=str(r.get(key,"")));vars[key]=v
            w=ttk.Combobox(row,textvariable=v,values=values,state="readonly") if values else ttk.Entry(row,textvariable=v)
            w.pack(fill="x")
        q=tk.Frame(frame,bg=PANEL);q.pack(fill="x",padx=18,pady=4)
        qv={}
        for j,(lab,key) in enumerate([("Real","real"),("Detectado","det")]):
            b=tk.Frame(q,bg=PANEL);b.grid(row=0,column=j,sticky="ew",padx=4);q.columnconfigure(j,weight=1)
            tk.Label(b,text=lab,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
            qv[key]=tk.StringVar(value=str(r[key]));ttk.Entry(b,textvariable=qv[key]).pack(fill="x")
        o=tk.Frame(frame,bg=PANEL);o.pack(fill="x",padx=18,pady=4)
        tk.Label(o,text="Observação",bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w")
        ov=tk.StringVar(value=r["obs"]);ttk.Entry(o,textvariable=ov).pack(fill="x")

        sec=self.card(frame,"DEFEITOS DETALHADOS — OPCIONAL",
                       "Use somente quando quiser saber quais tipos de defeito ocorreram. Uma inspeção pode ter vários tipos.")
        sec.pack(fill="both",expand=True,padx=18,pady=10)
        dt=tk.Frame(sec,bg=PANEL);dt.pack(fill="both",expand=True,padx=10,pady=5)
        cols=("tipo","real","det")
        dtt=ttk.Treeview(dt,columns=cols,show="headings",height=8)
        dscroll=ttk.Scrollbar(dt,orient="vertical",command=dtt.yview)
        dtt.configure(yscrollcommand=dscroll.set)
        for c,h,w in zip(cols,("Tipo de defeito","Real","Detectado"),(420,100,110)):
            dtt.heading(c,text=h);dtt.column(c,width=w,anchor="center")
        dtt.pack(side="left",fill="both",expand=True)
        dscroll.pack(side="right",fill="y")
        for d in r["defects"]:
            dtt.insert("", "end",values=(d["type"],d["real"],d["det"]))
        addbar=tk.Frame(sec,bg=PANEL);addbar.pack(fill="x",padx=10,pady=6)
        tv=tk.StringVar();rv=tk.StringVar(value="0");dv=tk.StringVar(value="0")
        ttk.Entry(addbar,textvariable=tv,width=35).pack(side="left",padx=3)
        ttk.Entry(addbar,textvariable=rv,width=8).pack(side="left",padx=3)
        ttk.Entry(addbar,textvariable=dv,width=8).pack(side="left",padx=3)
        def add_def():
            name=normalize_defect(tv.get())
            if not name:return
            try: rr=max(0,int(float(rv.get() or 0)));dd=max(0,int(float(dv.get() or 0)))
            except ValueError:
                messagebox.showerror("Defeito","Real e detectado devem ser numéricos.",parent=win);return
            dtt.insert("", "end",values=(name,rr,dd))
            tv.set("");rv.set("0");dv.set("0")
            _sync()
        ttk.Button(addbar,text="+ Adicionar tipo",command=add_def).pack(side="left",padx=5)
        def remove_def():
            for iid in dtt.selection(): dtt.delete(iid)
            _sync()
        ttk.Button(addbar,text="Excluir selecionado",command=remove_def).pack(side="right",padx=5)
        tk.Label(addbar,text="Tipo | Real | Detectado",bg=PANEL,fg=MUTED,
                 font=("Segoe UI",8)).pack(side="left",padx=5)
        _sync()

        # Rodapé fixo fora da área rolável: nunca fica escondido.
        footer=tk.Frame(win,bg=BG,highlightbackground=BORDER,highlightthickness=1)
        footer.pack(fill="x",side="bottom",padx=12,pady=(0,12))
        def save():
            # Monte um novo registro a partir do snapshot aberto nesta janela.
            # Cada campo é lido no momento do clique em Salvar para impedir
            # perda, repetição ou sobrescrita entre inspeções consecutivas.
            updated=normalize_row(r)
            updated["id"]=vars["id"].get().strip()
            updated["vision"]=vars["vision"].get() or "OK"
            updated["ref"]=vars["ref"].get() or "Não determinado"
            updated["obs"]=ov.get().strip()
            try:
                real=max(0,int(float(qv["real"].get() or 0)))
                detected=max(0,int(float(qv["det"].get() or 0)))
            except ValueError:
                messagebox.showerror("Dados","Real e Detectado devem ser numéricos.",parent=win);return
            updated["real"]=real
            updated["det"]=detected
            updated["defects"]=[]
            for iid in dtt.get_children():
                vals=dtt.item(iid,"values")
                updated["defects"].append({"type":normalize_defect(vals[0]),"real":int(vals[1]),"det":int(vals[2])})
            self.rows[index]=self.classify(updated)
            self.source_file=""
            win.destroy();self.refresh_all()
        ttk.Button(footer,text="Cancelar",command=win.destroy).pack(side="right",padx=5,pady=7)
        ttk.Button(footer,text="Salvar inspeção",style="Accent.TButton",command=save).pack(side="right",padx=5,pady=7)
        canvas.focus_set()
        def on_close():
            try: canvas.unbind_all("<MouseWheel>")
            except Exception: pass
            win.destroy()
        win.protocol("WM_DELETE_WINDOW",on_close)

    def delete_row(self):
        if not hasattr(self,"tree"):return
        sel=self.tree.selection()
        if not sel:return
        if not messagebox.askyesno("Excluir","Excluir a(s) inspeção(ões) selecionada(s)?"):return
        for i in sorted([int(x) for x in sel],reverse=True): self.rows.pop(i)
        self.refresh_all()

    # ---------- IMPORT ----------
    def _norm_header(self,v):
        return normalize_text(v).replace("º","").replace("°","")

    def _find_header_row(self,ws):
        targets=("IDENTIFICACAO","RESULTADO VISION AI","FALSO POSITIVO","FALSO NEGATIVO",
                 "VERDADEIRO POSITIVO","VERDADEIRO NEGATIVO","TP","FP","FN","VN")
        best=None
        for r in range(1,min(ws.max_row,80)+1):
            vals=[self._norm_header(ws.cell(r,c).value) for c in range(1,min(ws.max_column,50)+1)]
            score=sum(1 for t in targets if t in vals)
            if "IDENTIFICACAO" in vals: score+=2
            if score>=3 and (best is None or score>best[0]): best=(score,r)
        return best[1] if best else None

    def _header_map(self,ws,row):
        out={}
        for c in range(1,ws.max_column+1):
            h=self._norm_header(ws.cell(row,c).value)
            if h: out[h]=c
        return out

    def _num(self,v):
        try:
            if v is None or str(v).strip()=="":return 0
            return int(float(str(v).replace(",",".")))
        except Exception:return 0

    def _format_inspection_date(self, value):
        """Converte datas do Excel para dd/mm/aaaa sem inventar a data atual."""
        if value is None or str(value).strip()=="":
            return ""
        if isinstance(value, datetime):
            return value.strftime("%d/%m/%Y")
        try:
            from datetime import date as _date
            if isinstance(value, _date):
                return value.strftime("%d/%m/%Y")
        except Exception:
            pass
        text=str(value).strip()
        for fmt in ("%d/%m/%Y","%Y-%m-%d","%d-%m-%Y","%d/%m/%y"):
            try:
                return datetime.strptime(text,fmt).strftime("%d/%m/%Y")
            except Exception:
                pass
        return ""

    def _date_from_filename(self, path):
        """Tenta recuperar uma data explícita do nome do relatório (ex.: 21072026)."""
        name=Path(path).stem
        for raw in re.findall(r"(?<!\d)(\d{8})(?!\d)",name):
            for fmt in ("%d%m%Y","%Y%m%d"):
                try:return datetime.strptime(raw,fmt).strftime("%d/%m/%Y")
                except Exception:pass
        return ""

    def _metadata_aliases(self):
        """Aliases aceitos para preencher automaticamente o contexto da coleta."""
        return {
            "cliente": ["Cliente", "Customer", "Cliente / Customer", "OEM", "Montadora"],
            "local": ["Local / Planta", "Planta", "Plant", "Local", "Site", "Unidade", "Fábrica", "Fabrica"],
            "projeto": ["Projeto", "Project", "Projeto / Project", "Programa", "Program"],
            "linha": ["Linha / Estação", "Linha", "Line", "Estação", "Station", "Line / Station", "Linha de produção", "Processo"],
            "tecnico": ["Técnico", "Tecnico", "Technical", "Operador", "Operator", "Responsável", "Responsavel", "Responsável técnico", "Responsavel tecnico"],
            "data": ["Data da inspeção", "Data inspeção", "Data", "Date", "Inspection Date", "Data do teste", "Data do ensaio", "Dia"],
            "threshold": ["Threshold", "Confidence", "Confiança", "Grau de confiança", "Threshold / Grau de Confiança", "Limiar", "Confidence Threshold"],
            "gabarito": ["Método de referência", "Metodo de referencia", "Gabarito", "Método", "Metodo", "Reference", "Reference Method", "Ground Truth", "Truth Method"],
        }

    def _clean_metadata_value(self, value, key):
        if value is None:
            return ""
        if key == "data":
            d=self._format_inspection_date(value)
            if d:
                return d
        text=str(value).strip()
        if not text:
            return ""
        if self._norm_header(text) in {"N/D","ND","N/A","NA","NENHUM","NONE","NULL","-"}:
            return ""
        return text

    def _metadata_from_header_column(self, ws, header_row, aliases, key):
        hm=self._header_map(ws, header_row)
        normalized={self._norm_header(a) for a in aliases}
        candidates=[]
        for h,c in hm.items():
            if h in normalized or any(h.startswith(a+" ") or h.startswith(a+"/") for a in normalized):
                for rr in range(header_row+1, min(ws.max_row, header_row+30)+1):
                    v=self._clean_metadata_value(ws.cell(rr,c).value,key)
                    if v:
                        candidates.append(v)
        if not candidates:
            return ""
        if key == "data":
            from collections import Counter
            return Counter(candidates).most_common(1)[0][0]
        return candidates[0]

    def _metadata_from_labeled_cells(self, wb, aliases, key):
        """Procura campos em formulários/resumos: rótulo à esquerda ou acima do valor."""
        normalized={self._norm_header(a) for a in aliases}
        best=[]
        for ws in wb.worksheets:
            for r in range(1,min(ws.max_row,250)+1):
                for c in range(1,min(ws.max_column,40)+1):
                    h=self._norm_header(ws.cell(r,c).value)
                    if not h or (h not in normalized and not any(h.startswith(a+":") for a in normalized)):
                        continue
                    positions=[(r,c+1),(r,c+2),(r,c+3),(r+1,c),(r+2,c)]
                    for rr,cc in positions:
                        if rr>ws.max_row or cc>ws.max_column: continue
                        v=self._clean_metadata_value(ws.cell(rr,cc).value,key)
                        if v and self._norm_header(v) not in normalized:
                            best.append(v); break
        if not best:
            return ""
        if key == "data":
            from collections import Counter
            return Counter(best).most_common(1)[0][0]
        return best[0]

    def _extract_workbook_metadata(self, wb, data_ws, header_row, path):
        """Extrai contexto sem inventar valores: tabela > campos rotulados > data no nome do arquivo."""
        aliases=self._metadata_aliases()
        out={k:"" for k in aliases}; status={k:"não encontrado" for k in aliases}
        for key,als in aliases.items():
            v=self._metadata_from_header_column(data_ws, header_row, als, key)
            if not v:
                v=self._metadata_from_labeled_cells(wb, als, key)
            if not v and key=="data":
                v=self._date_from_filename(path)
            if v:
                out[key]=v; status[key]="identificado automaticamente"
        return out,status

    def import_excel(self):
        if openpyxl is None:
            messagebox.showerror("Excel","OpenPyXL não está instalado.\nExecute: py -3 -m pip install openpyxl")
            return
        path=filedialog.askopenfilename(title="Abrir planilha de assertividade",
                                        filetypes=[("Excel","*.xlsx"),("Todos","*.*")])
        if not path:return
        try:
            wb=openpyxl.load_workbook(path,data_only=True)
            # Não assumimos que a primeira aba é a de dados. Em relatórios AIA,
            # RESUMO costuma vir antes de INSPEÇÕES.
            candidates=[]
            for ws_try in wb.worksheets:
                hr_try=self._find_header_row(ws_try)
                if hr_try is not None:
                    hm_try=self._header_map(ws_try,hr_try)
                    score=0
                    for key in ("IDENTIFICACAO","RESULTADO VISION AI","TP","FP","FN","VN","REAL","DETECTADO"):
                        if key in hm_try: score += 1
                    candidates.append((score, ws_try, hr_try))
            if not candidates:
                raise ValueError("Não encontrei uma aba com a estrutura de INSPEÇÕES. Verifique se a planilha possui Identificação/Inspeção e colunas de resultado.")
            _, ws, hr = max(candidates, key=lambda x: x[0])
            hm=self._header_map(ws,hr)
            imported_meta, metadata_status = self._extract_workbook_metadata(wb, ws, hr, path)
            def col(*names):
                for n in names:
                    if self._norm_header(n) in hm:return hm[self._norm_header(n)]
                return None
            c_id=col("Identificação","N Inspeção","Nº Inspeção","Inspeção")
            c_vision=col("Resultado Vision AI","Vision AI","Vision")
            c_tp=col("Verdadeiro Positivo","VP - Verdadeiro Positivo","TP")
            c_tn=col("Verdadeiro Negativo","VN - Verdadeiro Negativo","VN","TN")
            c_fp=col("Falso Positivo","FP - Falso Positivo","FP")
            c_fn=col("Falso Negativo","FN - Falso Negativo","FN")
            c_real=col("Real","Real (defeitos)","Defeitos reais")
            c_det=col("Detectado","Detectados","Detecções","Detecções registradas")
            c_def=col("Defeito","Tipo de defeito")
            c_obs=col("Observação","Observações")
            c_date=col("Data da inspeção","Data inspeção","Data","Date")
            if c_id is None: raise ValueError("A planilha não possui coluna de Identificação/Inspeção.")
            rows=[];hist={"TP":0,"VN/TN":0,"FP":0,"FN":0,"N":0}
            imported_date=""
            imported_dates=[]
            for rr in range(hr+1,ws.max_row+1):
                ident=ws.cell(rr,c_id).value
                if ident is None or str(ident).strip()=="":continue
                if c_date:
                    row_date=self._format_inspection_date(ws.cell(rr,c_date).value)
                    if row_date:
                        imported_dates.append(row_date)
                r=blank_row()
                r["id"]=str(ident)
                r["vision"]=str(ws.cell(rr,c_vision).value or "OK").upper() if c_vision else "OK"
                r["ref"]="Importado — histórico"
                # Historical values remain historical. We store them in the row but do not
                # convert them into the new classification truth.
                r["tp"]=self._num(ws.cell(rr,c_tp).value) if c_tp else 0
                r["vn"]=self._num(ws.cell(rr,c_tn).value) if c_tn else 0
                r["fp"]=self._num(ws.cell(rr,c_fp).value) if c_fp else 0
                r["fn"]=self._num(ws.cell(rr,c_fn).value) if c_fn else 0
                r["detection_available"] = bool(c_real is not None and c_det is not None)
                r["real"]=self._num(ws.cell(rr,c_real).value) if c_real else 0
                r["det"]=self._num(ws.cell(rr,c_det).value) if c_det else 0
                r["obs"]=str(ws.cell(rr,c_obs).value or "") if c_obs else ""
                r["calc_status"]="Histórico"
                r["source"]="Excel histórico"
                # Legacy defect column: keep as optional detail with unknown counts.
                if c_def:
                    name=normalize_defect(ws.cell(rr,c_def).value)
                    if name:r["defects"]=[{"type":name,"real":0,"det":0}]
                rows.append(r)
                hist["TP"]+=r["tp"];hist["VN/TN"]+=r["vn"];hist["FP"]+=r["fp"];hist["FN"]+=r["fn"]
                hist["N"]+=r["tp"]+r["vn"]+r["fp"]+r["fn"]
            if not rows: raise ValueError("Nenhuma inspeção válida foi encontrada.")
            if imported_dates:
                from collections import Counter
                imported_date=Counter(imported_dates).most_common(1)[0][0]
            self._archive_current_collection()
            self.rows=rows;self.history=[hist];self.source_file=path
            # Importação inteligente do contexto. Nunca usa a data atual como fallback.
            if imported_date:
                imported_meta["data"] = imported_date
                metadata_status["data"] = "identificado automaticamente na tabela de inspeções"
            for key in self.meta:
                if key == "data":
                    continue
                value=imported_meta.get(key, "")
                if value:
                    self.meta[key].set(value)
            self.meta["data"].set(imported_meta.get("data", ""))
            if not self.meta["gabarito"].get():
                self.meta["gabarito"].set("Não determinado")
            self.refresh_all()
            labels=[("cliente","Cliente"),("local","Local / Planta"),("projeto","Projeto"),("linha","Linha / Estação"),("tecnico","Técnico"),("data","Data"),("threshold","Threshold"),("gabarito","Método de referência")]
            found=[label for key,label in labels if metadata_status.get(key,"").startswith("identificado")]
            missing=[label for key,label in labels if not metadata_status.get(key,"").startswith("identificado")]
            msg=(f"Planilha: {Path(path).name}\n\nInspeções importadas: {len(rows)}\n"
                 f"Histórico: TP={hist['TP']} | VN={hist['VN/TN']} | FP={hist['FP']} | FN={hist['FN']}\n\n"
                 f"IDENTIFICAÇÃO AUTOMÁTICA\nEncontrados: {len(found)}/8\n"
                 f"{', '.join(found) if found else 'Nenhum campo identificado'}\n")
            if missing:
                msg += f"\nPreencher/revisar: {', '.join(missing)}\n"
            msg += (f"\nData usada: {self.meta['data'].get() or 'N/D — preencher na tela Coleta'}\n\n"
                    "Os valores históricos foram preservados e não viraram gabarito do motor novo.\n"
                    "Revise os metadados antes de salvar/exportar, principalmente quando a planilha possuir mais de uma coleta.")
            messagebox.showinfo("Importação inteligente concluída", msg)
        except Exception as e:
            messagebox.showerror("Erro na importação",f"Não foi possível importar a planilha.\n\n{e}")

    # ---------- METRICS / OTHER VIEWS ----------
    def engineering_indicators(self, m=None, d=None):
        """Indicadores derivados para leitura técnica/engenharia.
        Não são métricas estatísticas oficiais; servem como indicadores operacionais.
        """
        m = m or self.display_metrics()
        d = d or self.defect_counts()
        n = m.get("N", 0) or 0
        real_positive = (m.get("TP", 0) or 0) + (m.get("FN", 0) or 0)
        predicted_positive = (m.get("TP", 0) or 0) + (m.get("FP", 0) or 0)
        return {
            "prevalencia": safe(real_positive, n),
            "taxa_alarme": safe(predicted_positive, n),
            "fp_100": safe(m.get("FP", 0) * 100, n),
            "fn_100": safe(m.get("FN", 0) * 100, n),
            "real_por_inspecao": safe(d.get("real", 0), n),
            "det_por_inspecao": safe(d.get("det", 0), n),
        }

    def technical_diagnosis(self, m, d):
        """Gera leitura técnica baseada nos indicadores, sem criar um score artificial."""
        notes = []
        n = m.get("N", 0) or 0
        recall = m.get("Recall / Sensibilidade")
        precision = m.get("Precision")
        fpr = m.get("FPR")
        fnr = m.get("FNR")

        if n < 30:
            notes.append("AMOSTRA PEQUENA: use os resultados como diagnóstico inicial; aumente o número de inspeções antes de concluir desempenho.")
        if recall is not None:
            if recall >= 0.90:
                notes.append("RECALL FORTE: a capacidade de encontrar inspeções realmente NOK está alta.")
            elif recall < 0.80:
                notes.append("RECALL DE ATENÇÃO: há risco relevante de deixar casos NOK passarem como OK.")
        if precision is not None:
            if precision < 0.70:
                notes.append("PRECISION DE ATENÇÃO: parte importante dos alarmes NOK não corresponde a casos realmente NOK.")
            elif precision >= 0.90:
                notes.append("PRECISION FORTE: a maior parte das classificações NOK corresponde a casos realmente NOK.")
        if fpr is not None and fpr >= 0.20:
            notes.append("FALSO POSITIVO ELEVADO: vale investigar threshold, ruído de imagem, iluminação, ROI e condições de aquisição.")
        if fnr is not None and fnr >= 0.10:
            notes.append("FALSO NEGATIVO RELEVANTE: vale investigar sensibilidade, cobertura da inspeção, velocidade e condições de aquisição.")
        if d.get("real", 0) and d.get("det", 0) > d.get("real", 0):
            notes.append("FP DE DETECÇÃO ELEVADO: existem mais detecções que defeitos reais registrados; avaliar duplicidade, granularidade, threshold, ruído e correspondência das ocorrências.")
        if not notes:
            notes.append("SEM ALERTA AUTOMÁTICO: os indicadores atuais não ultrapassaram os limites de diagnóstico configurados.")
        return notes

    def view_tecnica(self):
        fonte = "HISTÓRICO DA PLANILHA" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
        self.section_title(
            "ANÁLISE TÉCNICA  /  DIAGNÓSTICO",
            f"Autaza Inspection Analytics • {fonte} • métricas oficiais + leitura técnica sem misturar unidades")

        m = self.display_metrics()
        d = self.defect_counts()
        dt = self.detection_totals()
        e = self.engineering_indicators(m, d)

        grid = tk.Frame(self.content, bg=BG)
        grid.pack(fill="x")
        cards = [
            ("N • Inspeções", str(m["N"]), TEXT),
            ("Accuracy", pct(m["Accuracy"]), ACCENT2),
            ("Precision", pct(m["Precision"]), ACCENT2),
            ("Recall", pct(m["Recall / Sensibilidade"]), GOOD),
            ("Specificity", pct(m["Specificity"]), ACCENT2),
            ("FPR", pct(m["FPR"]), BAD),
            ("FNR", pct(m["FNR"]), WARN),
            ("F1", pct(m["F1"]), ACCENT2),
            ("MCC", "N/D" if m["MCC"] is None else f"{m['MCC']:.4f}", ACCENT2),
            ("Prevalência", pct(e["prevalencia"]), ACCENT2),
            ("Precision detecção", pct(dt["precision"]), ACCENT),
            ("Recall detecção", pct(dt["recall"]), GOOD),
        ]
        for i, (lab, val, col) in enumerate(cards):
            r, c = divmod(i, 5)
            self.metric_card(grid, lab, val, col).grid(row=r, column=c, sticky="ew", padx=4, pady=4)
        for c in range(5):
            grid.columnconfigure(c, weight=1)

        lower = tk.Frame(self.content, bg=BG)
        lower.pack(fill="both", expand=True, pady=8)

        matrix = self.chart_card(lower, "MATRIZ DE CONFUSÃO", "Unidade estatística: inspeção / carroceria / peça")
        matrix.grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))
        txt = (
            f"                         PREDITO NOK       PREDITO OK\n"
            f"REAL NOK / DEFEITO       TP = {m['TP']:<8}     FN = {m['FN']}\n"
            f"REAL OK                  FP = {m['FP']:<8}     VN = {m['VN/TN']}\n\n"
            f"N = {m['N']}\n\n"
            f"FPR = {pct(m['FPR'])}    FNR = {pct(m['FNR'])}"
        )
        tk.Label(matrix, text=txt, bg=PANEL, fg=TEXT, font=("Consolas", 10),
                 justify="left").pack(anchor="w", padx=16, pady=14)

        errors = self.chart_card(lower, "PERFIL DE ERROS", "Volume de erros e acertos da classificação")
        errors.grid(row=0, column=1, sticky="nsew", padx=(4, 0), pady=(0, 4))
        self.draw_bar_chart(
            errors,
            [("TP", m["TP"], GOOD), ("TN", m["VN/TN"], GOOD),
             ("FP", m["FP"], BAD), ("FN", m["FN"], WARN)],
            max_value=max(m["TP"], m["VN/TN"], m["FP"], m["FN"], 1)
        )

        diag = self.card(lower, "DIAGNÓSTICO TÉCNICO", "Leitura automática para apoiar investigação — não substitui análise de processo.")
        diag.grid(row=1, column=0, sticky="nsew", padx=(0, 4), pady=(4, 0))
        for note in self.technical_diagnosis(m, d):
            tk.Label(diag, text="• " + note, bg=PANEL, fg=TEXT,
                     wraplength=520, justify="left", font=("Segoe UI", 9)).pack(
                         anchor="w", padx=16, pady=5)

        detcard = self.card(lower, "DETecção × CLASSIFICAÇÃO", "FP/FN de detecção não são os mesmos FP/FN da matriz OK/NOK.")
        detcard.grid(row=2, column=0, columnspan=2, sticky="ew", padx=0, pady=(8, 0))
        txtd=(f"CLASSIFICAÇÃO: FP = {m['FP']} | FN = {m['FN']}\n"
              f"DETECÇÃO: TP = {dt['tp']} | FP = {dt['fp']} | FN = {dt['fn']}\n"
              f"Detecção por contagem: Precision = {pct(dt['precision'])} | Recall = {pct(dt['recall'])}\n\n"
              "A estimativa de detecção usa quantidade quando não existe correspondência individual; ela não afirma correspondência espacial entre defeitos.")
        tk.Label(detcard, text=txtd, bg=PANEL, fg=TEXT, font=("Consolas", 9), justify="left", wraplength=1100).pack(anchor="w", padx=16, pady=12)

        op = self.card(lower, "INDICADORES OPERACIONAIS", "Complementares às métricas estatísticas; não devem ser interpretados como Accuracy/Recall.")
        op.grid(row=1, column=1, sticky="nsew", padx=(4, 0), pady=(4, 0))
        operational = [
            ("Alarmes NOK / inspeção", pct(e["taxa_alarme"])),
            ("Falsos positivos / 100 inspeções", "N/D" if e["fp_100"] is None else f"{e['fp_100']:.1f}"),
            ("Falsos negativos / 100 inspeções", "N/D" if e["fn_100"] is None else f"{e['fn_100']:.1f}"),
            ("Defeitos reais / inspeção", "N/D" if e["real_por_inspecao"] is None else f"{e['real_por_inspecao']:.2f}"),
            ("Detecções / inspeção", "N/D" if e["det_por_inspecao"] is None else f"{e['det_por_inspecao']:.2f}"),
            ("FP detecção / inspeção", "N/D" if dt["available_rows"]==0 else f"{dt['fp']/max(dt['available_rows'],1):.2f}"),
        ]
        for lab, val in operational:
            row = tk.Frame(op, bg=PANEL)
            row.pack(fill="x", padx=16, pady=5)
            tk.Label(row, text=lab, bg=PANEL, fg=MUTED, font=("Segoe UI", 9)).pack(side="left")
            tk.Label(row, text=val, bg=PANEL, fg=TEXT,
                     font=("Segoe UI", 10, "bold")).pack(side="right")

        lower.columnconfigure(0, weight=1)
        lower.columnconfigure(1, weight=1)
        lower.rowconfigure(0, weight=1)
        lower.rowconfigure(1, weight=1)

    def _draw_audit_reconciliation_chart(self, parent, pairs):
        """Conciliação visual: cinza para histórico e azul para motor novo."""
        canvas=tk.Canvas(parent,bg=PANEL,highlightthickness=0,height=255)
        canvas.pack(fill="both",expand=True,padx=12,pady=(2,12)); canvas.update_idletasks()
        w,h=max(canvas.winfo_width(),700),255; left,right,top,bottom=142,64,32,34
        maxv=max([max(a,b) for _,a,b in pairs]+[1]); plot_w=w-left-right
        for i in range(5):
            x=left+plot_w*i/4; canvas.create_line(x,top,x,h-bottom,fill="#29404F")
            canvas.create_text(x,h-bottom+13,text=f"{maxv*i/4:.0f}",fill=MUTED,font=("Segoe UI",7))
        canvas.create_text(left,14,text="● Histórico",fill="#8FA9B8",anchor="w",font=("Segoe UI",8,"bold"))
        canvas.create_text(left+110,14,text="● Motor novo",fill=ACCENT2,anchor="w",font=("Segoe UI",8,"bold"))
        row_h=(h-top-bottom)/max(len(pairs),1)
        for i,(label,hist,motor) in enumerate(pairs):
            y=top+row_h*(i+.5); xh=left+plot_w*hist/maxv; xm=left+plot_w*motor/maxv
            color=GOOD if motor==hist else (BAD if motor>hist else WARN)
            canvas.create_text(left-10,y,text=label,fill=TEXT,anchor="e",font=("Segoe UI",9,"bold"))
            canvas.create_line(xh,y,xm,y,fill=color,width=4)
            canvas.create_oval(xh-5,y-5,xh+5,y+5,fill="#8FA9B8",outline="")
            canvas.create_oval(xm-6,y-6,xm+6,y+6,fill=ACCENT2,outline="")
            canvas.create_text(w-right+5,y,text=f"{motor-hist:+d}",fill=color,anchor="w",font=("Segoe UI",9,"bold"))
        canvas.create_text(w-right+4,h-bottom+13,text="Δ",fill=MUTED,anchor="w",font=("Segoe UI",7))

    def view_auditoria(self):
        self.section_title("AUDITORIA  /  CONCILIAÇÃO", "Histórico importado × motor estruturado. A diferença aponta o que revisar, sem mesclar metodologias.")
        if not self.history:
            f=self.card(self.content,"Nenhum histórico importado")
            f.pack(fill="x",pady=5)
            tk.Label(f,text="Abra uma planilha Excel para carregar os valores históricos.",
                     bg=PANEL,fg=TEXT).pack(anchor="w",padx=16,pady=16)
            return
        hist=self.history[-1];m=self.display_metrics()
        pairs=[("TP correto",int(hist.get("TP",0) or 0),int(m.get("TP",0) or 0)),
               ("TN/VN correto",int(hist.get("VN/TN",0) or 0),int(m.get("VN/TN",0) or 0)),
               ("FP alarme indevido",int(hist.get("FP",0) or 0),int(m.get("FP",0) or 0)),
               ("FN escape",int(hist.get("FN",0) or 0),int(m.get("FN",0) or 0)),
               ("N inspeções",int(hist.get("N",0) or 0),int(m.get("N",0) or 0))]
        deltas=[abs(motor-old) for _,old,motor in pairs]; same=sum(1 for _,old,motor in pairs if old==motor)
        max_pair=pairs[deltas.index(max(deltas))]
        cards=tk.Frame(self.content,bg=BG); cards.pack(fill="x",pady=(0,7))
        for i,(label,value,color) in enumerate([
            ("ITENS CONCILIADOS",f"{same}/{len(pairs)}",GOOD if same==len(pairs) else WARN),
            ("DIVERGÊNCIA ABSOLUTA",str(sum(deltas)),BAD if sum(deltas) else GOOD),
            ("MAIOR PONTO DE REVISÃO",f"{max_pair[0]}  {max(deltas):+d}",BAD if max(deltas) else GOOD),
            ("FONTE HISTÓRICA",Path(self.source_file).name if self.source_file else "Importada",ACCENT2),
        ]):
            c=self.metric_card(cards,label,value,color); c.grid(row=0,column=i,sticky="ew",padx=(0,5) if i<3 else 0); cards.columnconfigure(i,weight=1)
        visual=self.chart_card(self.content,"MAPA DE CONCILIAÇÃO", "Cinza = histórico | azul = motor novo | Δ mostra a diferença a investigar.")
        visual.pack(fill="x",pady=(0,7)); self._draw_audit_reconciliation_chart(visual,pairs)
        f=self.card(self.content,"RASTRO DE CONCILIAÇÃO", "Diferença = motor novo − histórico.");f.pack(fill="x")
        tree=ttk.Treeview(f,columns=("ind","hist","new","diff"),show="headings",height=6)
        for c,h in zip(("ind","hist","new","diff"),("Indicador","Histórico","Motor novo","Diferença")):
            tree.heading(c,text=h);tree.column(c,width=180,anchor="center")
        tree.pack(fill="x",padx=12,pady=12)
        tree.tag_configure("ok",foreground=GOOD); tree.tag_configure("review",foreground=BAD)
        for label,hv,mv in pairs:
            delta=mv-hv; tree.insert("", "end",values=(label,hv,mv,f"{delta:+d}"),tags=("ok" if delta==0 else "review",))
        a=self.card(self.content,"INTERPRETAÇÃO")
        a.pack(fill="both",expand=True,pady=8)
        tk.Label(a,text=("Priorize os itens em vermelho: confirme unidade estatística, filtros da planilha, regras de classificação e recorte da amostra. "
                         "Diferenças podem ser legítimas quando as metodologias forem distintas; o software não transforma o histórico em verdade de referência."),bg=PANEL,fg=WARN,wraplength=1100,
                 justify="left").pack(anchor="w",padx=16,pady=16)

    def view_engenharia(self):
        fonte = "HISTÓRICO DA PLANILHA" if (self.source_file and self.history) else "MOTOR NOVO / ENTRADA MANUAL"
        self.section_title(
            "ENGENHARIA  /  DECISÃO",
            f"Autaza Inspection Analytics • {fonte} • indicadores para investigação de processo, câmera, threshold e operação")

        m = self.display_metrics()
        d = self.defect_counts()
        dt = self.detection_totals()
        e = self.engineering_indicators(m, d)

        # Cabeçalho de contexto
        context = self.card(self.content, "CONTEXTO DO TESTE", "Condições registradas no momento da coleta/importação.")
        context.pack(fill="x", pady=(0, 6))
        ctx = tk.Frame(context, bg=PANEL)
        ctx.pack(fill="x", padx=14, pady=(0, 12))
        items = [
            ("Cliente", self.meta["cliente"].get() or "N/D"),
            ("Planta", self.meta["local"].get() or "N/D"),
            ("Projeto", self.meta["projeto"].get() or "N/D"),
            ("Linha / Estação", self.meta["linha"].get() or "N/D"),
            ("Threshold", self.meta["threshold"].get() or "N/D"),
            ("Referência", self.meta["gabarito"].get() or "N/D"),
        ]
        for i, (lab, val) in enumerate(items):
            r, c = divmod(i, 3)
            box = tk.Frame(ctx, bg=PANEL)
            box.grid(row=r, column=c, sticky="ew", padx=8, pady=4)
            tk.Label(box, text=lab, bg=PANEL, fg=MUTED,
                     font=("Segoe UI", 8, "bold")).pack(anchor="w")
            tk.Label(box, text=val, bg=PANEL, fg=TEXT,
                     font=("Segoe UI", 9)).pack(anchor="w")
            ctx.columnconfigure(c, weight=1)

        grid = tk.Frame(self.content, bg=BG)
        grid.pack(fill="x", pady=2)
        cards = [
            ("Prevalência", pct(e["prevalencia"]), ACCENT2),
            ("Taxa de alarme NOK", pct(e["taxa_alarme"]), ACCENT),
            ("FP / 100 inspeções", "N/D" if e["fp_100"] is None else f"{e['fp_100']:.1f}", BAD),
            ("FN / 100 inspeções", "N/D" if e["fn_100"] is None else f"{e['fn_100']:.1f}", WARN),
            ("Defeitos reais / inspeção", "N/D" if e["real_por_inspecao"] is None else f"{e['real_por_inspecao']:.2f}", GOOD),
            ("Detecções / inspeção", "N/D" if e["det_por_inspecao"] is None else f"{e['det_por_inspecao']:.2f}", ACCENT2),
        ]
        for i, (lab, val, col) in enumerate(cards):
            self.metric_card(grid, lab, val, col).grid(row=0, column=i, sticky="ew", padx=4, pady=4)
            grid.columnconfigure(i, weight=1)

        lower = tk.Frame(self.content, bg=BG)
        lower.pack(fill="both", expand=True, pady=8)

        priority = self.card(lower, "PRIORIZAÇÃO DE INVESTIGAÇÃO", "A ordem abaixo é derivada dos erros observados; não cria um score proprietário.")
        priority.grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))

        fpr = m["FPR"]
        fnr = m["FNR"]
        recall = m["Recall / Sensibilidade"]
        precision = m["Precision"]
        priorities = []
        if fpr is not None and fpr >= 0.20:
            priorities.append(("01", "REDUZIR FALSOS POSITIVOS", "Investigar threshold, ruído, iluminação, ROI, contraste e condições de aquisição."))
        if fnr is not None and fnr >= 0.10:
            priorities.append(("02", "REDUZIR FALSOS NEGATIVOS", "Investigar sensibilidade, cobertura da inspeção, velocidade, foco e condições de aquisição."))
        if precision is not None and precision < 0.70:
            priorities.append(("03", "VALIDAR CRITÉRIO DE NOK", "Muitos alarmes não confirmados indicam necessidade de revisar critério de decisão e/ou referência."))
        if recall is not None and recall < 0.80:
            priorities.append(("04", "VALIDAR COBERTURA", "A capacidade de encontrar casos reais está abaixo do esperado para uma inspeção robusta."))
        if not priorities:
            priorities.append(("01", "ACOMPANHAR ESTABILIDADE", "Os indicadores não acionaram prioridades críticas. Repetir testes em condições variadas antes de concluir estabilidade."))

        for num, title, desc in priorities:
            row = tk.Frame(priority, bg=PANEL)
            row.pack(fill="x", padx=16, pady=7)
            tk.Label(row, text=num, bg="#0E4560", fg=TEXT,
                     font=("Segoe UI", 9, "bold"), width=4, pady=5).pack(side="left")
            body = tk.Frame(row, bg=PANEL)
            body.pack(side="left", fill="x", expand=True, padx=10)
            tk.Label(body, text=title, bg=PANEL, fg=TEXT,
                     font=("Segoe UI", 9, "bold")).pack(anchor="w")
            tk.Label(body, text=desc, bg=PANEL, fg=MUTED,
                     wraplength=500, justify="left", font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))

        quality = self.chart_card(lower, "INDICADORES PARA DECISÃO", "Comparação visual das métricas mais relevantes para engenharia")
        quality.grid(row=0, column=1, sticky="nsew", padx=(4, 0), pady=(0, 4))
        self.draw_percent_chart(quality, [
            ("Recall", m["Recall / Sensibilidade"], GOOD),
            ("Precision", m["Precision"], ACCENT),
            ("Specificity", m["Specificity"], ACCENT2),
            ("FPR", m["FPR"], BAD),
            ("FNR", m["FNR"], WARN),
        ])

        actions = self.card(lower, "LEITURA DE ENGENHARIA", "Como usar os números sem transformar uma métrica em 'nota'.")
        actions.grid(row=1, column=0, sticky="nsew", padx=(0, 4), pady=(4, 0))
        lines = [
            f"• Recall {pct(m['Recall / Sensibilidade'])}: priorize a capacidade de encontrar o que realmente é defeito.",
            f"• Precision {pct(m['Precision'])}: avalie quanto retrabalho/alarmes podem estar vindo de falsos positivos.",
            f"• FPR {pct(m['FPR'])}: observe a carga de alarmes em peças realmente OK.",
            f"• FNR {pct(m['FNR'])}: observe o risco de defeitos passarem como OK.",
            f"• FP de detecção {dt['fp']}: representa detecções extras em relação aos defeitos reais registrados.",
            f"• FN de detecção {dt['fn']}: representa defeitos reais sem detecção correspondente por contagem.",
            "• Accuracy não deve ser usada isoladamente, principalmente quando OK e NOK são desbalanceados.",
        ]
        for line in lines:
            tk.Label(actions, text=line, bg=PANEL, fg=TEXT,
                     wraplength=520, justify="left", font=("Segoe UI", 9)).pack(anchor="w", padx=16, pady=5)

        occurrence = self.card(lower, "OCORRÊNCIAS / DETECÇÃO", "Unidade diferente da classificação da inspeção.")
        occurrence.grid(row=1, column=1, sticky="nsew", padx=(4, 0), pady=(4, 0))
        vals = [
            ("Defeitos reais registrados", d["real"]),
            ("Detecções registradas", d["det"]),
            ("TP de detecção (por contagem)", d["tp"]),
            ("FP de detecção (extras)", d["fp"]),
            ("FN de detecção (não detectados)", d["fn"]),
            ("Precision de detecção", pct(dt["precision"])),
            ("Recall de detecção", pct(dt["recall"])),
        ]
        for lab, val in vals:
            row = tk.Frame(occurrence, bg=PANEL)
            row.pack(fill="x", padx=16, pady=5)
            tk.Label(row, text=lab, bg=PANEL, fg=MUTED, font=("Segoe UI", 9)).pack(side="left")
            tk.Label(row, text=str(val), bg=PANEL, fg=TEXT,
                     font=("Segoe UI", 10, "bold")).pack(side="right")
        tk.Label(occurrence,
                 text=("A Coleta Rápida pode operar sem tipo de defeito. "
                       "Quando o tipo é informado, a aba 'Por Defeito' permite aprofundar a análise."),
                 bg=PANEL, fg=ACCENT2, wraplength=520, justify="left",
                 font=("Segoe UI", 8)).pack(anchor="w", padx=16, pady=(10, 14))

        lower.columnconfigure(0, weight=1)
        lower.columnconfigure(1, weight=1)
        lower.rowconfigure(0, weight=1)
        lower.rowconfigure(1, weight=1)

    def view_defeitos(self):
        tk.Label(self.content,text="ANÁLISE POR TIPO DE DEFEITO",bg=BG,fg=TEXT,
                 font=("Segoe UI",18,"bold")).pack(anchor="w")
        tk.Label(self.content,text="Análise opcional: só aparecem tipos quando o técnico decidiu detalhá-los.",
                 bg=BG,fg=MUTED).pack(anchor="w",pady=(2,10))
        d=self.defect_counts()
        if not d["by"]:
            f=self.card(self.content,"Nenhum tipo detalhado")
            f.pack(fill="both",expand=True)
            tk.Label(f,text=("Isso é normal na Coleta Rápida. O tipo do defeito não é obrigatório. "
                             "Use 'Detalhar defeitos' apenas quando a análise por categoria for necessária."),
                     bg=PANEL,fg=TEXT,wraplength=900,justify="left").pack(anchor="w",padx=18,pady=18)
            return
        f=self.card(self.content,"RESUMO POR TIPO")
        f.pack(fill="both",expand=True)
        tree=ttk.Treeview(f,columns=("def","real","det","tp","fp","fn","rec","prec"),
                          show="headings")
        heads=("Defeito","Real","Detectado","TP","FP","FN","Recall","Precision")
        for c,h in zip(tree["columns"],heads):
            tree.heading(c,text=h);tree.column(c,width=130,anchor="center")
        tree.column("def",width=260,anchor="w")
        tree.pack(fill="both",expand=True,padx=12,pady=12)
        for name,v in sorted(d["by"].items()):
            mm=classification_metrics(v["tp"],0,v["fp"],v["fn"])
            tree.insert("", "end",values=(name,v["real"],v["det"],v["tp"],v["fp"],v["fn"],
                                          pct(mm["Recall / Sensibilidade"]),pct(mm["Precision"])))

    # ---------- COMPARAÇÃO MULTICOLETA ----------
    def edit_calibration_project(self):
        """Identifica o conjunto de coletas sem misturar dados de projetos distintos."""
        win=tk.Toplevel(self); win.title("Projeto de calibração"); win.geometry("620x430"); win.configure(bg=BG); win.transient(self); win.grab_set()
        box=self.card(win,"PROJETO DE CALIBRAÇÃO","Agrupe as coletas e registre o contexto da decisão."); box.pack(fill="both",expand=True,padx=16,pady=16)
        fields=[("Código do projeto","id"),("Nome do projeto","name"),("Objetivo de calibração","objective"),("Responsável","responsible")]
        vals={k:tk.StringVar(value=self.calibration_project.get(k,"")) for _,k in fields}
        for label,key in fields:
            tk.Label(box,text=label,bg=PANEL,fg=MUTED,font=("Segoe UI",8,"bold")).pack(anchor="w",padx=18,pady=(10,1))
            ttk.Entry(box,textvariable=vals[key]).pack(fill="x",padx=18)
        def save():
            self.calibration_project={k:v.get().strip() for k,v in vals.items()}
            self.calibration_project["created_at"]=self.calibration_project.get("created_at") or datetime.now().isoformat()
            self.status.set("Projeto de calibração atualizado."); win.destroy(); self.refresh_all()
        ttk.Button(box,text="Salvar projeto de calibração",style="Accent.TButton",command=save).pack(anchor="e",padx=18,pady=18)

    def _calibration_score(self, item):
        """Prioriza aderência às metas; em empate privilegia F1, recall e menor FP."""
        m=item.get("metrics",{}); exp=item.get("expectation",{}) or {}
        checks=[("accuracy","Accuracy",True),("precision","Precision",True),("recall","Recall / Sensibilidade",True),("f1","F1",True)]
        met=total=0
        for key, metric, higher in checks:
            target=exp.get(key); actual=m.get(metric)
            if target is not None and actual is not None:
                total+=1; met += int(actual >= target if higher else actual <= target)
        n=m.get("N",0) or 0
        for key in ("fp100","fn100"):
            target=exp.get(key)
            if target is not None and n:
                total+=1; actual=m.get("FP" if key=="fp100" else "FN",0)/n*100
                met += int(actual <= target)
        return (met, total, m.get("F1") or -1, m.get("Recall / Sensibilidade") or -1, -(m.get("FP",0)))

    def _best_configuration(self, datasets=None):
        items=datasets if datasets is not None else self.calibration_history
        valid=[x for x in items if x.get("metrics",{}).get("N",0)>0]
        return max(valid,key=self._calibration_score) if valid else None

    def _threshold_value(self, item):
        try: return float(str(item.get("threshold","")).replace(",",".").replace("%",""))
        except (TypeError,ValueError): return None

    def _draw_threshold_chart(self,parent,datasets):
        points=[(self._threshold_value(x),x) for x in datasets]; points=[p for p in points if p[0] is not None and p[1].get("metrics",{}).get("F1") is not None]
        canvas=tk.Canvas(parent,bg=PANEL,highlightthickness=0,height=260); canvas.pack(fill="both",expand=True,padx=10,pady=8); canvas.update_idletasks()
        w=max(canvas.winfo_width(),620); h=260; left,right,top,bottom=55,28,22,45
        if not points:
            canvas.create_text(w/2,h/2,text="Informe um Threshold numérico em pelo menos uma coleta para gerar o gráfico.",fill=MUTED,font=("Segoe UI",9)); return
        points.sort(key=lambda x:x[0]); xs=[p[0] for p in points]; lo,hi=min(xs),max(xs)
        if lo==hi: lo-=1; hi+=1
        for i in range(5):
            y=top+(h-top-bottom)*(1-i/4); canvas.create_line(left,y,w-right,y,fill="#29404F")
            canvas.create_text(left-7,y,text=f"{i*25}%",anchor="e",fill=MUTED,font=("Segoe UI",7))
        prev=None
        for threshold,item in points:
            x=left+(threshold-lo)/(hi-lo)*(w-left-right); f1=item["metrics"]["F1"]; y=top+(1-f1)*(h-top-bottom)
            if prev: canvas.create_line(prev[0],prev[1],x,y,fill=ACCENT,width=2)
            canvas.create_oval(x-5,y-5,x+5,y+5,fill=GOOD if item is self._best_configuration(datasets) else ACCENT2,outline="")
            canvas.create_text(x,y-12,text=f"{f1*100:.1f}%",fill=TEXT,font=("Segoe UI",7,"bold")); canvas.create_text(x,h-bottom+15,text=str(threshold),fill=MUTED,font=("Segoe UI",7)); prev=(x,y)
        canvas.create_text(left,h-bottom+30,text="Threshold",fill="#6E8E9E",font=("Segoe UI",7))

    def view_calibracao(self):
        self.section_title("PROJETO DE CALIBRAÇÃO","Rastreabilidade, melhor configuração encontrada e relação Threshold x Performance.")
        p=self.calibration_project; info=self.card(self.content,"IDENTIFICAÇÃO DO PROJETO",p.get("name") or "Projeto ainda não identificado"); info.pack(fill="x",pady=4)
        tk.Label(info,text=f"Código: {p.get('id') or 'N/D'}   |   Responsável: {p.get('responsible') or 'N/D'}\nObjetivo: {p.get('objective') or 'N/D'}",bg=PANEL,fg=TEXT,justify="left").pack(anchor="w",padx=16,pady=12)
        ttk.Button(info,text="Editar identificação",command=self.edit_calibration_project).pack(anchor="e",padx=14,pady=(0,12))
        toolbar=tk.Frame(self.content,bg=BG); toolbar.pack(fill="x",pady=5)
        ttk.Button(toolbar,text="Registrar coleta atual no histórico",style="Accent.TButton",command=self._archive_current_from_view).pack(side="left",padx=3)
        ttk.Button(toolbar,text="Gerar relatório executivo",command=self.export_executive_report).pack(side="left",padx=3)
        datasets=self.calibration_history or self.comparison_archive
        best=self._best_configuration(datasets)
        bestbox=self.card(self.content,"MELHOR CONFIGURAÇÃO ENCONTRADA","Critério: maior atendimento às metas; depois F1, Recall e menor FP."); bestbox.pack(fill="x",pady=4)
        text="Ainda não há coleta classificada suficiente para recomendar uma configuração."
        if best:
            m=best["metrics"]; text=(f"{best.get('name')} | Threshold: {best.get('threshold','N/D')} | F1: {pct(m.get('F1'))} | Recall: {pct(m.get('Recall / Sensibilidade'))} | Precision: {pct(m.get('Precision'))} | N: {m.get('N',0)}")
        tk.Label(bestbox,text=text,bg=PANEL,fg=GOOD if best else MUTED,font=("Segoe UI",10,"bold"),wraplength=1100,justify="left").pack(anchor="w",padx=16,pady=14)
        chart=self.card(self.content,"THRESHOLD x PERFORMANCE","F1 por configuração registrada; o ponto verde representa a recomendação."); chart.pack(fill="both",expand=True,pady=4); self._draw_threshold_chart(chart,datasets)
        hist=self.card(self.content,"HISTÓRICO RASTREÁVEL",f"{len(datasets)} coleta(s) no projeto"); hist.pack(fill="x",pady=4)
        tree=ttk.Treeview(hist,columns=("id","data","coleta","threshold","n","f1","origem","rastro"),show="headings",height=min(9,max(3,len(datasets))))
        for col,label,width in [("id","Registro",105),("data","Data/Hora",150),("coleta","Coleta",145),("threshold","Threshold",90),("n","N",55),("f1","F1",80),("origem","Origem",210),("rastro","Hash",100)]: tree.heading(col,text=label); tree.column(col,width=width,anchor="center" if col not in ("coleta","origem") else "w")
        tree.pack(fill="x",padx=12,pady=12)
        for item in datasets:
            tree.insert("","end",values=(item.get("record_id","N/D"),item.get("created_at","N/D")[:19],item.get("name","Coleta"),item.get("threshold","N/D"),item.get("metrics",{}).get("N",0),pct(item.get("metrics",{}).get("F1")),Path(item.get("source_file","")).name or "Manual",item.get("trace_hash","")[:10]))

    def _current_collection_has_data(self):
        return bool(self.rows or self.source_file or self.history or any(v.get() for k, v in self.meta.items() if k not in ("data", "gabarito")))

    def _archive_current_collection(self, label=None):
        """Arquiva um snapshot da coleta ativa para comparação posterior."""
        if not self._current_collection_has_data():
            return None
        m = self.display_metrics()
        d = self.defect_counts()
        meta = {k: v.get() for k, v in self.meta.items()}
        seq = len(self.comparison_archive) + 1
        payload = {"metadata":meta,"metrics":m,"defects":d,"source_file":self.source_file or "","project":self.calibration_project}
        item = {
            "name": label or f"Coleta {seq:02d}",
            "created_at": datetime.now().isoformat(),
            "metadata": meta,
            "source_file": self.source_file or "",
            "metrics": dict(m),
            "defects": dict(d),
            "threshold": meta.get("threshold") or "N/D",
            "date": meta.get("data") or datetime.now().strftime("%d/%m/%Y"),
            "expectation": self.expectation_snapshot(),
            "project": dict(self.calibration_project),
            "record_id": f"CAL-{len(self.calibration_history)+1:03d}",
            "trace_hash": hashlib.sha256(json.dumps(payload,ensure_ascii=False,sort_keys=True,default=str).encode("utf-8")).hexdigest(),
        }
        self.comparison_archive.append(item)
        self.calibration_history.append(dict(item))
        if len(self.comparison_archive) > 50:
            self.comparison_archive = self.comparison_archive[-50:]
        if len(self.calibration_history) > 50:
            self.calibration_history = self.calibration_history[-50:]
        return item

    def _comparison_datasets(self):
        datasets = list(self.comparison_archive)
        if self._current_collection_has_data():
            m = self.display_metrics()
            d = self.defect_counts()
            meta = {k: v.get() for k, v in self.meta.items()}
            datasets.append({
                "name": "Coleta atual", "created_at": datetime.now().isoformat(),
                "metadata": meta, "source_file": self.source_file or "",
                "metrics": dict(m), "defects": dict(d),
                "threshold": meta.get("threshold") or "N/D",
                "date": meta.get("data") or datetime.now().strftime("%d/%m/%Y"),
                "expectation": self.expectation_snapshot(),
                "current": True,
            })
        return datasets[-51:]

    def _draw_trend_chart(self, parent, datasets, metric_key, percent=False):
        canvas = tk.Canvas(parent, bg=PANEL, highlightthickness=0, height=235)
        canvas.pack(fill="both", expand=True, padx=10, pady=(2,10))
        canvas.update_idletasks()
        w=max(canvas.winfo_width(),620); h=235
        left,right,top,bottom=58,25,22,45
        plot_w=max(1,w-left-right); plot_h=max(1,h-top-bottom)
        vals=[float(ds.get("metrics",{}).get(metric_key)) for ds in datasets if ds.get("metrics",{}).get(metric_key) is not None]
        if not vals:
            canvas.create_text(w/2,h/2,text="N/D — dados insuficientes",fill=MUTED,font=("Segoe UI",9,"bold")); return
        lo,hi=(0.0,1.0) if percent else (0.0,max(max(vals),1.0)*1.12)
        for i in range(5):
            frac=i/4; y=top+plot_h*(1-frac)
            canvas.create_line(left,y,w-right,y,fill="#29404F")
            lab=lo+(hi-lo)*frac
            canvas.create_text(left-8,y,text=(f"{lab*100:.0f}%" if percent else f"{lab:.0f}"),fill=MUTED,anchor="e",font=("Segoe UI",7))
        step=plot_w/max(len(datasets)-1,1); pts=[]
        for i,ds in enumerate(datasets):
            if metric_key == "__DET_FP__":
                v=ds.get("defects",{}).get("fp")
            elif metric_key == "__DET_FN__":
                v=ds.get("defects",{}).get("fn")
            else:
                v=ds.get("metrics",{}).get(metric_key)
            x=left+(i*step if len(datasets)>1 else plot_w/2)
            if v is None: continue
            v=float(v); frac=(v-lo)/(hi-lo) if hi>lo else 0; y=top+plot_h*(1-max(0,min(1,frac)))
            pts.append((x,y)); canvas.create_oval(x-4,y-4,x+4,y+4,fill=ACCENT2,outline="")
            label=ds.get("threshold") or ds.get("name","?"); label=str(label)[:9]
            canvas.create_text(x,h-bottom+12,text=label,fill=MUTED,font=("Segoe UI",7))
            canvas.create_text(x,y-10,text=(f"{v*100:.1f}%" if percent else f"{v:.0f}"),fill=TEXT,font=("Segoe UI",7,"bold"))
        for a,b in zip(pts,pts[1:]): canvas.create_line(*a,*b,fill=ACCENT,width=2)
        canvas.create_text(left,h-bottom+29,text="Coletas / Threshold",fill="#6E8E9E",font=("Segoe UI",7))

    def _draw_comparison_overview(self, parent, datasets):
        """Visão única e legível das quatro métricas de qualidade por coleta."""
        canvas=tk.Canvas(parent,bg=PANEL,highlightthickness=0,height=285)
        canvas.pack(fill="both",expand=True,padx=12,pady=(2,12)); canvas.update_idletasks()
        w,h=max(canvas.winfo_width(),720),285; left,right,top,bottom=55,25,38,48; pw,ph=w-left-right,h-top-bottom
        series=[("Accuracy","Accuracy",ACCENT2),("Precision","Precision",ACCENT),("Recall","Recall / Sensibilidade",GOOD),("F1","F1",WARN)]
        for i in range(5):
            y=top+ph*(1-i/4); canvas.create_line(left,y,w-right,y,fill="#29404F")
            canvas.create_text(left-8,y,text=f"{i*25}%",fill=MUTED,anchor="e",font=("Segoe UI",7))
        for i,(name,_,color) in enumerate(series):
            canvas.create_text(left+i*112,17,text=f"● {name}",fill=color,anchor="w",font=("Segoe UI",8,"bold"))
        step=pw/max(len(datasets)-1,1)
        for name,key,color in series:
            pts=[]
            for i,ds in enumerate(datasets):
                v=ds.get("metrics",{}).get(key)
                if v is None: continue
                x=left+(i*step if len(datasets)>1 else pw/2); y=top+ph*(1-max(0,min(1,float(v))))
                pts.append((x,y,float(v)))
            for a,b in zip(pts,pts[1:]): canvas.create_line(a[0],a[1],b[0],b[1],fill=color,width=2)
            for j,(x,y,v) in enumerate(pts):
                canvas.create_oval(x-4,y-4,x+4,y+4,fill=color,outline="")
                if j in (0,len(pts)-1): canvas.create_text(x,y-11,text=f"{v*100:.1f}%",fill=TEXT,font=("Segoe UI",7,"bold"))
        for i,ds in enumerate(datasets):
            x=left+(i*step if len(datasets)>1 else pw/2); label=str(ds.get("threshold") or ds.get("name","?"))[:10]
            canvas.create_text(x,h-bottom+15,text=label,fill=MUTED,font=("Segoe UI",7))
        canvas.create_text(left,h-bottom+30,text="Coletas / threshold",fill="#6E8E9E",font=("Segoe UI",7))

    def _draw_error_comparison(self, parent, datasets):
        items=[]
        for ds in datasets:
            m=ds.get("metrics",{}); items.extend([
                (f"{str(ds.get('threshold') or ds.get('name','?'))[:7]} FP",m.get("FP",0),BAD),
                (f"{str(ds.get('threshold') or ds.get('name','?'))[:7]} FN",m.get("FN",0),WARN)])
        self.draw_bar_chart(parent,items,max_value=max([v for _,v,_ in items]+[1]))

    def view_comparacao(self):
        self.section_title("COMPARAÇÃO  /  EVOLUÇÃO",
            "Compare de 2 a 50 coletas e observe o efeito de threshold, configuração ou motor no desempenho.")
        datasets=self._comparison_datasets()
        if not datasets:
            f=self.card(self.content,"NENHUMA COLETA DISPONÍVEL","Realize uma coleta ou importe uma planilha para iniciar a comparação.")
            f.pack(fill="x",pady=4)
            tk.Label(f,text="O AIA preserva as coletas anteriores quando você usa 'Nova coleta'. A coleta atual também pode ser comparada.",bg=PANEL,fg=TEXT,wraplength=900,justify="left").pack(anchor="w",padx=16,pady=14)
            return
        toolbar=tk.Frame(self.content,bg=BG); toolbar.pack(fill="x",pady=(0,7))
        ttk.Button(toolbar,text="Comparar selecionadas",style="Accent.TButton",command=lambda:self._run_comparison(datasets)).pack(side="left",padx=3)
        ttk.Button(toolbar,text="Selecionar todas",command=self._select_all_compare).pack(side="left",padx=3)
        ttk.Button(toolbar,text="Remover da seleção",command=self._remove_selected_compare).pack(side="left",padx=3)
        ttk.Button(toolbar,text="Limpar seleção",command=self._clear_compare).pack(side="left",padx=3)
        tk.Label(toolbar,text="Selecione 2 a 50 coletas.",bg=BG,fg=MUTED,font=("Segoe UI",8)).pack(side="right",padx=5)
        box=self.card(self.content,"COLETAS DISPONÍVEIS",f"{len(datasets)} conjunto(s) disponível(is) nesta sessão")
        box.pack(fill="x",pady=4)
        cols=("coleta","data","threshold","n","recall","precision","fp","fn","dfp","dfn","origem")
        tree=ttk.Treeview(box,columns=cols,show="headings",selectmode="extended",height=min(10,max(5,len(datasets))))
        self.compare_tree=tree; self.compare_datasets=datasets
        heads=("Coleta","Data","Threshold","N","Recall","Precision","FP-C","FN-C","FP-D","FN-D","Origem"); widths=(125,90,80,55,80,85,55,55,55,55,230)
        for c,h,w in zip(cols,heads,widths): tree.heading(c,text=h); tree.column(c,width=w,anchor="center" if c!="origem" else "w")
        tree.pack(fill="x",padx=12,pady=(4,12))
        for i,ds in enumerate(datasets):
            m=ds.get("metrics",{}); origin="ATUAL" if ds.get("current") else (Path(ds.get("source_file","")).name if ds.get("source_file") else "Manual")
            dds=ds.get("defects",{}); tree.insert("","end",iid=str(i),values=(ds.get("name","Coleta"),ds.get("date","N/D"),ds.get("threshold","N/D"),m.get("N",0),pct(m.get("Recall / Sensibilidade")),pct(m.get("Precision")),m.get("FP",0),m.get("FN",0),dds.get("fp","N/D"),dds.get("fn","N/D"),origin))
        # Delete remove apenas a(s) coleta(s) da comparação atual. O histórico
        # e os dados arquivados continuam íntegros.
        tree.bind("<Delete>", lambda _event: self._remove_selected_compare())
        if 2<=len(datasets)<=10: tree.selection_set(*[str(i) for i in range(len(datasets))])
        result=self.card(self.content,"RESULTADO DA COMPARAÇÃO","A análise é calculada somente sobre as coletas selecionadas.")
        result.pack(fill="x",pady=(6,4)); self.compare_result=result
        es=self.expectation_summary()
        if es["configured"]:
            tk.Label(result,text=f"🎯 Meta ativa: {es['met']}/{es['configured']} critérios atendidos na coleta atual • {es['status']}",bg=PANEL,fg=self._expectation_status_color(es["status"]),font=("Segoe UI",9,"bold")).pack(anchor="w",padx=16,pady=(8,2))
        tk.Label(result,text="Selecione pelo menos 2 coletas e clique em 'Comparar selecionadas'.",bg=PANEL,fg=MUTED,font=("Segoe UI",9)).pack(anchor="w",padx=16,pady=14)

    def _select_all_compare(self):
        if hasattr(self,"compare_tree"): self.compare_tree.selection_set(self.compare_tree.get_children())

    def _remove_selected_compare(self):
        """Remove uma ou mais coletas apenas da seleção da comparação."""
        tree=getattr(self,"compare_tree",None)
        if tree is None:
            return
        selected=tree.selection()
        if not selected:
            self.status.set("Nenhuma coleta selecionada para remover da comparação.")
            return
        tree.selection_remove(*selected)
        tree.focus("")
        self.status.set(f"{len(selected)} coleta(s) removida(s) da seleção de comparação.")

    def _clear_compare(self):
        """Limpa a seleção de forma explícita em todas as versões do Tk."""
        tree=getattr(self,"compare_tree",None)
        if tree is None:
            return
        # selection_set(()) é mais confiável que selection_remove(tuple) em
        # alguns ambientes Tk/Windows e também cobre seleção inicial automática.
        tree.selection_set(())
        tree.focus("")
        self.status.set("Seleção de comparação limpa.")

    def _archive_current_from_view(self):
        item=self._archive_current_collection()
        if item:
            self.status.set(f"Coleta arquivada para comparação: {item['name']}"); self.show("comparacao")
        else: messagebox.showinfo("Comparação","Não há uma coleta ativa para arquivar.")

    def _run_comparison(self,datasets):
        ids=self.compare_tree.selection() if hasattr(self,"compare_tree") else ()
        if len(ids)<2:
            messagebox.showwarning("Comparação","Selecione pelo menos 2 coletas."); return
        selected=[datasets[int(i)] for i in ids]
        result=getattr(self,"compare_result",None)
        if result is None: return
        for w in result.winfo_children(): w.destroy()
        tk.Label(result,text=f"{len(selected)} COLETAS EM COMPARAÇÃO",bg=PANEL,fg=ACCENT2,font=("Segoe UI",11,"bold")).pack(anchor="w",padx=16,pady=(12,3))
        exp=self.expectation_snapshot()
        configured=any(exp.get(k) is not None for k in ("accuracy","precision","recall","f1","fp100","fn100"))
        if configured:
            met_count=0
            total_count=0
            for ds in selected:
                em=ds.get("expectation") or exp; mds=ds.get("metrics",{}); n=mds.get("N",0)
                actuals={"accuracy":mds.get("Accuracy"),"precision":mds.get("Precision"),"recall":mds.get("Recall / Sensibilidade"),"f1":mds.get("F1"),"fp100":None if not n else mds.get("FP",0)/n*100.0,"fn100":None if not n else mds.get("FN",0)/n*100.0}
                for k,op in [("accuracy",">="),("precision",">="),("recall",">="),("f1",">="),("fp100","<="),("fn100","<=")]:
                    t=em.get(k); a=actuals.get(k)
                    if t is not None:
                        total_count+=1
                        if a is not None and ((a>=t) if op==">=" else (a<=t)): met_count+=1
            tk.Label(result,text=f"🎯 ATENDIMENTO À META NAS COLETAS SELECIONADAS: {met_count}/{total_count} critérios-coleta atendidos",bg=PANEL,fg=self._expectation_status_color("META ATINGIDA" if total_count and met_count==total_count else ("EM CALIBRAÇÃO" if met_count else "FORA DA EXPECTATIVA")),font=("Segoe UI",9,"bold")).pack(anchor="w",padx=16,pady=(0,6))
        table=ttk.Treeview(result,columns=("indicador","base","final","delta"),show="headings",height=9)
        for c,h,w in [("indicador","Indicador",210),("base","Primeira selecionada",150),("final","Última selecionada",150),("delta","Variação",150)]: table.heading(c,text=h); table.column(c,width=w,anchor="center" if c!="indicador" else "w")
        table.pack(fill="x",padx=12,pady=8)
        first=selected[0].get("metrics",{}); last=selected[-1].get("metrics",{})
        fd=selected[0].get("defects",{}); ld=selected[-1].get("defects",{})
        rows=[("N",first.get("N",0),last.get("N",0),False),("TP classificação",first.get("TP",0),last.get("TP",0),False),("TN / VN classificação",first.get("VN/TN",0),last.get("VN/TN",0),False),("FP classificação",first.get("FP",0),last.get("FP",0),False),("FN classificação",first.get("FN",0),last.get("FN",0),False),("FP detecção",fd.get("fp","N/D"),ld.get("fp","N/D"),False),("FN detecção",fd.get("fn","N/D"),ld.get("fn","N/D"),False),("Accuracy",first.get("Accuracy"),last.get("Accuracy"),True),("Precision",first.get("Precision"),last.get("Precision"),True),("Recall",first.get("Recall / Sensibilidade"),last.get("Recall / Sensibilidade"),True),("F1",first.get("F1"),last.get("F1"),True)]
        for label,a,b,is_pct in rows:
            if is_pct: av=pct(a); bv=pct(b); dv="N/D" if a is None or b is None else f"{(b-a)*100:+.2f} p.p."
            else: av=str(a if a is not None else "N/D"); bv=str(b if b is not None else "N/D"); dv="N/D" if a is None or b is None else f"{b-a:+g}"
            table.insert("","end",values=(label,av,bv,dv))
        summary=tk.Frame(result,bg=PANEL); summary.pack(fill="x",padx=12,pady=(0,8))
        changes=[
            ("Δ RECALL", "N/D" if first.get("Recall / Sensibilidade") is None or last.get("Recall / Sensibilidade") is None else f"{(last['Recall / Sensibilidade']-first['Recall / Sensibilidade'])*100:+.1f} p.p.", GOOD if (last.get("Recall / Sensibilidade") or 0)>=(first.get("Recall / Sensibilidade") or 0) else BAD),
            ("Δ PRECISION", "N/D" if first.get("Precision") is None or last.get("Precision") is None else f"{(last['Precision']-first['Precision'])*100:+.1f} p.p.", GOOD if (last.get("Precision") or 0)>=(first.get("Precision") or 0) else BAD),
            ("Δ F1", "N/D" if first.get("F1") is None or last.get("F1") is None else f"{(last['F1']-first['F1'])*100:+.1f} p.p.", GOOD if (last.get("F1") or 0)>=(first.get("F1") or 0) else BAD),
            ("ERROS FP + FN", f"{first.get('FP',0)+first.get('FN',0)} → {last.get('FP',0)+last.get('FN',0)}", GOOD if last.get('FP',0)+last.get('FN',0)<=first.get('FP',0)+first.get('FN',0) else BAD),
        ]
        for i,(label,value,color) in enumerate(changes):
            c=self.metric_card(summary,label,value,color); c.grid(row=0,column=i,sticky="ew",padx=(0,4) if i<3 else 0); summary.columnconfigure(i,weight=1)
        charts=tk.Frame(result,bg=PANEL); charts.pack(fill="both",expand=True,padx=10,pady=(0,10))
        overview=self.chart_card(charts,"PAINEL DE EVOLUÇÃO", "Quatro métricas no mesmo eixo: compare o equilíbrio, não apenas um indicador isolado.")
        overview.grid(row=0,column=0,sticky="nsew",padx=(0,4),pady=4); self._draw_comparison_overview(overview,selected)
        errors=self.chart_card(charts,"CARGA DE ERROS DE CLASSIFICAÇÃO", "FP e FN por coleta; cores distintas tornam escapes e alarmes indevidos visíveis.")
        errors.grid(row=0,column=1,sticky="nsew",padx=(4,0),pady=4); self._draw_error_comparison(errors,selected)
        charts.columnconfigure(0,weight=3); charts.columnconfigure(1,weight=2)
        read=self.card(result,"LEITURA COMPARATIVA"); read.pack(fill="x",padx=10,pady=(0,10))
        rv=[x.get("metrics",{}).get("Recall / Sensibilidade") for x in selected]; pv=[x.get("metrics",{}).get("Precision") for x in selected]
        lines=[]
        if all(v is not None for v in rv): lines.append(f"Recall: {(rv[-1]-rv[0])*100:+.2f} p.p. entre a primeira e a última coleta.")
        if all(v is not None for v in pv): lines.append(f"Precision: {(pv[-1]-pv[0])*100:+.2f} p.p. entre a primeira e a última coleta.")
        lines.append(f"Falsos positivos: {first.get('FP',0)} → {last.get('FP',0)}; falsos negativos: {first.get('FN',0)} → {last.get('FN',0)}.")
        lines.append("A leitura indica tendência nos dados; a escolha de configuração deve considerar contexto, amostra e risco operacional.")
        for line in lines: tk.Label(read,text="• "+line,bg=PANEL,fg=TEXT,wraplength=1050,justify="left",font=("Segoe UI",9)).pack(anchor="w",padx=14,pady=4)

    # ---------- PROJECT / EXPORT ----------
    def new_collection(self, confirm=True):
        """Inicia uma nova coleta totalmente vazia, separada das inspeções atuais.

        Diferente de "+ Nova inspeção", esta ação encerra o conjunto atual de
        dados em memória e prepara a tela para outro cliente/planta/projeto.
        Nenhum arquivo já salvo no computador é apagado.
        """
        if confirm:
            ok = messagebox.askyesno(
                "Nova coleta — iniciar do zero",
                "Isso vai limpar da tela todas as inspeções, histórico e informações "
                "da coleta atual.\n\n"
                "Os arquivos já salvos no computador NÃO serão apagados.\n\n"
                "Deseja iniciar uma nova coleta do zero?",
                icon="warning"
            )
            if not ok:
                return

        self._archive_current_collection()
        self.rows = []
        self.history = []
        self.source_file = ""
        for v in self.meta.values():
            v.set("")
        # Nova coleta começa realmente vazia: a data deve representar o dia
        # em que a nova inspeção/coleta foi realizada, não o dia em que o
        # software foi aberto. O campo permanece editável na tela Coleta.
        self.meta["data"].set("")
        self.meta["gabarito"].set("Não determinado")
        self.status.set("Nova coleta iniciada — nenhum dado registrado.")
        self.refresh_all()

    def new_test(self,confirm=True):
        # Compatibilidade com chamadas antigas: agora o conceito de novo teste
        # também inicia uma coleta realmente vazia.
        self.new_collection(confirm=confirm)

    def refresh_all(self):
        if hasattr(self,"tree"): self.refresh_table()
        self.update_side()
        if self.current_view:self.show(self.current_view)

    def update_side(self):
        g=self.meta["gabarito"].get() or "Não determinado"
        src="Excel histórico" if self.source_file else "Entrada manual"
        classified=sum(1 for r in self.rows if r.get("ref") in ("Defeito confirmado","OK confirmado"))
        self.side_status.configure(text=f"Fonte: {src}\nReferência: {g}\nInspeções: {len(self.rows)}\nClassificadas: {classified}")

    def save_project(self):
        path=filedialog.asksaveasfilename(title="Salvar projeto",defaultextension=".json",
                                          filetypes=[("Projeto JSON","*.json")])
        if not path:return
        data={"app":APP,"version":VERSION,"saved_at":datetime.now().isoformat(),
              "metadata":{k:v.get() for k,v in self.meta.items()},
              "rows":[normalize_row(r) for r in self.rows],
              "history":self.history,"source_file":self.source_file,
              "comparison_archive":self.comparison_archive,"calibration_history":self.calibration_history,
              "calibration_project":self.calibration_project,"expectation":self.expectation_snapshot()}
        Path(path).write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding="utf-8")
        self.status.set(f"Projeto salvo: {Path(path).name}")

    def _project_summary(self):
        return self.metrics(), (self.history[-1] if self.history else {"TP":0,"VN/TN":0,"FP":0,"FN":0,"N":0})

    def export_executive_report(self):
        """Cria um PDF compacto para decisão, sem expor a operação interna do AIA."""
        path=filedialog.asksaveasfilename(title="Gerar relatório executivo",defaultextension=".pdf",filetypes=[("PDF","*.pdf")])
        if not path: return
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import cm
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
            m=self.display_metrics(); d=self.defect_counts(); best=self._best_configuration(self.calibration_history)
            styles=getSampleStyleSheet(); title_style=ParagraphStyle("aia",parent=styles["Title"],textColor=colors.HexColor("#087EA4"),fontName="Helvetica-Bold",fontSize=20,spaceAfter=6)
            sub=ParagraphStyle("sub",parent=styles["Normal"],textColor=colors.HexColor("#406070"),fontSize=9,spaceAfter=12)
            doc=SimpleDocTemplate(path,pagesize=A4,rightMargin=1.6*cm,leftMargin=1.6*cm,topMargin=1.5*cm,bottomMargin=1.5*cm)
            story=[Paragraph("AUTAZA INSPECTION ANALYTICS",title_style),Paragraph(f"RELATÓRIO EXECUTIVO - V{VERSION} - Emitido em {datetime.now().strftime('%d/%m/%Y %H:%M')}",sub)]
            project=self.calibration_project
            info=[["Projeto",project.get("name") or self.meta["projeto"].get() or "N/D"],["Código",project.get("id") or "N/D"],["Cliente / Planta",f"{self.meta['cliente'].get() or 'N/D'} / {self.meta['local'].get() or 'N/D'}"],["Responsável",project.get("responsible") or self.meta["tecnico"].get() or "N/D"],["Objetivo",project.get("objective") or "Avaliar desempenho da inspeção"]]
            t=Table(info,colWidths=[4.1*cm,12.0*cm]); t.setStyle(TableStyle([("BACKGROUND",(0,0),(0,-1),colors.HexColor("#EAF5F8")),("TEXTCOLOR",(0,0),(0,-1),colors.HexColor("#07526B")),("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),("GRID",(0,0),(-1,-1),0.25,colors.HexColor("#B8D3DC")),("VALIGN",(0,0),(-1,-1),"TOP"),("PADDING",(0,0),(-1,-1),7)])); story += [t,Spacer(1,14)]
            story.append(Paragraph("Resultado da coleta atual",styles["Heading2"]))
            metrics=[["N",m.get("N",0),"Accuracy",pct(m.get("Accuracy"))],["Recall",pct(m.get("Recall / Sensibilidade")),"Precision",pct(m.get("Precision"))],["F1",pct(m.get("F1")),"FP / FN",f"{m.get('FP',0)} / {m.get('FN',0)}"],["Defeitos reais",d.get("real",0),"Detecções",d.get("det",0)]]
            t=Table(metrics,colWidths=[3.4*cm,3.2*cm,3.4*cm,6.1*cm]); t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#0E202D")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),0.25,colors.HexColor("#B8D3DC")),("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),("ALIGN",(1,0),(-1,-1),"CENTER"),("PADDING",(0,0),(-1,-1),7)])); story += [t,Spacer(1,14)]
            story.append(Paragraph("Recomendação de calibração",styles["Heading2"]))
            if best:
                bm=best["metrics"]; recommendation=f"A configuração recomendada é <b>{best.get('name')}</b>, com threshold <b>{best.get('threshold','N/D')}</b>, F1 de <b>{pct(bm.get('F1'))}</b>, Recall de <b>{pct(bm.get('Recall / Sensibilidade'))}</b> e Precision de <b>{pct(bm.get('Precision'))}</b>. A recomendação ordena o atendimento às metas e, em empate, F1, Recall e menor FP."
            else: recommendation="Ainda não há coleta classificada suficiente no histórico para recomendar uma configuração."
            story += [Paragraph(recommendation,styles["BodyText"]),Spacer(1,12)]
            es=self.expectation_summary(); story.append(Paragraph("Critério de entrega",styles["Heading2"])); story.append(Paragraph(f"Status atual: <b>{es.get('status','N/D')}</b>. Critérios atendidos: <b>{es.get('met',0)}/{es.get('configured',0)}</b>.",styles["BodyText"]))
            story += [Spacer(1,12),Paragraph("Rastreabilidade",styles["Heading2"]),Paragraph(f"Histórico do projeto: {len(self.calibration_history)} coleta(s). A coleta atual contém {len(self.rows)} inspeção(ões). Fonte: {Path(self.source_file).name if self.source_file else 'entrada manual'}.",styles["BodyText"])]
            doc.build(story); self.status.set(f"Relatório executivo gerado: {Path(path).name}"); messagebox.showinfo("Relatório executivo","PDF gerado com sucesso.")
        except Exception as e:
            messagebox.showerror("Relatório executivo",f"Não foi possível gerar o PDF:\n{e}")

    def export_csv(self):
        path=filedialog.asksaveasfilename(title="Exportar CSV estruturado",defaultextension=".csv",
                                          filetypes=[("CSV","*.csv")])
        if not path:return
        m,hist=self._project_summary();d=self.defect_counts()
        with open(path,"w",newline="",encoding="utf-8-sig") as f:
            w=csv.writer(f,delimiter=";")
            w.writerow(["AUTAZA INSPECTION ANALYTICS",f"V{VERSION}"])
            w.writerow(["Data de exportação",datetime.now().strftime("%d/%m/%Y %H:%M:%S")]);w.writerow([])
            w.writerow(["IDENTIFICAÇÃO DO TESTE"])
            for key,label in [("cliente","Cliente"),("local","Local / Planta"),("projeto","Projeto"),
                              ("linha","Linha / Estação"),("tecnico","Técnico"),("data","Data"),
                              ("threshold","Threshold"),("gabarito","Método de referência")]:
                w.writerow([label,self.meta[key].get()])
            w.writerow(["Fonte do teste","Excel histórico" if self.source_file else "Entrada manual"])
            w.writerow(["Arquivo de origem",self.source_file or ""]);w.writerow([])
            w.writerow(["EXPECTATIVA / CRITÉRIO DE ENTREGA"])
            es=self.expectation_snapshot()
            for key,label in [("accuracy","Meta Accuracy"),("precision","Meta Precision"),("recall","Meta Recall"),("f1","Meta F1"),("fp100","Meta FP/100"),("fn100","Meta FN/100"),("source","Origem do critério"),("responsible","Responsável"),("date","Data do critério"),("notes","Observações")]:
                w.writerow([label,es.get(key)])
            w.writerow([])
            w.writerow(["CLASSIFICAÇÃO DAS INSPEÇÕES"])
            w.writerow(["Indicador","Valor"])
            for k in ("TP","VN/TN","FP","FN","N"):w.writerow([k,m.get(k,0)])
            for k in ("Accuracy","Precision","Recall / Sensibilidade","Specificity","FPR","FNR","F1","MCC"):
                v=m.get(k);w.writerow([k,"N/D" if v is None else v])
            w.writerow([])
            w.writerow(["DETECÇÃO DE OCORRÊNCIAS"])
            w.writerow(["Indicador","Valor"])
            for k,label in (("real","Defeitos reais"),("det","Detecções"),("tp","TP de detecção"),("fp","FP de detecção"),("fn","FN de detecção")):w.writerow([label,d[k]])
            w.writerow(["Análise por tipo","Detalhada" if d["detailed"] else "Coleta rápida — sem tipo"]);w.writerow([])
            w.writerow(["INSPEÇÕES"])
            w.writerow(["Inspeção","Identificação","Vision AI","Referência","Real","Detectado",
                        "TP-C","VN-C","FP-C","FN-C","TP-D","FP-D","FN-D","Tipos detalhados","Observação","Origem"])
            for i,r0 in enumerate(self.rows,1):
                r=normalize_row(r0); dd=self.detection_for_row(r)
                dv=(dd["tp"],dd["fp"],dd["fn"]) if dd["available"] else ("N/D","N/D","N/D")
                w.writerow([i,r["id"],r["vision"],r["ref"],r["real"],r["det"],r["tp"],r["vn"],
                            r["fp"],r["fn"],dv[0],dv[1],dv[2],len(r["defects"]),r["obs"],r["source"]])
            w.writerow([])
            w.writerow(["DEFEITOS DETALHADOS"])
            w.writerow(["Inspeção","Identificação","Tipo","Real","Detectado","TP por contagem","FP","FN"])
            for i,r0 in enumerate(self.rows,1):
                r=normalize_row(r0)
                for dd in r["defects"]:
                    tt=min(dd["real"],dd["det"]);ff=max(dd["det"]-dd["real"],0);nn=max(dd["real"]-dd["det"],0)
                    w.writerow([i,r["id"],dd["type"],dd["real"],dd["det"],tt,ff,nn])
            w.writerow([])
            w.writerow(["AUDITORIA HISTÓRICO × MOTOR NOVO"])
            w.writerow(["Indicador","Histórico","Motor novo","Diferença"])
            for k in ("TP","VN/TN","FP","FN","N"):
                hv=hist.get(k,0);mv=m.get(k,0);w.writerow([k,hv,mv,mv-hv])
        self.status.set(f"CSV estruturado exportado: {Path(path).name}")
        messagebox.showinfo("Exportação concluída",
                            "CSV estruturado criado com classificação, ocorrências, inspeções e detalhes.")

    def export_excel(self):
        if openpyxl is None:
            messagebox.showerror("Excel","OpenPyXL não está instalado.\nExecute: py -3 -m pip install openpyxl");return
        path=filedialog.asksaveasfilename(title="Exportar relatório completo",defaultextension=".xlsx",
                                          filetypes=[("Excel","*.xlsx")])
        if not path:return
        try:
            from openpyxl import Workbook
            from openpyxl.styles import Font,PatternFill,Border,Side,Alignment
            wb=Workbook();ws=wb.active;ws.title="RESUMO"
            dark="182532";blue="2F8CFF";border="6B7F8D";green="45D39A";yellow="F4C15D";red="FF6B78"
            thin=Side(style="thin",color=border)
            def title(sh,text,cols=6):
                sh["A1"]=text;sh["A1"].font=Font(size=18,bold=True,color="FFFFFF")
                sh["A1"].fill=PatternFill("solid",fgColor=dark);sh.merge_cells(start_row=1,start_column=1,end_row=1,end_column=cols)
            def hdr(sh,row,labels):
                for c,v in enumerate(labels,1):
                    x=sh.cell(row,c,v);x.font=Font(bold=True,color="FFFFFF")
                    x.fill=PatternFill("solid",fgColor=dark);x.border=Border(bottom=thin)
            def autosize(sh):
                # Ignora MergedCell: títulos mesclados não possuem column_letter.
                # Mede somente células reais e aplica largura por coluna.
                from openpyxl.cell.cell import Cell
                widths = {}
                for row in sh.iter_rows():
                    for x in row:
                        if not isinstance(x, Cell):
                            continue
                        value = "" if x.value is None else str(x.value)
                        widths[x.column] = max(widths.get(x.column, 10), min(55, len(value) + 2))
                for idx, width in widths.items():
                    from openpyxl.utils import get_column_letter
                    sh.column_dimensions[get_column_letter(idx)].width = max(10, min(55, width))
            m,hist=self._project_summary();d=self.defect_counts()
            title(ws,"AUTAZA INSPECTION ANALYTICS — RELATÓRIO",6)
            info=[("Versão",VERSION),("Data de exportação",datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
                  ("Fonte","Excel histórico" if self.source_file else "Entrada manual"),
                  ("Cliente",self.meta["cliente"].get()),("Local / Planta",self.meta["local"].get()),
                  ("Projeto",self.meta["projeto"].get()),("Linha / Estação",self.meta["linha"].get()),
                  ("Técnico",self.meta["tecnico"].get()),("Data",self.meta["data"].get()),
                  ("Threshold",self.meta["threshold"].get()),("Método de referência",self.meta["gabarito"].get())]
            for i,(a,b) in enumerate(info,3):ws.cell(i,1,a);ws.cell(i,2,b)
            hdr(ws,16,["Indicador","Valor","Interpretação"])
            interp={"TP":"Inspeções reais NOK classificadas como NOK","VN/TN":"Inspeções reais OK classificadas como OK",
                    "FP":"Inspeções reais OK classificadas como NOK","FN":"Inspeções reais NOK classificadas como OK"}
            vals=[("TP",m["TP"]),("VN/TN",m["VN/TN"]),("FP",m["FP"]),("FN",m["FN"]),
                  ("N",m["N"]),("Accuracy",m["Accuracy"]),("Precision",m["Precision"]),
                  ("Recall / Sensibilidade",m["Recall / Sensibilidade"]),("Specificity",m["Specificity"]),
                  ("F1",m["F1"])]
            for i,(k,v) in enumerate(vals,17):
                ws.cell(i,1,k);ws.cell(i,2,v);ws.cell(i,3,interp.get(k,""))
                if isinstance(v,float):ws.cell(i,2).number_format="0.00%"
            autosize(ws)

            we=wb.create_sheet("EXPECTATIVA");title(we,"EXPECTATIVA / CRITÉRIO DE ENTREGA",6)
            hdr(we,3,["Indicador","Meta","Resultado atual","Critério","Status","Origem"])
            for i,(label,key,op,a,t,status) in enumerate(self._expectation_rows(),4):
                we.cell(i,1,label);we.cell(i,2,"N/D" if t is None else t);we.cell(i,3,"N/D" if a is None else a);we.cell(i,4,op);we.cell(i,5,status);we.cell(i,6,self.expectation["source"].get())
                if key in ("accuracy","precision","recall","f1"):
                    if isinstance(t,float):we.cell(i,2).number_format="0.00%"
                    if isinstance(a,float):we.cell(i,3).number_format="0.00%"
            autosize(we)

            wi=wb.create_sheet("INSPEÇÕES");title(wi,"INSPEÇÕES",16)
            hdr(wi,3,["Inspeção","Identificação","Vision AI","Referência","Real","Detectado","TP-C","VN-C","FP-C","FN-C","TP-D","FP-D","FN-D","Tipos detalhados","Observação","Origem"])
            for i,r0 in enumerate(self.rows,4):
                r=normalize_row(r0); dd=self.detection_for_row(r)
                dv=(dd["tp"],dd["fp"],dd["fn"]) if dd["available"] else ("N/D","N/D","N/D")
                vals=[i-3,r["id"],r["vision"],r["ref"],r["real"],r["det"],r["tp"],r["vn"],r["fp"],r["fn"],
                      dv[0],dv[1],dv[2],len(r["defects"]),r["obs"],r["source"]]
                for c,v in enumerate(vals,1):wi.cell(i,c,v)
            wi.freeze_panes="A4";autosize(wi)

            wm=wb.create_sheet("MÉTRICAS");title(wm,"MÉTRICAS DE CLASSIFICAÇÃO",3)
            hdr(wm,3,["Métrica","Valor","Definição"])
            defs={"Accuracy":"Classificações corretas / total","Precision":"TP/(TP+FP)",
                  "Recall / Sensibilidade":"TP/(TP+FN)","Specificity":"TN/(TN+FP)",
                  "FPR":"FP/(FP+TN)","FNR":"FN/(FN+TP)","F1":"Média harmônica de Precision e Recall",
                  "Balanced Accuracy":"(Recall+Specificity)/2","MCC":"Correlação da matriz de confusão"}
            for i,k in enumerate(("Accuracy","Precision","Recall / Sensibilidade","Specificity","FPR","FNR","F1","Balanced Accuracy","MCC"),4):
                v=m[k];wm.cell(i,1,k);wm.cell(i,2,v if v is not None else "N/D");wm.cell(i,3,defs.get(k,""))
                if isinstance(v,float) and k!="MCC":wm.cell(i,2).number_format="0.00%"
            autosize(wm)

            wa=wb.create_sheet("AUDITORIA");title(wa,"AUDITORIA — HISTÓRICO × MOTOR NOVO",4)
            hdr(wa,3,["Indicador","Histórico","Motor novo","Diferença"])
            for i,k in enumerate(("TP","VN/TN","FP","FN","N"),4):
                hv=hist.get(k,0);mv=m.get(k,0);wa.cell(i,1,k);wa.cell(i,2,hv);wa.cell(i,3,mv);wa.cell(i,4,mv-hv)
            autosize(wa)

            wd=wb.create_sheet("POR DEFEITO");title(wd,"ANÁLISE POR TIPO DE DEFEITO",9)
            hdr(wd,3,["Defeito","Real","Detectado","TP","FP","FN","Recall","Precision","Inspeções"])
            for i,(name,v) in enumerate(sorted(d["by"].items()),4):
                mm=classification_metrics(v["tp"],0,v["fp"],v["fn"])
                vals=[name,v["real"],v["det"],v["tp"],v["fp"],v["fn"],
                      mm["Recall / Sensibilidade"],mm["Precision"],v["inspecoes"]]
                for c,x in enumerate(vals,1):wd.cell(i,c,x)
                wd.cell(i,7).number_format="0.00%";wd.cell(i,8).number_format="0.00%"
            autosize(wd)

            wdo=wb.create_sheet("DETECCAO");title(wdo,"DETECÇÃO DE OCORRÊNCIAS",6)
            hdr(wdo,3,["Indicador","Valor","Observação"])
            vals=[("Defeitos reais",d["real"],"Contagem registrada"),
                  ("Detecções",d["det"],"Contagem registrada"),
                  ("TP de detecção",d["tp"],"min(real,detectado) quando não há correspondência individual"),
                  ("FP de detecção",d["fp"],"max(detectado-real,0) — estimativa por contagem"),
                  ("FN de detecção",d["fn"],"max(real-detectado,0) — estimativa por contagem"),
                  ("Classificação por tipo","Sim" if d["detailed"] else "Não","Opcional")]
            for i,row in enumerate(vals,4):
                for c,v in enumerate(row,1):wdo.cell(i,c,v)
            autosize(wdo)

            wg=wb.create_sheet("GLOSSÁRIO");title(wg,"GLOSSÁRIO",2)
            hdr(wg,3,["Sigla / Métrica","Definição"])
            for i,(term,desc) in enumerate(GLOSSARY.items(),4):
                wg.cell(i,1,term);wg.cell(i,2,desc);wg.cell(i,2).alignment=Alignment(wrap_text=True,vertical="top")
            wg.column_dimensions["A"].width=42;wg.column_dimensions["B"].width=110
            wb.save(path)
            self.status.set(f"Relatório Excel exportado: {Path(path).name}")
            messagebox.showinfo("Exportação concluída",
                "Relatório Excel criado:\nRESUMO • INSPEÇÕES • MÉTRICAS • AUDITORIA • POR DEFEITO • DETECCAO • GLOSSÁRIO")
        except Exception as e:
            messagebox.showerror("Erro na exportação Excel",str(e))

    def glossary(self):
        win=tk.Toplevel(self);win.title("Glossário");win.geometry("920x720");win.configure(bg=BG)
        text=tk.Text(win,bg=PANEL3,fg=TEXT,insertbackground=TEXT,relief="flat",wrap="word")
        text.pack(fill="both",expand=True,padx=14,pady=14)
        for term,desc in GLOSSARY.items():
            text.insert("end",term+"\n","term");text.insert("end",desc+"\n\n")
        text.tag_configure("term",foreground=ACCENT2,font=("Segoe UI",11,"bold"))
        text.configure(state="disabled")

if __name__=="__main__":
    App().mainloop()
