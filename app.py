# -*- coding: utf-8 -*-
"""Etiquetas Freedom — imprime etiquetas de volume em qualquer impressora do Windows."""
import datetime
import json
import os
import queue
import shutil
import socket
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk

import historico
from controle import verifica
import arquivo_zip
import danfe
import impressora_virtual
import ocr
from comunicado import desenha_comunicado
from etiqueta_core import Planilha, desenha_etiqueta, ler_tabela, pasta_dados, pasta_exe, pasta_saida, recurso


def carrega_fontes():
    """Registra a Roboto embutida (só para este programa). Sem ela, cai na Segoe UI."""
    if sys.platform != "win32":
        return "Segoe UI"
    try:
        import ctypes
        for nome in ("Regular", "Medium", "Bold", "Black"):
            ctypes.windll.gdi32.AddFontResourceExW(recurso("fonts", f"Roboto-{nome}.ttf"), 0x10, 0)
        return "Roboto"
    except Exception:
        return "Segoe UI"


FONTE = carrega_fontes()

PASTA = pasta_exe()
MINIMIZADO = "--minimizado" in sys.argv        # aberto com o Windows: fica na barra até chegar um DANFE
MODO_DANFE = "--danfe" in sys.argv          # aberto pela tarefa agendada quando algo é impresso
PORTA_LOCAL = 47653                        # garante uma única cópia do programa aberta
CONFIG = os.path.join(pasta_dados(), "config.json")

# paleta Freedom
BG, CARD, TX, TX2, TX3 = "#eef2f7", "#ffffff", "#0f1b3d", "#4b5875", "#7b879e"
ROSA, LINHA = "#e0457f", "#e1e7ef"


# ---------------------------------------------------------------- arquivo de clientes (vigiado)
EXT_CLIENTES = (".csv", ".xlsx", ".xlsm")
INTERVALO_VIGIA = 15_000  # ms


def arquivo_clientes(origem):
    """O arquivo a ler: o próprio arquivo, ou (se for pasta) o mais recente com 'client' no nome."""
    if not origem:
        return None
    if os.path.isfile(origem):
        return origem
    if os.path.isdir(origem):
        c = [os.path.join(origem, f) for f in os.listdir(origem)
             if f.lower().endswith(EXT_CLIENTES) and "client" in f.lower() and not f.startswith("~$")]
        return max(c, key=os.path.getmtime) if c else None
    return None


def assinatura(caminho):
    st = os.stat(caminho)
    return f"{os.path.abspath(caminho)}|{st.st_mtime_ns}|{st.st_size}"


# ---------------------------------------------------------------- impressão
VIRTUAIS = ("pdf", "xps", "onenote", "fax", "anydesk", "teamviewer", "document writer", "snagit",
            "print to file", "enviar para", "send to", "webex", "zoom", "virtual")
PORTAS_VIRTUAIS = ("portprompt", "nul", "file", "shrfax", "xpsport", "microsoft.")
ST_PAUSADA, ST_EXCLUINDO, ST_OFFLINE, ST_INDISPONIVEL, ST_SERVIDOR_OFFLINE = 0x1, 0x4, 0x80, 0x1000, 0x2000000
ATTR_OFFLINE = 0x400


def impressora_ativa(nome, porta="", status=0, atributos=0):
    """True para impressora real e pronta: não é virtual (PDF, XPS, Fax...), nem offline/pausada."""
    n, p = (nome or "").lower(), (porta or "").lower()
    if n == impressora_virtual.NOME.lower() or p.endswith((".pdf", ".xps")):
        return False
    if any(v in n for v in VIRTUAIS) or p.startswith(PORTAS_VIRTUAIS):
        return False
    if status & (ST_PAUSADA | ST_EXCLUINDO | ST_OFFLINE | ST_INDISPONIVEL | ST_SERVIDOR_OFFLINE):
        return False
    return not (atributos & ATTR_OFFLINE)


def listar_impressoras(todas=False):
    """Impressoras do Windows. Por padrão só as ativas (reais e online)."""
    try:
        import win32print
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        nomes = []
        for p in win32print.EnumPrinters(flags, None, 2):
            nome = p["pPrinterName"]
            if todas or impressora_ativa(nome, p.get("pPortName", ""), p.get("Status", 0), p.get("Attributes", 0)):
                nomes.append(nome)
        return sorted(set(nomes), key=str.lower)
    except Exception:
        return []


def impressora_padrao():
    try:
        import win32print
        return win32print.GetDefaultPrinter()
    except Exception:
        return ""


MODO_ZPL = [False]   # ligado em ⚙ Configurações: manda ZPL direto (só Zebra/compatíveis), bem mais rápido


def impressora_esta_ativa(nome, estrito=True):
    """Confere só a impressora escolhida (rápido), sem listar todas as do Windows."""
    if sys.platform != "win32":
        return True
    try:
        import win32print
        h = win32print.OpenPrinter(nome)
        try:
            info = win32print.GetPrinter(h, 2)
        finally:
            win32print.ClosePrinter(h)
        return (not estrito) or impressora_ativa(nome, info.get("pPortName", ""), info.get("Status", 0),
                                                 info.get("Attributes", 0))
    except Exception:
        return False


def imprimir_zpl(nome_impressora, imagens):
    import win32print
    import zpl
    h = win32print.OpenPrinter(nome_impressora)
    try:
        win32print.StartDocPrinter(h, 1, ("Etiquetas Freedom", None, "RAW"))
        try:
            win32print.StartPagePrinter(h)
            win32print.WritePrinter(h, zpl.trabalho_zpl(imagens))
            win32print.EndPagePrinter(h)
        finally:
            win32print.EndDocPrinter(h)
    finally:
        win32print.ClosePrinter(h)


def imprimir(nome_impressora, imagens, w_mm, h_mm):
    """Imprime. Normal: pelo driver do Windows com bitmap cinza (3x menor que colorido). Modo ZPL: direto na Zebra."""
    if MODO_ZPL[0]:
        return imprimir_zpl(nome_impressora, imagens)
    import win32ui
    from PIL import ImageWin

    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(nome_impressora)
    dc.StartDoc("Etiquetas Freedom")
    try:
        dpi_x, dpi_y = dc.GetDeviceCaps(88), dc.GetDeviceCaps(90)  # LOGPIXELSX/Y
        larg, alt = int(w_mm / 25.4 * dpi_x), int(h_mm / 25.4 * dpi_y)
        for img in imagens:
            dc.StartPage()
            ImageWin.Dib(img.convert("L")).draw(dc.GetHandleOutput(), (0, 0, larg, alt))
            dc.EndPage()
        dc.EndDoc()
    except Exception:
        dc.AbortDoc()
        raise
    finally:
        dc.DeleteDC()


# ---------------------------------------------------------------- app
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()
        self.title("Etiquetas Freedom")
        self._icone()
        self.configure(bg=BG)
        self.geometry("1000x720")
        self.minsize(940, 680)
        self.cfg = self._ler_cfg()
        self.bloqueado = False
        self._pronto = False
        self._salva_id = None
        self._fila = queue.SimpleQueue()

        # controle remoto: consulta rápida antes de abrir a tela
        bloq, msg = verifica()
        if bloq:
            self.bloqueado = True
            messagebox.showerror("Etiquetas Freedom", msg)
            self.destroy()
            return
        self._aviso_controle(msg)

        self.plan = None
        self.cliente = None
        self.vol_prev = 1
        self._foto = None
        self.fila_danfe, self.danfe_atual, self._danfe_sigs = [], None, {}
        self._msgs = queue.SimpleQueue()
        self._res_danfe, self._lendo_danfe, self._aviso_ativo = queue.SimpleQueue(), 0, False
        self._res_imp, self._res_print, self._res_excel = queue.SimpleQueue(), queue.SimpleQueue(), queue.SimpleQueue()
        self._imprimindo, self._excel_rodando, self._excel_id, self._tam_prev = False, False, None, {}
        MODO_ZPL[0] = bool(self.cfg.get("modo_zpl"))

        self._estilo()
        self._monta()
        self._carrega_base()
        if self.plan is None:
            self.destroy()
            self.bloqueado = True
            return
        self.bind("<F5>", lambda e: self.imprimir_todas())
        self.bind("<Control-u>", lambda e: self.atualiza_clientes())

        # salva a configuração sempre que algo mudar e ao fechar
        for v in (self.v_imp, self.v_rem, self.v_tr, self.v_w, self.v_h, self.v_dpi):
            v.trace_add("write", lambda *a: self._agenda_salva())
        self.protocol("WM_DELETE_WINDOW", self._fechar)
        self._pronto = True
        self._adota_origem_clientes()
        self._salva_cfg()
        if MINIMIZADO and not MODO_DANFE:
            self.iconify()                       # aberto com o Windows: espera na barra de tarefas
        elif not MODO_DANFE:                     # aberto pela tarefa do ERP: só aparece quando chegar um DANFE
            self.deiconify()
        self._lendo = False
        self._fila_cli = queue.SimpleQueue()
        self.after(1000, self._vigia_clientes)
        self.after(5 * 60 * 1000, self._ciclo_controle)
        self.after(400, self._vigia_danfe)
        self.after(4000, self._ciclo_zip)
        threading.Thread(target=ocr.aquece, daemon=True).start()     # deixa o leitor de imagem pronto
        if self.cfg.get("excel_pendente"):
            self._excel_id = self.after(12000, self._excel_bg)
        if MODO_DANFE:
            self.after(6000, self._sai_se_ocioso)

    # ---------- controle remoto
    def _aviso_controle(self, msg):
        if msg and msg != self.cfg.get("ultima_mensagem"):
            messagebox.showinfo("Aviso", msg)
            self.cfg["ultima_mensagem"] = msg
            self._salva_cfg()

    def _ciclo_controle(self):
        """A cada 5 min consulta o controle em segundo plano (sem travar a tela)."""
        threading.Thread(target=lambda: self._fila.put(verifica()), daemon=True).start()
        self.after(1500, self._le_fila)
        self.after(5 * 60 * 1000, self._ciclo_controle)

    def _le_fila(self):
        try:
            bloq, msg = self._fila.get_nowait()
        except queue.Empty:
            self.after(1000, self._le_fila)
            return
        if bloq:
            self.bloqueado = True
            self._salva_cfg()
            messagebox.showerror("Etiquetas Freedom", msg)
            self.destroy()
        else:
            self._aviso_controle(msg)

    # ---------- config
    def _ler_cfg(self):
        try:
            with open(CONFIG, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _agenda_salva(self):
        if not self._pronto:
            return
        if self._salva_id:
            self.after_cancel(self._salva_id)
        self._salva_id = self.after(400, self._salva_cfg)

    def _salva_cfg(self):
        self._salva_id = None
        try:
            self.cfg.update(
                impressora=self.v_imp.get(), remetente=self.v_rem.get(),
                transportadora=self.v_tr.get(), w=self.v_w.get(), h=self.v_h.get(), dpi=self.v_dpi.get(),
                todas_impressoras=self.v_todas.get(),
            )
        except AttributeError:
            pass
        try:
            os.makedirs(os.path.dirname(CONFIG), exist_ok=True)
            tmp = CONFIG + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.cfg, f, ensure_ascii=False, indent=1)
            os.replace(tmp, CONFIG)
        except Exception:
            pass

    def _fechar(self):
        if self.cfg.get("excel_pendente"):               # histórico pequeno: atualiza o Excel antes de sair
            try:
                if os.path.getsize(historico.caminho_csv()) < 400_000:
                    ok, _ = historico.atualizar_excel(self.xlsx_hist())
                    self.cfg["excel_pendente"] = not ok
            except OSError:
                pass
        self._salva_cfg()
        self.destroy()

    # ---------- visual
    def _icone(self):
        """Ícone da janela e da barra de tarefas: logo rosa da Freedom (em vez da pena padrão)."""
        try:
            if sys.platform == "win32":
                self.iconbitmap(default=recurso("assets", "icone.ico"))
            else:
                self._img_icone = ImageTk.PhotoImage(Image.open(recurso("assets", "icone.png")))
                self.iconphoto(True, self._img_icone)
        except Exception:
            pass

    def _estilo(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".", background=BG, foreground=TX, font=(FONTE, 10))
        s.configure("Card.TFrame", background=CARD)
        s.configure("TLabel", background=CARD, foreground=TX)
        s.configure("Rot.TLabel", background=CARD, foreground=TX3, font=(FONTE, 8, "bold"))
        s.configure("Tit.TLabel", background=BG, foreground=TX, font=(FONTE, 14, "bold"))
        s.configure("Sub.TLabel", background=BG, foreground=TX3, font=(FONTE, 9))
        s.configure("Cli.TLabel", background="#f6f8fc", foreground=TX, font=(FONTE, 10))
        s.configure("TEntry", fieldbackground="#f6f8fc", bordercolor=LINHA, padding=6)
        s.configure("TCombobox", fieldbackground="#f6f8fc", bordercolor=LINHA, padding=5)
        s.configure("TSpinbox", fieldbackground="#f6f8fc", bordercolor=LINHA, padding=5)
        s.configure("Rosa.TButton", background=ROSA, foreground="white", font=(FONTE, 12, "bold"),
                    padding=12, borderwidth=0)
        s.map("Rosa.TButton", background=[("active", "#c93a70"), ("disabled", "#e9b3c9")])
        s.configure("Treeview", rowheight=26, font=(FONTE, 10), background=CARD, fieldbackground=CARD,
                    foreground=TX, bordercolor=LINHA, borderwidth=0)
        s.configure("Treeview.Heading", background=TX, foreground="white", font=(FONTE, 9, "bold"),
                    relief="flat", padding=6)
        s.map("Treeview", background=[("selected", ROSA)], foreground=[("selected", "white")])
        s.map("Treeview.Heading", background=[("active", "#1c2c5c")])
        s.configure("TNotebook", background=BG, borderwidth=0)
        s.configure("TNotebook.Tab", padding=(16, 8), font=(FONTE, 10, "bold"), background="#e8edf5", foreground=TX2)
        s.map("TNotebook.Tab", background=[("selected", CARD)], foreground=[("selected", TX)])
        s.configure("Aviso.TButton", background="#fbefe0", foreground="#9a5a06", padding=7, borderwidth=0,
                    font=(FONTE, 10, "bold"))
        s.map("Aviso.TButton", background=[("active", "#f6e1c4")])
        s.configure("Leve.TButton", background="#e8edf5", foreground=TX, padding=7, borderwidth=0)
        s.map("Leve.TButton", background=[("active", "#dbe3ef")])

    def _card(self, pai, **kw):
        f = ttk.Frame(pai, style="Card.TFrame", padding=14)
        f.pack(**kw)
        return f

    def _rot(self, pai, txt):
        ttk.Label(pai, text=txt.upper(), style="Rot.TLabel").pack(anchor="w", pady=(8, 3))

    def _monta(self):
        topo = tk.Frame(self, bg=BG)
        topo.pack(fill="x", padx=16, pady=(12, 0))
        try:
            lg = Image.open(recurso("logo.png")).convert("RGBA")
            lg = lg.resize((int(lg.width * 46 / lg.height), 46), Image.LANCZOS)
            fundo = Image.new("RGBA", lg.size, tuple(int(BG[i:i + 2], 16) for i in (1, 3, 5)) + (255,))
            fundo.alpha_composite(lg)                      # logo rosa direto sobre o fundo, sem caixa branca
            self._logo = ImageTk.PhotoImage(fundo.convert("RGB"))
            tk.Label(topo, image=self._logo, bg=BG, bd=0).pack(side="left")
        except Exception:
            pass
        tit = tk.Frame(topo, bg=BG)
        tit.pack(side="left", padx=12)
        ttk.Label(tit, text="Etiquetas de volume", style="Tit.TLabel").pack(anchor="w")
        self.lbl_base = ttk.Label(tit, text="", style="Sub.TLabel", wraplength=330, justify="left")
        self.lbl_base.pack(anchor="w")
        ttk.Button(topo, text="＋ Cadastros", style="Leve.TButton", command=self.abrir_cadastros).pack(side="right")
        self.btn_fila = ttk.Button(topo, text="", style="Aviso.TButton", command=self._menu_fila)
        self._chip_on = False
        ttk.Button(topo, text="📋 Histórico", style="Leve.TButton", command=self.abrir_historico).pack(side="right")
        ttk.Button(topo, text="⟳ Clientes", style="Leve.TButton", command=self.atualiza_clientes).pack(side="right", padx=8)

        tk.Frame(self, bg=ROSA, height=2).pack(fill="x", padx=16, pady=(10, 0))
        corpo = tk.Frame(self, bg=BG)
        corpo.pack(fill="both", expand=True, padx=16, pady=8)
        esq = tk.Frame(corpo, bg=BG)
        esq.pack(side="left", fill="both", expand=True)
        dir_ = tk.Frame(corpo, bg=BG)
        dir_.pack(side="right", fill="y", padx=(12, 0))

        # --- cartão: dados da etiqueta
        c = self._card(esq, fill="x", pady=(0, 10))
        self._rot(c, "Cliente (código ou nome)")
        self.v_busca = tk.StringVar()
        e = ttk.Entry(c, textvariable=self.v_busca, font=(FONTE, 12))
        e.pack(fill="x")
        e.bind("<KeyRelease>", self.busca)
        e.bind("<Return>", self.confirma_busca)
        e.bind("<Down>", lambda ev: (self.lista.focus_set(), self.lista.selection_set(0)) if self.lista.size() else None)
        self.entry_busca = e
        self.lista = tk.Listbox(c, height=4, bd=0, highlightthickness=1, highlightbackground=LINHA,
                                font=(FONTE, 10), activestyle="none", selectbackground=ROSA)
        self.lista.bind("<Return>", self.escolhe_lista)
        self.lista.bind("<Double-Button-1>", self.escolhe_lista)
        self.lbl_cli = ttk.Label(c, text="Nenhum cliente selecionado", style="Cli.TLabel", justify="left",
                                 padding=10, wraplength=480)
        self.lbl_cli.pack(fill="x", pady=(8, 0))

        linha = tk.Frame(c, bg=CARD)
        linha.pack(fill="x")
        a = tk.Frame(linha, bg=CARD); a.pack(side="left", fill="x", expand=True, padx=(0, 8))
        b = tk.Frame(linha, bg=CARD); b.pack(side="left", padx=(0, 8))
        cc = tk.Frame(linha, bg=CARD); cc.pack(side="left")
        self._rot(a, "Nota fiscal")
        self.v_nf = tk.StringVar()
        en = ttk.Entry(a, textvariable=self.v_nf, font=(FONTE, 12))
        en.pack(fill="x")
        self.entry_nf = en
        en.bind("<Return>", lambda ev: self.imprimir_todas())
        self._rot(b, "Qtd. volumes")
        self.v_vol = tk.StringVar(value="1")
        self.spin_vol = ttk.Spinbox(b, from_=1, to=99, width=5, textvariable=self.v_vol, font=(FONTE, 12),
                                    command=self.preview)
        self.spin_vol.pack()
        self._rot(cc, "Só o volume nº")
        self.v_so = tk.StringVar()
        ttk.Entry(cc, textvariable=self.v_so, width=6, font=(FONTE, 12)).pack()
        for v in (self.v_nf, self.v_vol, self.v_so):
            v.trace_add("write", lambda *a: self.after(50, self.preview))

        ttk.Label(c, text="Ex.: 3 volumes imprime sozinho 1/3, 2/3 e 3/3 (use \"Só o volume nº\" para reimprimir um).",
                  foreground=TX3, background=CARD, font=(FONTE, 9)).pack(anchor="w", pady=(4, 0))
        self._rot(c, "Transportadora")
        self.v_tr = tk.StringVar()
        lt = tk.Frame(c, bg=CARD); lt.pack(fill="x")
        self.cb_tr = ttk.Combobox(lt, textvariable=self.v_tr, font=(FONTE, 11))
        self.cb_tr.pack(side="left", fill="x", expand=True)
        ttk.Button(lt, text="＋ Nova", style="Leve.TButton",
                   command=lambda: self.abrir_cadastros("transportadoras")).pack(side="left", padx=(6, 0))
        self.cb_tr.bind("<<ComboboxSelected>>", lambda e: self.preview())
        self.v_tr.trace_add("write", lambda *a: self.after(50, self.preview))
        self._rot(c, "Remetente")
        self.v_rem = tk.StringVar()
        lr = tk.Frame(c, bg=CARD); lr.pack(fill="x")
        self.cb_rem = ttk.Combobox(lr, textvariable=self.v_rem, state="readonly", font=(FONTE, 11))
        self.cb_rem.pack(side="left", fill="x", expand=True)
        ttk.Button(lr, text="＋ Novo", style="Leve.TButton",
                   command=lambda: self.abrir_cadastros("remetentes")).pack(side="left", padx=(6, 0))
        self.cb_rem.bind("<<ComboboxSelected>>", lambda e: self.preview())

        # --- cartão: impressora
        p = self._card(esq, fill="x")
        self._rot(p, "Impressora ativa")
        self.v_imp = tk.StringVar()
        self.v_w, self.v_h, self.v_dpi = (tk.StringVar(value=self.cfg.get(k, d)) for k, d in
                                          (("w", "100"), ("h", "150"), ("dpi", "203")))
        self.v_todas = tk.BooleanVar(value=bool(self.cfg.get("todas_impressoras", False)))
        linha_imp = tk.Frame(p, bg=CARD)
        linha_imp.pack(fill="x")
        self.v_imp.set(self.cfg.get("impressora", ""))        # a lista real chega em segundo plano
        self.cb_imp = ttk.Combobox(linha_imp, textvariable=self.v_imp, state="readonly", font=(FONTE, 11),
                                   values=[self.v_imp.get()] if self.v_imp.get() else [])
        self.cb_imp.pack(fill="x", side="left", expand=True)
        ttk.Button(linha_imp, text="⟳", width=3, style="Leve.TButton",
                   command=self.refresh_impressoras).pack(side="left", padx=(6, 0))
        ttk.Button(linha_imp, text="⚙ Configurações", style="Leve.TButton",
                   command=self.abrir_configuracoes).pack(side="left", padx=(6, 0))
        self.lbl_imp = ttk.Label(p, text="", foreground=TX3, background=CARD, font=(FONTE, 9))
        self.lbl_imp.pack(anchor="w", pady=(6, 0))
        self.refresh_impressoras(inicial=True)
        lb = tk.Frame(esq, bg=BG)
        lb.pack(fill="x", pady=(10, 0))
        ttk.Button(lb, text="📥 Abrir DANFE (PDF)…", style="Leve.TButton",
                   command=self.abrir_danfe_manual).pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(lb, text="⚠ Comunicado de avaria…", style="Leve.TButton",
                   command=self.abrir_comunicado).pack(side="left", fill="x", expand=True, padx=(5, 0))

        # --- direita: pré-visualização + botões
        pv = self._card(dir_, fill="x")
        self._rot(pv, "Pré-visualização")
        self.canvas = tk.Canvas(pv, width=270, height=405, bg="#f6f8fc", highlightthickness=1,
                                highlightbackground=LINHA)
        self.canvas.pack()
        nav = tk.Frame(pv, bg=CARD); nav.pack(pady=(8, 0))
        ttk.Button(nav, text="‹", width=3, style="Leve.TButton", command=lambda: self.navega(-1)).pack(side="left")
        self.lbl_nav = ttk.Label(nav, text="volume 1", width=14, anchor="center")
        self.lbl_nav.pack(side="left", padx=6)
        ttk.Button(nav, text="›", width=3, style="Leve.TButton", command=lambda: self.navega(1)).pack(side="left")

        self.btn = ttk.Button(dir_, text="Imprimir etiquetas  (F5)", style="Rosa.TButton", command=self.imprimir_todas)
        self.btn.pack(fill="x", pady=(8, 4))
        ttk.Button(dir_, text="Salvar como imagem (PNG)", style="Leve.TButton", command=self.salvar_png).pack(fill="x")

    # ---------- planilha
    def _carrega_base(self):
        self.plan = None
        try:
            self.config(cursor="watch"); self.update()
            self.plan = Planilha()
        except Exception as ex:
            messagebox.showerror("Erro ao abrir os dados", f"{ex}\n\nPasta de dados: {pasta_dados()}")
            return
        finally:
            self.config(cursor="")
        self.recarrega_cadastros(inicial=True)
        self._status()
        self.entry_busca.focus_set()
        self.preview()
        self._agenda_salva()

    def recarrega_cadastros(self, inicial=False):
        """Atualiza as listas de remetente e transportadora (depois de cadastrar, editar ou remover)."""
        rem = self.cfg.get("remetente") if inicial else self.v_rem.get()
        tr = self.cfg.get("transportadora") if inicial else self.v_tr.get()
        self.cb_rem["values"] = list(self.plan.remetentes)
        self.cb_tr["values"] = self.plan.transportadoras
        self.v_rem.set(rem if rem in self.plan.remetentes else next(iter(self.plan.remetentes), ""))
        self.v_tr.set(tr or (self.plan.transportadoras[0] if self.plan.transportadoras else ""))

    def abrir_cadastros(self, aba="transportadoras"):
        j = getattr(self, "_janela_cad", None)
        if j is not None and j.winfo_exists():
            j.mostra(aba)
            return
        self._janela_cad = CadastrosJanela(self, aba)

    def _garante_transportadora(self):
        """Transportadora digitada que ainda não existe é cadastrada sozinha."""
        digitada = self.v_tr.get().strip()
        if not digitada:
            return
        oficial = self.plan.add_transportadora(digitada)
        self.cb_tr["values"] = self.plan.transportadoras
        if oficial and oficial != digitada:
            self.v_tr.set(oficial)

    def _status(self):
        import datetime
        aviso = getattr(self, "_aviso_texto", "") if getattr(self, "_aviso_ativo", False) else ""
        if self.danfe_atual:
            n = len(self.fila_danfe)
            self.lbl_base.config(text=f"📥 DANFE da NF {self.danfe_atual['nf']}: confira os volumes, a transportadora e "
                                      f"o remetente e tecle F5" + (f"  ·  +{n} na fila" if n else "")
                                      + (f"\n{aviso}" if aviso else ""))
            return
        if aviso:                                       # aviso temporário (ex.: NF ignorada) na tela
            return
        n = f"{len(self.plan.clientes):,}".replace(",", ".")
        if len(self.plan.clientes) == 0:
            self.lbl_base.config(text="Nenhum cliente carregado. Coloque o clientes.csv ao lado do programa "
                                      "ou use ⟳ Atualizar clientes.")
            return
        try:
            quando = datetime.datetime.fromtimestamp(os.path.getmtime(self.plan.base)).strftime("%d/%m/%Y %H:%M")
            txt = f"{n} clientes · atualizado em {quando}"
            origem = self.cfg.get("clientes_origem")
            if origem:
                txt += f" · vigiando {os.path.basename(origem.rstrip('/\\')) or origem}"
        except OSError:
            txt = f"{n} clientes · base da planilha (ainda não atualizada)"
        self.lbl_base.config(text=txt)

    # ---------- clientes: vigia o arquivo e atualiza sozinho
    def _adota_origem_clientes(self):
        """Sem origem definida? Usa um 'clientes.csv/xlsx' que estiver ao lado do programa ou da planilha."""
        if self.cfg.get("clientes_origem") or not self.plan:
            return
        for nome in ("clientes.csv", "clientes.xlsx"):
            c = os.path.join(PASTA, nome)
            if os.path.exists(c):
                self.cfg["clientes_origem"] = c
                return

    def definir_origem_clientes(self, caminho):
        """Passa a vigiar este arquivo (ou pasta). Já lê agora."""
        self.cfg["clientes_origem"] = caminho
        self.cfg["clientes_sig"] = ""
        self._salva_cfg()
        self._vigia_clientes(agora=True)

    def _vigia_clientes(self, agora=False):
        """A cada 15 s olha se o arquivo mudou (data/tamanho). Mudou? Lê em segundo plano."""
        if not agora:
            self.after(INTERVALO_VIGIA, self._vigia_clientes)
        if self._lendo or not self.plan:
            return
        alvo = arquivo_clientes(self.cfg.get("clientes_origem"))
        if not alvo:
            return
        try:
            sig = assinatura(alvo)
            recente = time.time() - os.stat(alvo).st_mtime < 3  # ainda sendo gravado
        except OSError:
            return
        if sig == self.cfg.get("clientes_sig") or recente:
            return
        self._lendo = True

        def tarefa():
            try:
                self._fila_cli.put(("ok", alvo, sig, ler_tabela(alvo)))
            except (PermissionError, OSError):
                self._fila_cli.put(("espera", alvo, sig, None))  # aberto/travado: tenta de novo depois
            except Exception as ex:
                self._fila_cli.put(("erro", alvo, sig, str(ex)))

        threading.Thread(target=tarefa, daemon=True).start()
        self.after(400, self._le_clientes)

    def _le_clientes(self):
        try:
            tipo, alvo, sig, dado = self._fila_cli.get_nowait()
        except queue.Empty:
            self.after(400, self._le_clientes)
            return
        self._lendo = False
        if tipo == "espera":
            return
        if tipo == "erro":
            if self.cfg.get("clientes_erro_sig") != sig:
                self.cfg["clientes_erro_sig"] = sig
                self._salva_cfg()
                messagebox.showwarning("Clientes", f"Não consegui ler {os.path.basename(alvo)}:\n{str(dado)[:300]}\n\nA base de clientes foi mantida como estava.")
            return
        d = self._aplica_clientes(sig, dado)
        if d["novos"] or d["alterados"]:
            partes = [f"{d['novos']} cliente(s) novo(s) adicionado(s)"]
            if d["alterados"]:
                partes.append(f"{d['alterados']} atualizado(s)")
            self.lbl_base.config(text="✔ " + " · ".join(partes))
            self.after(10000, self._status)

    def _aplica_clientes(self, sig, df):
        """Arquivo cheio do sistema: quem já está cadastrado é ignorado; só os novos entram."""
        d = self.plan.importar_df(df, bool(self.cfg.get("atualizar_existentes")))
        self.cfg["clientes_sig"] = sig
        self._salva_cfg()
        self._status()
        if self.cliente:  # atualiza a etiqueta aberta, se for o caso
            novo = self.plan.cliente(self.cliente["cod"])
            if novo:
                self.cliente = novo
                self.preview()
        return d

    def atualiza_clientes(self):
        """Escolha manual do arquivo cheio: ele passa a ser vigiado (mudou, os novos entram sozinhos)."""
        if not self.plan:
            return
        c = filedialog.askopenfilename(title="Arquivo de clientes exportado do sistema",
                                       filetypes=[("Clientes (CSV ou Excel)", "*.csv *.xlsx *.xlsm"), ("Todos", "*.*")])
        if not c:
            return
        try:
            self.config(cursor="watch"); self.update()
            df = ler_tabela(c)
            sig = assinatura(c)
        except Exception as ex:
            messagebox.showerror("Não consegui atualizar", str(ex))
            return
        finally:
            self.config(cursor="")
        d = self._aplica_clientes(sig, df)
        self.cfg["clientes_origem"] = c
        self._salva_cfg()
        self.busca()
        extra = f"\n{d['alterados']} atualizado(s)" if d["alterados"] else ""
        messagebox.showinfo("Clientes",
                            f"{d['novos']} cliente(s) novo(s) adicionado(s)\n"
                            f"{d['ignorados']} já cadastrado(s): ignorado(s){extra}\n"
                            f"Total na base: {d['total']}\n\n"
                            "Esse arquivo agora é vigiado: quando você mandar um novo por cima dele, "
                            "os clientes novos entram sozinhos.")

    def refresh_impressoras(self, inicial=False):
        """Relê as impressoras do Windows EM SEGUNDO PLANO (impressoras de rede podem demorar) e mantém a escolhida."""
        todas = self.v_todas.get()
        self.lbl_imp.config(text="Procurando impressoras…", foreground=TX3)
        threading.Thread(target=lambda: self._res_imp.put((inicial, listar_impressoras(todas))), daemon=True).start()
        self.after(120, self._le_res_imp)

    def _le_res_imp(self):
        try:
            inicial, vals = self._res_imp.get_nowait()
        except queue.Empty:
            self.after(120, self._le_res_imp)
            return
        self.cb_imp["values"] = vals
        atual = self.v_imp.get()
        for cand in (atual, self.cfg.get("impressora") if inicial else None, impressora_padrao()):
            if cand and cand in vals:
                self.v_imp.set(cand)
                break
        else:
            self.v_imp.set(vals[0] if vals else "")
        if vals:
            self.lbl_imp.config(text=f"● {len(vals)} impressora(s) ativa(s) encontrada(s)", foreground="#067a54")
        elif sys.platform == "win32":
            self.lbl_imp.config(text="Nenhuma impressora ativa. Ligue a impressora e clique em ⟳.", foreground="#9a5a06")
        else:
            self.lbl_imp.config(text="Lista de impressoras disponível só no Windows.", foreground=TX3)

    def abrir_configuracoes(self):
        j = getattr(self, "_janela_cfg", None)
        if j is not None and j.winfo_exists():
            j.lift()
            return
        self._janela_cfg = ConfigJanela(self)

    # ---------- DANFE do ERP (impressora "Etiqueta Freedom")
    def _vigia_danfe(self):
        self.after(500, self._vigia_danfe)
        self._varre_tudo()

    def _ciclo_zip(self):
        self.arquiva_pendentes()
        self.after(3600 * 1000, self._ciclo_zip)

    def _varre_tudo(self):
        self._varre_entrada()
        self._varre_pasta_erp()

    # ---- pasta onde o próprio ERP salva o DANFE em PDF (alternativa à impressora virtual)
    def _carrega_vistos(self):
        if not hasattr(self, "_vistos"):
            try:
                with open(os.path.join(pasta_dados(), "pdfs_vistos.json"), encoding="utf-8") as f:
                    self._vistos = set(json.load(f))
            except Exception:
                self._vistos = set()
        return self._vistos

    def _marca_visto(self, chave):
        v = self._carrega_vistos()
        v.add(chave)
        try:
            with open(os.path.join(pasta_dados(), "pdfs_vistos.json"), "w", encoding="utf-8") as f:
                json.dump(sorted(v)[-5000:], f)
        except OSError:
            pass

    def definir_pasta_erp(self, pasta):
        """A partir de agora, todo PDF NOVO salvo nessa pasta é lido na hora (os antigos são ignorados)."""
        self.cfg["pasta_erp"] = pasta
        self.cfg["pasta_erp_desde"] = time.time() - 2 if pasta else 0
        self._salva_cfg()

    def _varre_pasta_erp(self):
        pasta = self.cfg.get("pasta_erp")
        if not pasta or not os.path.isdir(pasta):
            return
        desde = self.cfg.get("pasta_erp_desde", 0)
        try:
            arquivos = [os.path.join(pasta, f) for f in os.listdir(pasta) if f.lower().endswith(".pdf")]
        except OSError:
            return
        for arq in arquivos:
            try:
                st = os.stat(arq)
            except OSError:
                continue
            chave = f"{os.path.abspath(arq)}|{st.st_mtime_ns}|{st.st_size}"
            if (st.st_mtime < desde or st.st_size == 0 or time.time() - st.st_mtime < 1.2
                    or chave in self._carrega_vistos()):
                continue
            self._marca_visto(chave)
            self._processa_danfe(arq, mover=False, silencioso=True)

    def _varre_entrada(self):
        """Procura PDFs novos na pasta de entrada (onde a impressora virtual grava o DANFE)."""
        pasta = impressora_virtual.pasta_entrada()
        try:
            arquivos = [os.path.join(pasta, f) for f in os.listdir(pasta) if f.lower().endswith(".pdf")]
        except OSError:
            return
        for arq in arquivos:
            try:
                st = os.stat(arq)
            except OSError:
                continue
            sig = (st.st_mtime_ns, st.st_size)
            anterior = self._tam_prev.get(arq)
            self._tam_prev[arq] = st.st_size
            if (st.st_size == 0 or anterior != st.st_size or time.time() - st.st_mtime < 0.3
                    or self._danfe_sigs.get(arq) == sig):
                continue  # vazio, ainda crescendo (tamanho mudou desde a última olhada), ou já visto
            self._danfe_sigs[arq] = sig
            self._processa_danfe(arq, mover=True)

    def _pasta_zips(self):
        return pasta_saida()

    def _arquiva(self, arq, nome, apagar, mover):
        """DANFE lido -> ZIP do mês (10-2026.zip) na pasta do programa. Se o ZIP estiver travado, o PDF da impressora
        vai para 'processados' e entra no ZIP depois. Devolve o caminho final, ou None se o PDF ainda está em uso."""
        try:
            return arquivo_zip.guarda(arq, nome, self._pasta_zips(), apagar=apagar)
        except OSError:
            if not mover:
                return arq                        # PDF do ERP/manual: fica onde está; só não entra no ZIP agora
            try:
                return self._guarda_pdf(arq, "processados", nome)
            except OSError:
                return None

    def arquiva_pendentes(self):
        """PDFs soltos de antes (ou de quando o ZIP estava aberto) entram nos ZIPs dos meses."""
        base = impressora_virtual.pasta_entrada()
        try:
            arquivo_zip.guarda_pendentes([os.path.join(base, "processados"), os.path.join(base, "ignorados")],
                                         self._pasta_zips())
        except Exception:
            pass

    def _guarda_pdf(self, arq, subpasta, nome):
        """Tira o PDF da pasta de entrada (para a impressora poder gravar o próximo)."""
        destino_dir = os.path.join(impressora_virtual.pasta_entrada(), subpasta)
        os.makedirs(destino_dir, exist_ok=True)
        destino = os.path.join(destino_dir, nome)
        shutil.move(arq, destino)
        return destino

    def _processa_danfe(self, arq, mover, silencioso=False):
        """Lê o DANFE em segundo plano (imagens levam alguns segundos por causa do OCR) e segue em _conclui_danfe.
        silencioso: PDF que chegou sozinho da pasta do ERP (não é DANFE? só avisa na barra, sem janela)."""
        carimbo = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        if not hasattr(self, "_silenciosos"):
            self._silenciosos = {}
        self._silenciosos[arq] = silencioso
        self._lendo_danfe += 1
        if not self.danfe_atual:
            self.lbl_base.config(text="⏳ Lendo o DANFE…")

        def tarefa():
            try:
                self._res_danfe.put(("ok", arq, mover, carimbo, danfe.le_danfes(arq, self.plan)))
            except Exception as ex:
                self._res_danfe.put(("erro", arq, mover, carimbo, str(ex)))

        threading.Thread(target=tarefa, daemon=True).start()
        self.after(300, self._le_res_danfe)

    def _le_res_danfe(self):
        try:
            while True:
                tipo, arq, mover, carimbo, dado = self._res_danfe.get_nowait()
                self._lendo_danfe = max(0, self._lendo_danfe - 1)
                self._conclui_danfe(tipo, arq, mover, carimbo, dado)
        except queue.Empty:
            pass
        if self._lendo_danfe > 0:
            self.after(300, self._le_res_danfe)
        elif not self.danfe_atual and not self._aviso_ativo:
            self._status()

    def _conclui_danfe(self, tipo, arq, mover, carimbo, dado):
        silencioso = getattr(self, "_silenciosos", {}).pop(arq, False)
        if tipo == "erro":
            if silencioso and not mover:
                self._aviso_temporario(f"ℹ PDF ignorado ({os.path.basename(arq)}): não é um DANFE legível")
            elif mover:
                try:
                    novo = self._guarda_pdf(arq, "com_erro", f"erro_{carimbo}.pdf")
                except OSError:
                    novo = arq
                messagebox.showwarning("DANFE", f"{dado}\n\nO arquivo foi guardado em:\n{novo}")
            else:
                messagebox.showwarning("DANFE", dado)
            return
        aceitas, ignoradas = dado
        nfs = [r["nf"] for r in aceitas] + [i["nf"] for i in ignoradas if i["nf"] != "?"]
        nome = ("NF" + "_".join(nfs[:4]) + ("_etc" if len(nfs) > 4 else "") if nfs else "sem_NF") + f"_{carimbo}.pdf"
        if mover or silencioso:                   # impressora virtual ou pasta do ERP (abrir PDF à mão não arquiva)
            destino = self._arquiva(arq, nome, apagar=mover or bool(self.cfg.get("apagar_pdf_erp")), mover=mover)
        else:
            destino = arq
        if destino is None:                       # PDF ainda em uso: tenta de novo no próximo ciclo
            self._danfe_sigs.pop(arq, None)
            return
        for r in aceitas:
            r["arquivo"] = destino
        novas = 0
        for res in aceitas:
            atual = self.danfe_atual["nf"] if self.danfe_atual else None
            if res["nf"] == atual or any(f["nf"] == res["nf"] for f in self.fila_danfe):
                continue  # a mesma NF impressa duas vezes: ignora
            self.fila_danfe.append(res)
            novas += 1
        if ignoradas:
            self._aviso_temporario(self._texto_ignoradas(ignoradas))
        if novas:
            self.mostra_janela()
            if self.danfe_atual is None:
                self._proximo_danfe()
            else:
                self._atualiza_chip_danfe()
                self._status()
        elif not mover and not silencioso and not aceitas:
            messagebox.showinfo("DANFE", "Nenhuma NF desse PDF tem transportadora cadastrada.\n\n"
                                         + self._texto_ignoradas(ignoradas, longo=True))

    @staticmethod
    def _texto_ignoradas(ignoradas, longo=False):
        nfs = [i["nf"] for i in ignoradas if i["nf"] != "?"]
        if longo:
            return "\n".join(f"NF {i['nf']}: {i['motivo']}" for i in ignoradas)
        motivos = {i.get("curto", "transportadora não cadastrada") for i in ignoradas}
        motivo = motivos.pop() if len(motivos) == 1 else "veja o motivo ao abrir o PDF pelo botão DANFE"
        if nfs and len(nfs) <= 3:
            return "ℹ Ignorada(s): NF " + ", ".join(nfs) + f" ({motivo})"
        return f"ℹ {len(ignoradas)} NF(s) ignorada(s): {motivo}"

    def _aviso_temporario(self, texto, ms=12000):
        self._aviso_ativo = True
        self._aviso_texto = texto
        self._aviso_n = getattr(self, "_aviso_n", 0) + 1
        n = self._aviso_n
        self.lbl_base.config(text=texto)

        def fim():
            if n == self._aviso_n:          # só o aviso mais recente apaga a si mesmo
                self._aviso_ativo = False
                self._status()
        self.after(ms, fim)

    def _proximo_danfe(self):
        if not self.fila_danfe:
            self.danfe_atual = None
            self._atualiza_chip_danfe()
            self._status()
            return
        self.danfe_atual = self.fila_danfe.pop(0)
        self._carrega_danfe(self.danfe_atual)

    def _carrega_danfe(self, res):
        """Preenche a tela com o DANFE para o usuário só conferir."""
        if res["cliente"]:
            self.define_cliente(res["cliente"])
        else:                                   # DANFE como imagem: só NF e remetente; o cliente é escolhido na hora
            self.cliente = None
            self.v_busca.set("")
            self.lbl_cli.config(text="Cliente não identificado: digite o código ou o nome acima e tecle Enter.")
        self.v_nf.set(res["nf"])
        self.v_vol.set(str(res["volumes"]))
        self.v_so.set("")
        if res["remetente"]:
            self.v_rem.set(res["remetente"])
        self.v_tr.set(res["transportadora"])
        if res["avisos"]:
            self.lbl_cli.config(text=self.lbl_cli.cget("text") + "\n⚠ " + " ".join(res["avisos"]))
        self._atualiza_chip_danfe()
        self._status()
        self.preview()
        if res["cliente"]:
            self.after(150, lambda: (self.spin_vol.focus_set(), self.spin_vol.selection_range(0, "end")))
        else:
            self.after(150, self.entry_busca.focus_set)

    def _danfe_concluido(self):
        if self.danfe_atual:
            self.danfe_atual = None
            self._proximo_danfe()

    def _atualiza_chip_danfe(self):
        n = len(self.fila_danfe) + (1 if self.danfe_atual else 0)
        if n:
            self.btn_fila.config(text=f"📥 DANFE: {n}")
            if not self._chip_on:
                self.btn_fila.pack(side="right", padx=(0, 8))
                self._chip_on = True
        elif self._chip_on:
            self.btn_fila.pack_forget()
            self._chip_on = False

    def _menu_fila(self):
        m = tk.Menu(self, tearoff=0)
        if self.danfe_atual:
            m.add_command(label=f"Descartar a NF {self.danfe_atual['nf']} e ir para a próxima", command=self._descartar_danfe)
        m.add_command(label="Abrir DANFE (PDF)…", command=self.abrir_danfe_manual)
        m.add_command(label="Abrir a pasta de entrada", command=lambda: self.abre_pasta(impressora_virtual.pasta_entrada()))
        m.tk_popup(self.winfo_pointerx(), self.winfo_pointery())

    def _descartar_danfe(self):
        self.danfe_atual = None
        self._proximo_danfe()

    def abrir_danfe_manual(self):
        c = filedialog.askopenfilename(title="DANFE em PDF", filetypes=[("PDF", "*.pdf")])
        if c:
            self._processa_danfe(c, mover=False)

    def mostra_janela(self):
        """Traz o programa para a frente (quando chega um DANFE)."""
        self.deiconify()
        self.state("normal")
        self.lift()
        self.attributes("-topmost", True)
        self.after(800, lambda: self.attributes("-topmost", False))
        self.focus_force()
        try:
            self.bell()
        except tk.TclError:
            pass

    def _sai_se_ocioso(self):
        """Aberto pela tarefa agendada sem nenhum DANFE novo (ex.: impressão de outra coisa): fecha em silêncio."""
        if self._lendo_danfe == 0 and not self.danfe_atual and not self.fila_danfe:
            self._fechar()
        elif self._lendo_danfe > 0:
            self.after(3000, self._sai_se_ocioso)

    def inicia_escuta(self, srv):
        """Outra abertura do programa avisa esta (ex.: tarefa do ERP) para ler a pasta de entrada na hora."""
        def escuta():
            while True:
                try:
                    c, _ = srv.accept()
                    c.close()
                    self._msgs.put("varrer")
                except OSError:
                    return
        threading.Thread(target=escuta, daemon=True).start()
        self._le_msgs()

    def _le_msgs(self):
        try:
            while True:
                self._msgs.get_nowait()
                self._varre_tudo()
        except queue.Empty:
            pass
        self.after(500, self._le_msgs)

    def abrir_comunicado(self):
        j = getattr(self, "_janela_com", None)
        if j is not None and j.winfo_exists():
            j.lift(); j.focus_force()
            return
        self._janela_com = ComunicadoJanela(self)

    def imprime_comunicado(self, copias=1):
        """Imprime o comunicado de avaria no tamanho da etiqueta configurada."""
        if not self.v_imp.get() or not impressora_esta_ativa(self.v_imp.get(), estrito=not self.v_todas.get()):
            messagebox.showwarning("Impressora", "Nenhuma impressora ativa. Ligue a impressora e tente de novo.")
            return False
        w, h, dpi = self._params()
        try:
            self.config(cursor="watch"); self.update()
            imprimir(self.v_imp.get(), [desenha_comunicado(w, h, dpi)] * max(1, int(copias)), w, h)
        except Exception as ex:
            messagebox.showerror("Erro ao imprimir", str(ex))
            return False
        finally:
            self.config(cursor="")
        return True

    def abre_pasta(self, pasta):
        try:
            os.startfile(pasta)  # só Windows
        except AttributeError:
            messagebox.showinfo("Pasta", pasta)
        except OSError as ex:
            messagebox.showwarning("Pasta", str(ex))

    def trocar_local_historico(self, parent=None):
        atual = self.xlsx_hist()
        c = filedialog.asksaveasfilename(parent=parent or self, title="Onde guardar o Excel do histórico",
                                         defaultextension=".xlsx", initialdir=os.path.dirname(atual),
                                         initialfile=os.path.basename(atual), filetypes=[("Excel", "*.xlsx")],
                                         confirmoverwrite=False)
        if not c:
            return False
        self.cfg["historico_xlsx"] = c
        self._salva_cfg()
        ok, erro = historico.atualizar_excel(c)
        messagebox.showinfo("Excel", f"Histórico será guardado em:\n{c}" + ("" if ok else "\n\n" + erro), parent=parent or self)
        return True

    def configurar_impressora(self):
        """Abre as preferências da impressora no Windows (papel 100x150, escuridão, velocidade)."""
        nome = self.v_imp.get()
        if not nome:
            messagebox.showwarning("Impressora", "Escolha a impressora primeiro.")
            return
        if sys.platform != "win32":
            messagebox.showinfo("Impressora", "Disponível só no Windows.")
            return
        import subprocess
        try:
            subprocess.Popen(["rundll32", "printui.dll,PrintUIEntry", "/e", "/n", nome])
        except Exception as ex:
            messagebox.showerror("Impressora", f"Não consegui abrir as preferências:\n{ex}")

    def busca(self, ev=None):
        if ev and ev.keysym in ("Return", "Down", "Up", "Tab"):
            return
        self.achados = self.plan.buscar(self.v_busca.get()) if self.plan else []
        self.lista.delete(0, "end")
        for cod, nome, cid, uf in self.achados:
            self.lista.insert("end", f"{cod}  ·  {nome}  ·  {cid}/{uf}")
        self._mostra_lista(bool(self.achados))

    def _mostra_lista(self, mostrar):
        if mostrar:
            self.lista.configure(height=min(4, max(1, self.lista.size())))
            self.lista.pack(fill="x", pady=(4, 0), before=self.lbl_cli)
        else:
            self.lista.pack_forget()

    def confirma_busca(self, ev=None):
        if not self.plan:
            return
        cli = self.plan.cliente(self.v_busca.get())
        if cli:
            self.define_cliente(cli)
        elif getattr(self, "achados", None):
            self.define_cliente(self.plan.cliente(self.achados[0][0]))

    def escolhe_lista(self, ev=None):
        sel = self.lista.curselection()
        if sel:
            self.define_cliente(self.plan.cliente(self.achados[sel[0]][0]))

    def define_cliente(self, cli):
        self.cliente = cli
        self.lista.delete(0, "end")
        self._mostra_lista(False)
        aviso = "\n⚠ Cadastro INATIVO no sistema" if cli.get("situacao") == "I" else ""
        self.lbl_cli.config(text=f'{cli["cod"]} · {cli["nome"]}\n{cli["endereco"]}\n{cli["cidade"]}/{cli["uf"]} · CEP {cli["cep"]}{aviso}')
        self.v_busca.set(cli["cod"])
        self.preview()
        self.focus_nf()

    def focus_nf(self):
        self.after(10, self.entry_nf.focus_set)

    # ---------- geração
    def _params(self):
        try:
            return float(self.v_w.get().replace(",", ".")), float(self.v_h.get().replace(",", ".")), int(self.v_dpi.get())
        except ValueError:
            return 100.0, 150.0, 203

    def _volumes(self):
        try:
            total = max(1, int(self.v_vol.get()))
        except ValueError:
            total = 1
        so = self.v_so.get().strip()
        vols = [int(so)] if so.isdigit() and 1 <= int(so) <= total else list(range(1, total + 1))
        return total, vols

    def _gera(self, volume, total):
        w, h, dpi = self._params()
        return desenha_etiqueta(self.plan.remetentes[self.v_rem.get()], self.cliente, self.v_tr.get() or "-",
                                self.v_nf.get().strip() or "-", volume, total, w, h, dpi)

    def navega(self, d):
        total, vols = self._volumes()
        self.vol_prev = min(max(1, self.vol_prev + d), total)
        self.preview()

    def preview(self):
        if not (self.plan and self.cliente and self.v_rem.get()):
            return
        total, vols = self._volumes()
        self.vol_prev = min(max(self.vol_prev, 1), total)
        try:
            img = self._gera(self.vol_prev, total).convert("L")
        except Exception:
            return
        img.thumbnail((270, 405))
        self._foto = ImageTk.PhotoImage(img)
        self.canvas.delete("all")
        self.canvas.create_image(135, 202, image=self._foto)
        self.lbl_nav.config(text=f"volume {self.vol_prev} de {total}")

    def _valida(self):
        if not self.cliente:
            messagebox.showwarning("Falta o cliente", "Digite o código ou o nome do cliente e tecle Enter.")
            return False
        if not self.v_nf.get().strip():
            messagebox.showwarning("Falta a nota fiscal", "Informe o número da NF.")
            return False
        if not self.v_tr.get().strip():
            messagebox.showwarning("Falta a transportadora", "Escolha ou digite a transportadora.")
            return False
        return True

    def imprimir_todas(self):
        if self._imprimindo or not self._valida():
            return
        if not self.v_imp.get():
            messagebox.showwarning("Impressora", "Nenhuma impressora escolhida. Ligue a impressora e clique em ⟳.")
            return
        if not impressora_esta_ativa(self.v_imp.get(), estrito=not self.v_todas.get()):
            messagebox.showwarning("Impressora", f"A impressora “{self.v_imp.get()}” não está ativa. Ligue-a e clique em ⟳.")
            return
        self._garante_transportadora()
        total, vols = self._volumes()
        w, h, _ = self._params()
        if self.cfg.get("confirmar_impressao") and not messagebox.askyesno(
                "Imprimir", f"Imprimir {len(vols)} etiqueta(s) ({vols[0]}/{total} a {vols[-1]}/{total}) da NF "
                            f"{self.v_nf.get()} para\n{self.cliente['nome']}\nna impressora {self.v_imp.get()}?"):
            return
        try:
            imagens = [self._gera(v, total) for v in vols]
        except Exception as ex:
            messagebox.showerror("Erro ao preparar a etiqueta", str(ex))
            return
        snap = {"nf": self.v_nf.get().strip(), "cliente": self.cliente, "tr": self.v_tr.get(), "rem": self.v_rem.get(),
                "imp": self.v_imp.get(), "vols": vols, "total": total}
        self._imprimindo = True
        self.btn.state(["disabled"])
        self._aviso_ativo = False
        self.lbl_base.config(text=f"🖨 Enviando {len(vols)} etiqueta(s) para a impressora…")

        def tarefa():
            try:
                imprimir(snap["imp"], imagens, w, h)
                self._res_print.put(("ok", snap))
            except Exception as ex:
                self._res_print.put(("erro", str(ex)))

        threading.Thread(target=tarefa, daemon=True).start()
        self.after(80, self._le_res_print)

    def _le_res_print(self):
        try:
            tipo, dado = self._res_print.get_nowait()
        except queue.Empty:
            self.after(80, self._le_res_print)
            return
        self._imprimindo = False
        self.btn.state(["!disabled"])
        if tipo == "erro":
            messagebox.showerror("Erro ao imprimir", dado)
            self._status()
            return
        self._salva_cfg()
        self._registra(dado)
        if self.cfg.get("comunicado_junto"):
            self.imprime_comunicado(1)
        self._danfe_concluido()

    # ---------- histórico
    def xlsx_hist(self):
        return self.cfg.get("historico_xlsx") or os.path.join(pasta_saida(), historico.NOME_EXCEL)

    def _registra(self, snap):
        """Guarda cada etiqueta no histórico (rápido). O Excel é atualizado depois, em segundo plano."""
        linhas = [historico.nova_linha(snap["nf"], v, snap["total"], snap["cliente"], snap["tr"], snap["rem"], snap["imp"])
                  for v in snap["vols"]]
        try:
            historico.registrar(linhas)
        except Exception as ex:
            messagebox.showwarning("Histórico", f"A impressão foi feita, mas não consegui registrar no histórico:\n{ex}")
            return
        self._agenda_excel()
        agora = datetime.datetime.now().strftime("%H:%M")
        self._aviso_temporario(f"✔ {len(snap['vols'])} etiqueta(s) enviada(s) e registrada(s) às {agora}", ms=8000)

    def _agenda_excel(self, ms=15000):
        """Refazer o Excel inteiro a cada impressão ficava lento com o histórico grande: agora espera uns segundos e roda
        em segundo plano (várias impressões seguidas viram uma atualização só)."""
        self.cfg["excel_pendente"] = True
        if self._excel_id:
            self.after_cancel(self._excel_id)
        self._excel_id = self.after(ms, self._excel_bg)

    def _excel_bg(self):
        self._excel_id = None
        if self._excel_rodando:
            self._excel_id = self.after(15000, self._excel_bg)
            return
        self._excel_rodando = True
        caminho = self.xlsx_hist()
        threading.Thread(target=lambda: self._res_excel.put(historico.atualizar_excel(caminho)), daemon=True).start()
        self.after(500, self._le_res_excel)

    def _le_res_excel(self):
        try:
            ok, erro = self._res_excel.get_nowait()
        except queue.Empty:
            self.after(500, self._le_res_excel)
            return
        self._excel_rodando = False
        if ok:
            self.cfg["excel_pendente"] = False
            self._salva_cfg()
        else:                                          # ex.: Excel aberto: tenta de novo em 1 minuto, sem incomodar
            self._excel_id = self.after(60000, self._excel_bg)

    def abrir_historico(self):
        j = getattr(self, "_janela_hist", None)
        if j is not None and j.winfo_exists():
            j.deiconify(); j.lift(); j.focus_force()
            j.aplicar()
            return
        self._janela_hist = HistoricoJanela(self)

    def salvar_png(self):
        if not self._valida():
            return
        total, vols = self._volumes()
        pasta = filedialog.askdirectory(title="Salvar PNGs em…")
        if not pasta:
            return
        for v in vols:
            self._gera(v, total).convert("L").save(os.path.join(pasta, f"etiqueta_NF{self.v_nf.get()}_vol{v}.png"))
        messagebox.showinfo("Pronto", f"{len(vols)} imagem(ns) salva(s).")


class ComunicadoJanela(tk.Toplevel):
    """Prévia e impressão do comunicado de avaria, no tamanho da etiqueta."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Comunicado de avaria")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.transient(app)
        w, h, dpi = app._params()
        corpo = tk.Frame(self, bg=BG)
        corpo.pack(padx=16, pady=14)
        ttk.Label(corpo, text="Comunicado de avaria", style="Tit.TLabel").pack(anchor="w")
        tk.Frame(corpo, bg=ROSA, height=2).pack(fill="x", pady=(8, 10))

        card = ttk.Frame(corpo, style="Card.TFrame", padding=14)
        card.pack()
        self._img = desenha_comunicado(w, h, dpi).convert("L")
        prev = self._img.copy()
        prev.thumbnail((330, 500))
        self._foto = ImageTk.PhotoImage(prev)
        tk.Label(card, image=self._foto, bg=CARD, bd=1, relief="solid").pack()
        ttk.Label(card, text=f"Tamanho: {w:g} × {h:g} mm · {dpi} dpi (muda em ⚙ Configurações)",
                  foreground=TX3, background=CARD, font=(FONTE, 9)).pack(anchor="w", pady=(8, 0))

        lin = tk.Frame(corpo, bg=BG)
        lin.pack(fill="x", pady=(12, 0))
        f = tk.Frame(lin, bg=BG); f.pack(side="left")
        ttk.Label(f, text="CÓPIAS", style="Sub.TLabel").pack(anchor="w")
        self.v_cop = tk.StringVar(value="1")
        ttk.Spinbox(f, from_=1, to=99, width=5, textvariable=self.v_cop, font=(FONTE, 12)).pack()
        ttk.Button(lin, text="Imprimir", style="Rosa.TButton", command=self.imprimir).pack(side="right")
        ttk.Button(lin, text="Salvar PNG", style="Leve.TButton", command=self.salvar).pack(side="right", padx=8)
        self.update_idletasks()
        self.geometry(f"+{app.winfo_rootx() + 60}+{max(0, app.winfo_rooty() + 10)}")

    def imprimir(self):
        try:
            n = max(1, min(99, int(self.v_cop.get())))
        except ValueError:
            n = 1
        if self.app.imprime_comunicado(n):
            self.destroy()

    def salvar(self):
        c = filedialog.asksaveasfilename(parent=self, title="Salvar comunicado", defaultextension=".png",
                                         initialfile="Comunicado_de_avaria.png", filetypes=[("Imagem PNG", "*.png")])
        if c:
            self._img.save(c)


class CadastrosJanela(tk.Toplevel):
    """Cria, edita e remove transportadoras e remetentes. Tudo é salvo na hora, na pasta do programa."""

    def __init__(self, app, aba="transportadoras"):
        super().__init__(app)
        self.app = app
        self.title("Cadastros")
        self.configure(bg=BG)
        self.geometry("820x540")
        self.minsize(760, 500)
        self.transient(app)

        topo = tk.Frame(self, bg=BG)
        topo.pack(fill="x", padx=16, pady=(12, 0))
        ttk.Label(topo, text="Cadastros", style="Tit.TLabel").pack(side="left")
        self.lbl_ok = ttk.Label(topo, text="", style="Sub.TLabel")
        self.lbl_ok.pack(side="left", padx=14, pady=(6, 0))
        tk.Frame(self, bg=ROSA, height=2).pack(fill="x", padx=16, pady=(10, 10))

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        self._aba_transportadoras()
        self._aba_remetentes()
        self.mostra(aba)

    def mostra(self, aba):
        self.nb.select(0 if aba == "transportadoras" else 1)
        self.lift()
        self.focus_force()

    def _ok(self, txt):
        self.lbl_ok.config(text="✔ " + txt)
        self.after(5000, lambda: self.lbl_ok.config(text="") if self.winfo_exists() else None)

    def _lista(self, pai):
        cx = tk.Frame(pai, bg=CARD)
        lb = tk.Listbox(cx, bd=0, highlightthickness=1, highlightbackground=LINHA, font=(FONTE, 10),
                        activestyle="none", selectbackground=ROSA, selectforeground="white", exportselection=False)
        sb = ttk.Scrollbar(cx, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        return cx, lb

    # ---------------------------------------------------------- transportadoras
    def _aba_transportadoras(self):
        t = ttk.Frame(self.nb, style="Card.TFrame", padding=14)
        self.nb.add(t, text="Transportadoras")
        esq = tk.Frame(t, bg=CARD); esq.pack(side="left", fill="both", expand=True)
        ttk.Label(esq, text="CADASTRADAS", style="Rot.TLabel").pack(anchor="w", pady=(0, 4))
        cx, self.lb_t = self._lista(esq)
        cx.pack(fill="both", expand=True)
        self.lb_t.bind("<<ListboxSelect>>", self._sel_t)

        dir_ = tk.Frame(t, bg=CARD); dir_.pack(side="left", fill="y", padx=(16, 0))
        ttk.Label(dir_, text="NOME DA TRANSPORTADORA", style="Rot.TLabel").pack(anchor="w", pady=(0, 4))
        self.v_nome_t = tk.StringVar()
        e = ttk.Entry(dir_, textvariable=self.v_nome_t, width=32, font=(FONTE, 11))
        e.pack(fill="x")
        e.bind("<Return>", lambda ev: self.add_t())
        self.entry_t = e
        ttk.Button(dir_, text="＋ Adicionar nova", style="Rosa.TButton", command=self.add_t).pack(fill="x", pady=(10, 4))
        ttk.Button(dir_, text="Salvar novo nome (da selecionada)", style="Leve.TButton", command=self.renomeia_t).pack(fill="x", pady=2)
        ttk.Button(dir_, text="Remover selecionada", style="Leve.TButton", command=self.remove_t).pack(fill="x", pady=2)
        ttk.Label(dir_, text="Dica: também dá para digitar uma transportadora nova direto\n"
                             "na tela principal. Ela é cadastrada sozinha ao imprimir.",
                  foreground=TX3, background=CARD, font=(FONTE, 9), justify="left").pack(anchor="w", pady=(14, 0))
        self._enche_t()

    def _enche_t(self, selecionar=None):
        self.lb_t.delete(0, "end")
        for t in self.app.plan.transportadoras:
            self.lb_t.insert("end", t)
        if selecionar in self.app.plan.transportadoras:
            i = self.app.plan.transportadoras.index(selecionar)
            self.lb_t.selection_set(i); self.lb_t.see(i)

    def _sel_t(self, ev=None):
        sel = self.lb_t.curselection()
        if sel:
            self.v_nome_t.set(self.lb_t.get(sel[0]))

    def _sel_nome_t(self):
        sel = self.lb_t.curselection()
        return self.lb_t.get(sel[0]) if sel else None

    def add_t(self):
        nome = self.v_nome_t.get().strip()
        if not nome:
            messagebox.showwarning("Transportadora", "Digite o nome da transportadora.", parent=self)
            return
        existia = any(t.upper() == nome.upper() for t in self.app.plan.transportadoras)
        oficial = self.app.plan.add_transportadora(nome)
        self.app.recarrega_cadastros()
        self.app.v_tr.set(oficial)
        self._enche_t(oficial)
        self.v_nome_t.set("")
        self._ok(f"{oficial} já estava cadastrada" if existia else f"{oficial} cadastrada")
        self.entry_t.focus_set()

    def renomeia_t(self):
        antigo, novo = self._sel_nome_t(), self.v_nome_t.get().strip()
        if not antigo:
            messagebox.showinfo("Transportadora", "Clique em uma transportadora da lista primeiro.", parent=self)
            return
        if not novo:
            return
        era = self.app.v_tr.get() == antigo
        try:
            nome = self.app.plan.renomear_transportadora(antigo, novo)
        except ValueError as ex:
            messagebox.showwarning("Transportadora", str(ex), parent=self)
            return
        self.app.recarrega_cadastros()
        if era:
            self.app.v_tr.set(nome)
        self._enche_t(nome)
        self._ok("nome alterado")

    def remove_t(self):
        nome = self._sel_nome_t()
        if not nome:
            messagebox.showinfo("Transportadora", "Clique em uma transportadora da lista primeiro.", parent=self)
            return
        if not messagebox.askyesno("Remover", f"Remover {nome} da lista?\n(O histórico continua como está.)", parent=self):
            return
        self.app.plan.remove_transportadora(nome)
        self.app.recarrega_cadastros()
        self.v_nome_t.set("")
        self._enche_t()
        self._ok(f"{nome} removida")

    # ---------------------------------------------------------- remetentes
    CAMPOS_R = ("nome", "endereco", "cidade", "uf", "cep", "fone", "cnpj")

    def _aba_remetentes(self):
        t = ttk.Frame(self.nb, style="Card.TFrame", padding=14)
        self.nb.add(t, text="Remetentes")
        esq = tk.Frame(t, bg=CARD); esq.pack(side="left", fill="y")
        ttk.Label(esq, text="CADASTRADOS", style="Rot.TLabel").pack(anchor="w", pady=(0, 4))
        cx, self.lb_r = self._lista(esq)
        cx.pack(fill="both", expand=True)
        self.lb_r.configure(width=30)
        self.lb_r.bind("<<ListboxSelect>>", self._sel_r)
        self._orig_r = None

        dir_ = tk.Frame(t, bg=CARD); dir_.pack(side="left", fill="both", expand=True, padx=(16, 0))
        self.v_r = {k: tk.StringVar() for k in self.CAMPOS_R}
        grade = tk.Frame(dir_, bg=CARD); grade.pack(fill="x")
        # linhas: empresa, endereço (largas) | cidade+UF | CEP+fone | CNPJ
        ent = lambda k, larg: ttk.Entry(grade, textvariable=self.v_r[k], width=larg, font=(FONTE, 11))
        ent("nome", 40).grid(row=1, column=0, columnspan=2, sticky="we")
        ent("endereco", 40).grid(row=3, column=0, columnspan=2, sticky="we")
        ent("cidade", 24).grid(row=5, column=0, sticky="we", padx=(0, 8))
        ent("uf", 5).grid(row=5, column=1, sticky="w")
        ent("cep", 14).grid(row=7, column=0, sticky="we", padx=(0, 8))
        ent("fone", 16).grid(row=7, column=1, sticky="w")
        ent("cnpj", 24).grid(row=9, column=0, sticky="we")
        grade.columnconfigure(0, weight=1)
        for linha, rot, col in ((0, "EMPRESA *", 0), (2, "ENDEREÇO", 0), (4, "CIDADE", 0), (4, "UF", 1),
                                (6, "CEP", 0), (6, "TELEFONE", 1), (8, "CNPJ", 0)):
            ttk.Label(grade, text=rot, style="Rot.TLabel").grid(row=linha, column=col, sticky="w", pady=(8, 2))

        bt = tk.Frame(dir_, bg=CARD); bt.pack(fill="x", pady=(16, 0))
        ttk.Button(bt, text="Salvar", style="Rosa.TButton", command=self.salva_r).pack(side="left")
        ttk.Button(bt, text="＋ Novo remetente", style="Leve.TButton", command=self.novo_r).pack(side="left", padx=8)
        ttk.Button(bt, text="Remover", style="Leve.TButton", command=self.remove_r).pack(side="left")
        ttk.Label(dir_, text="CEP e CNPJ são formatados sozinhos. Para criar outro, clique em “Novo remetente”.",
                  foreground=TX3, background=CARD, font=(FONTE, 9)).pack(anchor="w", pady=(10, 0))
        self._enche_r()

    def _enche_r(self, selecionar=None):
        self.lb_r.delete(0, "end")
        for n in self.app.plan.remetentes:
            self.lb_r.insert("end", n)
        if selecionar in self.app.plan.remetentes:
            i = list(self.app.plan.remetentes).index(selecionar)
            self.lb_r.selection_set(i); self.lb_r.see(i)

    def _sel_r(self, ev=None):
        sel = self.lb_r.curselection()
        if not sel:
            return
        nome = self.lb_r.get(sel[0])
        r = self.app.plan.remetentes.get(nome, {})
        for k in self.v_r:
            self.v_r[k].set(r.get(k, ""))
        self._orig_r = nome

    def novo_r(self):
        self.lb_r.selection_clear(0, "end")
        for v in self.v_r.values():
            v.set("")
        self._orig_r = None

    def salva_r(self):
        dados = {k: v.get() for k, v in self.v_r.items()}
        era = self._orig_r is not None and self.app.v_rem.get() == self._orig_r
        novo = self._orig_r is None
        try:
            nome = self.app.plan.salva_remetente(dados, original=self._orig_r)
        except ValueError as ex:
            messagebox.showwarning("Remetente", str(ex), parent=self)
            return
        self.app.recarrega_cadastros()
        if novo or era:
            self.app.v_rem.set(nome)
        self._orig_r = nome
        self._enche_r(nome)
        self._sel_r()
        self.app.preview()
        self._ok(f"{nome} cadastrado" if novo else f"{nome} salvo")

    def remove_r(self):
        if not self._orig_r:
            messagebox.showinfo("Remetente", "Clique em um remetente da lista primeiro.", parent=self)
            return
        if len(self.app.plan.remetentes) <= 1:
            messagebox.showwarning("Remetente", "Mantenha pelo menos um remetente cadastrado.", parent=self)
            return
        if not messagebox.askyesno("Remover", f"Remover {self._orig_r}?\n(O histórico continua como está.)", parent=self):
            return
        nome = self._orig_r
        self.app.plan.remove_remetente(nome)
        self.app.recarrega_cadastros()
        self.novo_r()
        self._enche_r()
        self.app.preview()
        self._ok(f"{nome} removido")


class ConfigJanela(tk.Toplevel):
    PRESETS = {"100 × 150 mm (padrão)": (100, 150), "100 × 100 mm": (100, 100), "100 × 80 mm": (100, 80),
               "80 × 50 mm": (80, 50), "60 × 40 mm": (60, 40), "Personalizado": None}

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Configurações")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.transient(app)

        self.v_w, self.v_h = tk.StringVar(value=app.v_w.get()), tk.StringVar(value=app.v_h.get())
        self.v_dpi, self.v_todas = tk.StringVar(value=app.v_dpi.get()), tk.BooleanVar(value=app.v_todas.get())
        self.v_pre = tk.StringVar()
        self.v_atu = tk.BooleanVar(value=bool(app.cfg.get("atualizar_existentes")))
        self.v_com = tk.BooleanVar(value=bool(app.cfg.get("comunicado_junto")))
        self.v_zpl = tk.BooleanVar(value=bool(app.cfg.get("modo_zpl")))
        self.v_conf = tk.BooleanVar(value=bool(app.cfg.get("confirmar_impressao")))
        self.v_apaga = tk.BooleanVar(value=bool(app.cfg.get("apagar_pdf_erp")))
        self.v_ini = tk.BooleanVar(value=impressora_virtual.inicio_automatico_ativo())
        self._ini_antes = self.v_ini.get()

        corpo = tk.Frame(self, bg=BG)
        corpo.pack(padx=16, pady=10)
        tk.Frame(corpo, bg=ROSA, height=2).pack(fill="x", pady=(0, 8))
        nb = ttk.Notebook(corpo)
        nb.pack(fill="x", pady=(0, 8))
        t1, t2, t3 = (tk.Frame(nb, bg=BG, pady=8) for _ in range(3))
        nb.add(t1, text="Etiqueta e impressora")
        nb.add(t2, text="Arquivos e clientes")
        nb.add(t3, text="ERP · DANFE")

        # --- etiqueta
        c = ttk.Frame(t1, style="Card.TFrame", padding=10); c.pack(fill="x", pady=(0, 6))
        ttk.Label(c, text="ETIQUETA", style="Rot.TLabel").pack(anchor="w")
        ttk.Label(c, text="Tamanho do papel (o mesmo da impressora)", foreground=TX2,
                  background=CARD, font=(FONTE, 9)).pack(anchor="w", pady=(0, 4))
        self.cb_pre = ttk.Combobox(c, textvariable=self.v_pre, values=list(self.PRESETS), state="readonly",
                                   font=(FONTE, 11), width=26)
        self.cb_pre.pack(anchor="w")
        self.cb_pre.bind("<<ComboboxSelected>>", self._preset)
        lin = tk.Frame(c, bg=CARD); lin.pack(anchor="w", pady=(6, 0))
        for txt, var, larg in (("Largura (mm)", self.v_w, 8), ("Altura (mm)", self.v_h, 8)):
            f = tk.Frame(lin, bg=CARD); f.pack(side="left", padx=(0, 12))
            ttk.Label(f, text=txt.upper(), style="Rot.TLabel").pack(anchor="w", pady=(0, 3))
            ttk.Entry(f, textvariable=var, width=larg, font=(FONTE, 11)).pack()
            var.trace_add("write", lambda *a: self._acha_preset())
        f = tk.Frame(lin, bg=CARD); f.pack(side="left")
        ttk.Label(f, text="RESOLUÇÃO (DPI)", style="Rot.TLabel").pack(anchor="w", pady=(0, 3))
        ttk.Combobox(f, textvariable=self.v_dpi, values=("203", "300", "600"), width=7, font=(FONTE, 11)).pack()
        self._acha_preset()

        # --- impressora
        c = ttk.Frame(t1, style="Card.TFrame", padding=10); c.pack(fill="x", pady=(0, 6))
        ttk.Label(c, text="IMPRESSORA", style="Rot.TLabel").pack(anchor="w")
        ttk.Checkbutton(c, text="Mostrar também impressoras offline e virtuais (PDF, XPS, Fax…)",
                        variable=self.v_todas).pack(anchor="w", pady=(8, 2))
        ttk.Checkbutton(c, text="Imprimir também o comunicado de avaria depois das etiquetas (1 por NF)",
                        variable=self.v_com).pack(anchor="w", pady=(0, 2))
        ttk.Checkbutton(c, text="Modo rápido ZPL: manda direto para a impressora (só Zebra / compatíveis com ZPL)",
                        variable=self.v_zpl).pack(anchor="w", pady=(0, 2))
        ttk.Checkbutton(c, text="Perguntar “Sim/Não” antes de imprimir (padrão: imprime direto ao teclar F5)",
                        variable=self.v_conf).pack(anchor="w", pady=(0, 6))
        ttk.Button(c, text="Preferências da impressora no Windows (papel, escuridão, velocidade)…",
                   style="Leve.TButton", command=app.configurar_impressora).pack(anchor="w")

        # --- arquivos
        c = ttk.Frame(t2, style="Card.TFrame", padding=10); c.pack(fill="x", pady=(0, 6))
        ttk.Label(c, text="ARQUIVOS", style="Rot.TLabel").pack(anchor="w")
        self.lbl_cli = self._arquivo(c, "Clientes exportados do sistema (atualiza sozinho quando o arquivo mudar)",
                                     app.cfg.get("clientes_origem") or "— escolha o arquivo ou a pasta onde você salva a exportação —",
                                     self._troca_clientes_arq, extra=("Pasta…", self._troca_clientes_pasta))
        ttk.Checkbutton(c, text="Também atualizar os dados de clientes que já estão cadastrados (padrão: ignorar)",
                        variable=self.v_atu).pack(anchor="w", pady=(6, 0))
        self.lbl_hist = self._arquivo(c, "Excel do histórico", app.xlsx_hist(), self._troca_hist)
        self._arquivo(c, "Pasta dos dados (clientes, cadastros, histórico e configurações)", pasta_dados(),
                      lambda: self.app.abre_pasta(pasta_dados()), rotulo="Abrir pasta")

        # --- ERP: impressora virtual "Etiqueta Freedom"
        c = ttk.Frame(t3, style="Card.TFrame", padding=12); c.pack(fill="x", pady=(0, 6))
        ttk.Label(c, text="IMPRESSORA “ETIQUETA FREEDOM” (DANFE DO ERP)", style="Rot.TLabel").pack(anchor="w")
        ttk.Label(c, text="No ERP, em Imprimir DANFE, escolha “Etiqueta Freedom”: o programa abre com NF, cliente e remetente "
                          "preenchidos. Confira volumes e transportadora e tecle F5.", foreground=TX2, background=CARD,
                  font=(FONTE, 10), justify="left", wraplength=560).pack(anchor="w", pady=(4, 6))
        self.lbl_erp = ttk.Label(c, text="", background=CARD, font=(FONTE, 10, "bold"), wraplength=560, justify="left")
        self.lbl_erp.pack(anchor="w", pady=(0, 6))
        lb = tk.Frame(c, bg=CARD); lb.pack(anchor="w")
        ttk.Button(lb, text="Instalar a impressora", style="Rosa.TButton", command=self._instalar_erp).pack(side="left")
        ttk.Button(lb, text="Remover", style="Leve.TButton", command=self._remover_erp).pack(side="left", padx=8)
        ttk.Button(lb, text="Abrir a pasta de entrada", style="Leve.TButton",
                   command=lambda: self.app.abre_pasta(impressora_virtual.pasta_entrada())).pack(side="left")
        ttk.Label(c, text="Pede permissão de administrador uma vez. O ERP e este programa precisam estar no mesmo computador.",
                  foreground=TX3, background=CARD, font=(FONTE, 9), justify="left", wraplength=560).pack(anchor="w", pady=(6, 0))
        self._atualiza_erp()

        c = ttk.Frame(t3, style="Card.TFrame", padding=12); c.pack(fill="x", pady=(0, 6))
        ttk.Label(c, text="OU: O ERP SALVA O DANFE EM PDF SOZINHO", style="Rot.TLabel").pack(anchor="w")
        ttk.Label(c, text="Se o ERP salvar o DANFE em PDF numa pasta, escolha-a: PDF novo é lido na hora (com texto, leitura "
                          "exata). Os antigos são ignorados.", foreground=TX2, background=CARD, font=(FONTE, 10),
                  justify="left", wraplength=560).pack(anchor="w", pady=(4, 0))
        self.lbl_pasta_erp = self._arquivo(c, "Pasta onde o ERP salva", app.cfg.get("pasta_erp") or "— nenhuma —",
                                           self._escolhe_pasta_erp, rotulo="Escolher…",
                                           extra=("Limpar", self._limpa_pasta_erp))
        ttk.Checkbutton(c, text="Abrir o programa junto com o Windows (minimizado, aparece quando chegar um DANFE)",
                        variable=self.v_ini).pack(anchor="w", pady=(8, 0))
        ttk.Checkbutton(c, text="Apagar o PDF da pasta do ERP depois de guardar no ZIP (economiza espaço)",
                        variable=self.v_apaga).pack(anchor="w", pady=(2, 0))
        ttk.Label(c, text="Todo DANFE lido é guardado em um ZIP por mês na pasta do programa (ex.: 10-2026.zip).",
                  foreground=TX3, background=CARD, font=(FONTE, 9)).pack(anchor="w", pady=(6, 0))

        # --- botões
        bt = tk.Frame(corpo, bg=BG); bt.pack(fill="x")
        ttk.Button(bt, text="Salvar", style="Rosa.TButton", command=self.salvar).pack(side="right")
        ttk.Button(bt, text="Cancelar", style="Leve.TButton", command=self.destroy).pack(side="right", padx=8)
        self.update_idletasks()
        self.geometry(f"+{app.winfo_rootx() + 40}+{max(0, app.winfo_rooty() + 10)}")
        self.grab_set()

    def _escolhe_pasta_erp(self):
        p = filedialog.askdirectory(parent=self, title="Pasta onde o ERP salva o DANFE em PDF")
        if p:
            self.app.definir_pasta_erp(p)
            self.lbl_pasta_erp.config(text=p)

    def _limpa_pasta_erp(self):
        self.app.definir_pasta_erp("")
        self.lbl_pasta_erp.config(text="— nenhuma —")

    def _atualiza_erp(self):
        imp, tar = impressora_virtual.impressora_instalada(), impressora_virtual.tarefa_instalada()
        if imp and tar:
            self.lbl_erp.config(text="✔ Instalada: a impressora “Etiqueta Freedom” e a abertura automática estão ativas.",
                                foreground="#067a54")
        elif imp or tar:
            self.lbl_erp.config(text=f"Instalação incompleta (impressora: {'sim' if imp else 'não'} · abertura automática: "
                                     f"{'sim' if tar else 'não'}). Clique em Instalar para refazer.", foreground="#9a5a06")
        else:
            self.lbl_erp.config(text="Ainda não instalada neste computador."
                                     + ("" if impressora_virtual.eh_exe() else " (A instalação só funciona no programa .exe.)"),
                                foreground=TX2)

    def _roda_admin(self, funcao, pergunta, ok_msg):
        if not impressora_virtual.eh_exe():
            messagebox.showinfo("ERP", "Isso só funciona no programa .exe (não no código-fonte).", parent=self)
            return
        if not messagebox.askyesno("ERP", pergunta, parent=self):
            return
        try:
            self.config(cursor="watch"); self.update()
            log = funcao()
        except Exception as ex:
            messagebox.showerror("ERP", str(ex), parent=self)
            return
        finally:
            self.config(cursor="")
        self.app.refresh_impressoras()
        self._atualiza_erp()
        if log.rstrip().endswith("OK"):
            messagebox.showinfo("ERP", ok_msg, parent=self)
        else:
            messagebox.showwarning("ERP", "Não deu certo" + (":\n\n" + log if log else
                                   " (a permissão de administrador foi negada?)."), parent=self)

    def _instalar_erp(self):
        self._roda_admin(impressora_virtual.instalar,
                         "O Windows vai pedir permissão de administrador (uma vez).\nInstalar a impressora “Etiqueta Freedom”?",
                         "Pronto! No ERP, escolha a impressora “Etiqueta Freedom” ao imprimir o DANFE.")

    def _remover_erp(self):
        self._roda_admin(impressora_virtual.remover,
                         "Remover a impressora “Etiqueta Freedom” e a abertura automática?", "Removido.")

    def _arquivo(self, pai, titulo, caminho, comando, extra=None, rotulo=None):
        ttk.Label(pai, text=titulo, foreground=TX2, background=CARD, font=(FONTE, 9)).pack(anchor="w", pady=(6, 2))
        lin = tk.Frame(pai, bg=CARD); lin.pack(fill="x")
        lbl = ttk.Label(lin, text=caminho, background="#f6f8fc", padding=6, width=48, anchor="w")
        lbl.pack(side="left", fill="x", expand=True)
        ttk.Button(lin, text=rotulo or ("Arquivo…" if extra else "Trocar…"), style="Leve.TButton", command=comando).pack(side="left", padx=(6, 0))
        if extra:
            ttk.Button(lin, text=extra[0], style="Leve.TButton", command=extra[1]).pack(side="left", padx=(6, 0))
        return lbl

    def _troca_clientes_arq(self):
        c = filedialog.askopenfilename(parent=self, title="Arquivo de clientes",
                                       filetypes=[("Clientes (CSV ou Excel)", "*.csv *.xlsx *.xlsm"), ("Todos", "*.*")])
        if c:
            self.app.definir_origem_clientes(c)
            self.lbl_cli.config(text=c)

    def _troca_clientes_pasta(self):
        p = filedialog.askdirectory(parent=self, title="Pasta onde você salva a exportação de clientes")
        if p:
            self.app.definir_origem_clientes(p)
            self.lbl_cli.config(text=p + "   (usa o mais recente com “client” no nome)")

    def _troca_hist(self):
        if self.app.trocar_local_historico(self):
            self.lbl_hist.config(text=self.app.xlsx_hist())

    def _preset(self, ev=None):
        medidas = self.PRESETS.get(self.v_pre.get())
        if medidas:
            self.v_w.set(str(medidas[0])); self.v_h.set(str(medidas[1]))

    def _acha_preset(self):
        try:
            atual = (float(self.v_w.get().replace(",", ".")), float(self.v_h.get().replace(",", ".")))
        except ValueError:
            atual = None
        for nome, m in self.PRESETS.items():
            if m and atual == (float(m[0]), float(m[1])):
                self.v_pre.set(nome)
                return
        self.v_pre.set("Personalizado")

    def salvar(self):
        try:
            w, h = float(self.v_w.get().replace(",", ".")), float(self.v_h.get().replace(",", "."))
            dpi = int(self.v_dpi.get())
            if not (20 <= w <= 300 and 20 <= h <= 400 and 100 <= dpi <= 1200):
                raise ValueError
        except ValueError:
            messagebox.showwarning("Configurações", "Confira os números: largura e altura em mm (ex.: 100 e 150) e DPI (ex.: 203).", parent=self)
            return
        a = self.app
        a.v_w.set(f"{w:g}"); a.v_h.set(f"{h:g}"); a.v_dpi.set(str(dpi))
        a.v_todas.set(self.v_todas.get())
        a.cfg["atualizar_existentes"] = bool(self.v_atu.get())
        a.cfg["comunicado_junto"] = bool(self.v_com.get())
        a.cfg["apagar_pdf_erp"] = bool(self.v_apaga.get())
        a.cfg["modo_zpl"] = bool(self.v_zpl.get())
        a.cfg["confirmar_impressao"] = bool(self.v_conf.get())
        MODO_ZPL[0] = a.cfg["modo_zpl"]
        if self.v_ini.get() != self._ini_antes:
            try:
                impressora_virtual.define_inicio_automatico(self.v_ini.get())
            except Exception as ex:
                messagebox.showwarning("Abrir com o Windows", str(ex), parent=self)
        a.refresh_impressoras()
        a._salva_cfg()
        a.preview()
        self.destroy()


class HistoricoJanela(tk.Toplevel):
    LIMITE = 3000  # linhas mostradas na tela (o Excel leva tudo)

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Histórico de etiquetas")
        self.configure(bg=BG)
        self.geometry("1060x660")
        self.minsize(940, 540)
        self._df = None

        topo = tk.Frame(self, bg=BG)
        topo.pack(fill="x", padx=16, pady=(12, 0))
        ttk.Label(topo, text="Histórico de etiquetas", style="Tit.TLabel").pack(side="left")
        self.lbl_res = ttk.Label(topo, text="", style="Sub.TLabel")
        self.lbl_res.pack(side="left", padx=14, pady=(6, 0))
        tk.Frame(self, bg=ROSA, height=2).pack(fill="x", padx=16, pady=(10, 10))

        # --- filtros
        f = ttk.Frame(self, style="Card.TFrame", padding=12)
        f.pack(fill="x", padx=16)
        l1 = tk.Frame(f, bg=CARD); l1.pack(fill="x")
        self.v_de, self.v_ate = tk.StringVar(), tk.StringVar()
        self.v_tr, self.v_nf, self.v_cli = tk.StringVar(value="Todas"), tk.StringVar(), tk.StringVar()

        def campo(pai, rotulo, var, larg, **kw):
            c = tk.Frame(pai, bg=CARD); c.pack(side="left", padx=(0, 10))
            ttk.Label(c, text=rotulo.upper(), style="Rot.TLabel").pack(anchor="w", pady=(0, 3))
            if kw.get("combo"):
                w = ttk.Combobox(c, textvariable=var, width=larg, values=kw["combo"], font=(FONTE, 10))
            else:
                w = ttk.Entry(c, textvariable=var, width=larg, font=(FONTE, 10))
            w.pack()
            w.bind("<Return>", lambda e: self.aplicar())
            return w

        campo(l1, "De (dd/mm/aaaa)", self.v_de, 12)
        campo(l1, "Até", self.v_ate, 12)
        trs = ["Todas"] + sorted(set(app.plan.transportadoras if app.plan else []) |
                                 set(historico.ler()["TRANSPORTADORA"]) - {""})
        campo(l1, "Transportadora", self.v_tr, 24, combo=trs)
        campo(l1, "NF", self.v_nf, 10)
        campo(l1, "Cliente (cód. ou nome)", self.v_cli, 24)

        l2 = tk.Frame(f, bg=CARD); l2.pack(fill="x", pady=(10, 0))
        for txt, cmd in (("Hoje", self.hoje), ("Ontem", self.ontem), ("Últimos 7 dias", self.sete), ("Este mês", self.mes), ("Tudo", self.tudo)):
            ttk.Button(l2, text=txt, style="Leve.TButton", command=cmd).pack(side="left", padx=(0, 6))
        ttk.Button(l2, text="Filtrar", style="Rosa.TButton", command=self.aplicar).pack(side="right")
        ttk.Button(l2, text="Limpar", style="Leve.TButton", command=self.tudo).pack(side="right", padx=8)

        # --- tabela
        cx = tk.Frame(self, bg=BG)
        cx.pack(fill="both", expand=True, padx=16, pady=10)
        cols = (("data", "Data", 90), ("hora", "Hora", 70), ("nf", "NF", 80), ("vol", "Vol.", 60),
                ("cliente", "Cliente", 300), ("cidade", "Cidade/UF", 160), ("transp", "Transportadora", 170),
                ("imp", "Impressora", 160))
        self.tree = ttk.Treeview(cx, columns=[c[0] for c in cols], show="headings", selectmode="extended")
        for c, t, w in cols:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="center" if c in ("data", "hora", "nf", "vol") else "w", stretch=c in ("cliente", "transp"))
        self.tree.tag_configure("par", background="#f6f8fc")
        sb = ttk.Scrollbar(cx, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # --- rodapé
        rod = tk.Frame(self, bg=BG)
        rod.pack(fill="x", padx=16, pady=(0, 14))
        ttk.Button(rod, text="Exportar este filtro para Excel", style="Rosa.TButton", command=self.exportar).pack(side="left")
        ttk.Button(rod, text="Abrir Excel completo", style="Leve.TButton", command=self.abrir_excel).pack(side="left", padx=8)
        ttk.Button(rod, text="Trocar local do Excel…", style="Leve.TButton", command=self.trocar_local).pack(side="left")
        self.lbl_local = ttk.Label(rod, text="", style="Sub.TLabel")
        self.lbl_local.pack(side="right")

        self.hoje()

    # ---- períodos rápidos
    def _periodo(self, de, ate):
        fmt = lambda d: d.strftime("%d/%m/%Y") if d else ""
        self.v_de.set(fmt(de)); self.v_ate.set(fmt(ate))
        self.aplicar()

    def hoje(self):
        h = datetime.date.today(); self._periodo(h, h)

    def ontem(self):
        o = datetime.date.today() - datetime.timedelta(days=1); self._periodo(o, o)

    def sete(self):
        h = datetime.date.today(); self._periodo(h - datetime.timedelta(days=6), h)

    def mes(self):
        h = datetime.date.today(); self._periodo(h.replace(day=1), h)

    def tudo(self):
        self.v_tr.set("Todas"); self.v_nf.set(""); self.v_cli.set("")
        self._periodo(None, None)

    # ---- filtro
    def _filtrado(self):
        de, ate = historico.data_br(self.v_de.get()), historico.data_br(self.v_ate.get())
        if de and ate and de > ate:
            de, ate = ate, de
        return historico.filtrar(historico.ler(), de, ate, self.v_tr.get(), self.v_nf.get(), self.v_cli.get())

    def aplicar(self):
        try:
            df = self._filtrado()
        except ValueError as ex:
            messagebox.showwarning("Data", str(ex), parent=self)
            return
        self._df = df
        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(df.head(self.LIMITE).itertuples(index=False)):
            self.tree.insert("", "end", tags=("par",) if i % 2 else (), values=(
                historico.data_exibicao(r[0]), r[1][:5], r[2], f"{r[3]}/{r[4]}", r[6], f"{r[7]}/{r[8]}", r[9], r[11]))
        res = historico.resumo(df)
        txt = f"{res['etiquetas']} etiqueta(s) · {res['nfs']} NF(s) · {res['clientes']} cliente(s)"
        if len(df) > self.LIMITE:
            txt += f" · mostrando as {self.LIMITE} mais recentes (o Excel leva todas)"
        self.lbl_res.config(text=txt)
        self.lbl_local.config(text=os.path.basename(self.app.xlsx_hist()))

    # ---- Excel
    def _abre(self, caminho):
        try:
            os.startfile(caminho)  # só Windows
        except AttributeError:
            messagebox.showinfo("Excel", f"Arquivo salvo em:\n{caminho}", parent=self)
        except OSError as ex:
            messagebox.showwarning("Excel", f"Não consegui abrir:\n{ex}", parent=self)

    def exportar(self):
        if self._df is None or self._df.empty:
            messagebox.showinfo("Exportar", "Não há etiquetas neste filtro.", parent=self)
            return
        de, ate = self.v_de.get().replace("/", "-"), self.v_ate.get().replace("/", "-")
        nome = "Etiquetas" + (f"_{de}" if de else "") + (f"_a_{ate}" if ate and ate != de else "") + ".xlsx"
        c = filedialog.asksaveasfilename(parent=self, title="Salvar Excel", defaultextension=".xlsx",
                                         initialfile=nome, filetypes=[("Excel", "*.xlsx")])
        if not c:
            return
        try:
            historico.exportar_xlsx(self._df.sort_values(["DATA", "HORA"], kind="stable"), c)
        except PermissionError:
            messagebox.showwarning("Exportar", "Esse arquivo está aberto no Excel. Feche-o ou escolha outro nome.", parent=self)
            return
        except Exception as ex:
            messagebox.showerror("Exportar", str(ex), parent=self)
            return
        if messagebox.askyesno("Exportar", f"{len(self._df)} etiqueta(s) salvas.\nAbrir agora?", parent=self):
            self._abre(c)

    def abrir_excel(self):
        caminho = self.app.xlsx_hist()
        ok, erro = historico.atualizar_excel(caminho)
        if not ok and not os.path.exists(caminho):
            messagebox.showwarning("Excel", erro, parent=self)
            return
        if not ok:
            messagebox.showinfo("Excel", "Abrindo a versão anterior do arquivo.\n" + erro, parent=self)
        self._abre(caminho)

    def trocar_local(self):
        if self.app.trocar_local_historico(self):
            self.aplicar()


def instancia_unica():
    """Devolve o socket de escuta; ou None se o programa já está aberto (ele é avisado e esta cópia sai)."""
    try:
        srv = socket.socket()
        srv.bind(("127.0.0.1", PORTA_LOCAL))
        srv.listen(5)
        return srv
    except OSError:
        try:
            socket.create_connection(("127.0.0.1", PORTA_LOCAL), timeout=1).close()
        except OSError:
            pass
        return None


if __name__ == "__main__":
    servidor = instancia_unica()
    if servidor is None:
        sys.exit(0)
    app = App()
    if not app.bloqueado:
        app.inicia_escuta(servidor)
        app.mainloop()
