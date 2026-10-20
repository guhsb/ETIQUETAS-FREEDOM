# -*- coding: utf-8 -*-
"""Histórico de etiquetas impressas: base em CSV (segura) + Excel gerado a partir dela."""
import csv
import datetime as dt
import getpass
import os
import re
import socket
import unicodedata

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from etiqueta_core import pasta_dados

COLS = ["DATA", "HORA", "NF", "VOLUME", "TOTAL", "COD CLIENTE", "CLIENTE", "CIDADE", "UF",
        "TRANSPORTADORA", "REMETENTE", "IMPRESSORA", "USUARIO"]
TITULOS = {"DATA": "Data", "HORA": "Hora", "NF": "NF", "VOLUME": "Volume", "TOTAL": "De",
           "COD CLIENTE": "Cód. cliente", "CLIENTE": "Cliente", "CIDADE": "Cidade", "UF": "UF",
           "TRANSPORTADORA": "Transportadora", "REMETENTE": "Remetente", "IMPRESSORA": "Impressora",
           "USUARIO": "Usuário"}
LARGURAS = {"DATA": 12, "HORA": 9, "NF": 11, "VOLUME": 9, "TOTAL": 7, "COD CLIENTE": 13, "CLIENTE": 46,
            "CIDADE": 24, "UF": 6, "TRANSPORTADORA": 24, "REMETENTE": 26, "IMPRESSORA": 28, "USUARIO": 22}
NOME_EXCEL = "HISTORICO_ETIQUETAS.xlsx"


def caminho_csv():
    return os.path.join(pasta_dados(), "historico_etiquetas.csv")


def _sem_acento(t):
    return "".join(c for c in unicodedata.normalize("NFD", str(t)) if unicodedata.category(c) != "Mn").upper()


def nova_linha(nf, volume, total, cli, transportadora, remetente, impressora, agora=None):
    agora = agora or dt.datetime.now()
    try:
        usuario = f"{getpass.getuser()}@{socket.gethostname()}"
    except Exception:
        usuario = ""
    return {"DATA": agora.date().isoformat(), "HORA": agora.strftime("%H:%M:%S"), "NF": str(nf),
            "VOLUME": str(volume), "TOTAL": str(total), "COD CLIENTE": cli["cod"], "CLIENTE": cli["nome"],
            "CIDADE": cli["cidade"], "UF": cli["uf"], "TRANSPORTADORA": transportadora,
            "REMETENTE": remetente, "IMPRESSORA": impressora, "USUARIO": usuario}


def registrar(linhas):
    """Acrescenta as linhas na base (CSV). Nunca fica 'travada' como um Excel aberto."""
    arq = caminho_csv()
    novo = not os.path.exists(arq)
    with open(arq, "a", newline="", encoding="utf-8-sig" if novo else "utf-8") as f:
        w = csv.DictWriter(f, COLS, delimiter=";")
        if novo:
            w.writeheader()
        for l in linhas:
            w.writerow({c: l.get(c, "") for c in COLS})


def ler():
    try:
        df = pd.read_csv(caminho_csv(), sep=";", dtype=str, encoding="utf-8-sig").fillna("")
    except (FileNotFoundError, pd.errors.EmptyDataError):
        df = pd.DataFrame(columns=COLS)
    for c in COLS:
        if c not in df.columns:
            df[c] = ""
    return df[COLS]


def data_br(texto):
    """'20/10/2026', '20/10/26', '20/10' ou '20102026' -> date. Vazio -> None."""
    t = (texto or "").strip()
    if not t:
        return None
    t = re.sub(r"[.\-]", "/", t)
    if re.fullmatch(r"\d{8}", t):
        t = f"{t[:2]}/{t[2:4]}/{t[4:]}"
    p = t.split("/")
    try:
        d, m = int(p[0]), int(p[1])
        a = int(p[2]) if len(p) > 2 and p[2] else dt.date.today().year
        if a < 100:
            a += 2000
        return dt.date(a, m, d)
    except (ValueError, IndexError):
        raise ValueError(f"Data inválida: “{texto}”. Use dia/mês/ano, por exemplo 20/10/2026.")


def filtrar(df, de=None, ate=None, transportadora="", nf="", cliente="", texto=""):
    r = df
    if de:
        r = r[r["DATA"] >= de.isoformat()]
    if ate:
        r = r[r["DATA"] <= ate.isoformat()]
    if transportadora and transportadora.upper() not in ("TODAS", "(TODAS)"):
        r = r[r["TRANSPORTADORA"].str.upper() == transportadora.upper()]
    if nf.strip():
        r = r[r["NF"].str.contains(re.escape(nf.strip()), regex=True)]
    if cliente.strip():
        q = _sem_acento(cliente.strip())
        alvo = (r["COD CLIENTE"] + " " + r["CLIENTE"]).map(_sem_acento)
        r = r[alvo.str.contains(re.escape(q), regex=True)]
    if texto.strip():
        q = _sem_acento(texto.strip())
        r = r[r.apply(lambda x: q in _sem_acento(" ".join(x.astype(str))), axis=1)]
    return r.sort_values(["DATA", "HORA"], ascending=False, kind="stable")


def resumo(df):
    return {"etiquetas": len(df), "nfs": df["NF"].nunique(), "clientes": df["COD CLIENTE"].nunique()}


def data_exibicao(iso):
    try:
        return dt.date.fromisoformat(iso).strftime("%d/%m/%Y")
    except ValueError:
        return iso


def exportar_xlsx(df, caminho):
    """Gera o Excel formatado (padrão Freedom). Levanta PermissionError se o arquivo estiver aberto."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Etiquetas"
    azul, zebra, linha = "0F1B3D", "F6F8FC", "E1E7EF"
    borda = Border(bottom=Side(style="thin", color=linha))
    ws.append([TITULOS[c] for c in COLS])
    for i, c in enumerate(COLS, 1):
        cel = ws.cell(row=1, column=i)
        cel.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cel.fill = PatternFill("solid", fgColor=azul)
        cel.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = LARGURAS[c]
    ws.row_dimensions[1].height = 22
    for n, (_, r) in enumerate(df.iterrows(), start=2):
        try:
            data = dt.date.fromisoformat(r["DATA"])
        except ValueError:
            data = r["DATA"]
        vals = []
        for c in COLS:
            v = r[c]
            if c == "DATA":
                v = data
            elif c in ("NF", "VOLUME", "TOTAL", "COD CLIENTE") and str(v).isdigit():
                v = int(v)
            vals.append(v)
        ws.append(vals)
        for i, c in enumerate(COLS, 1):
            cel = ws.cell(row=n, column=i)
            cel.font = Font(name="Arial", size=10, color="2B3550")
            cel.border = borda
            if n % 2 == 1:
                cel.fill = PatternFill("solid", fgColor=zebra)
            if c == "DATA":
                cel.number_format = "dd/mm/yyyy"
            if c in ("DATA", "HORA", "NF", "VOLUME", "TOTAL", "COD CLIENTE", "UF"):
                cel.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{max(2, ws.max_row)}"
    tmp = caminho + ".tmp"
    wb.save(tmp)
    os.replace(tmp, caminho)


def atualizar_excel(caminho):
    """Regera o Excel completo a partir da base. Devolve (ok, mensagem_de_erro)."""
    try:
        exportar_xlsx(ler().sort_values(["DATA", "HORA"], kind="stable"), caminho)
        return True, ""
    except PermissionError:
        return False, "O arquivo está aberto no Excel. Feche-o para atualizar (os dados estão salvos e entram na próxima impressão)."
    except Exception as ex:
        return False, str(ex)

