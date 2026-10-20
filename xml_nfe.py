# -*- coding: utf-8 -*-
"""Lê o XML da NF-e: dados exatos e instantâneos (sem OCR). A chave de acesso (do código de barras do DANFE) acha o arquivo."""
import os
import re
import time
import xml.etree.ElementTree as ET

from etiqueta_core import fmt_cep, limpa


def _t(e):
    return e.tag.split("}")[-1]


def _filho(e, nome):
    if e is None:
        return None
    for c in e:
        if _t(c) == nome:
            return c
    return None


def _txt(e, *caminho):
    for n in caminho:
        e = _filho(e, n)
        if e is None:
            return ""
    return limpa(e.text or "")


def acha_xml(pasta, chave, nf="", emitentes=(), max_arquivos=600, dias=15):
    """1) arquivo com a chave no nome; 2) a chave dentro dos XMLs mais recentes; 3) sem chave legível: o nº da NF
    com o CNPJ de um emitente conhecido. Devolve o caminho ou ''."""
    if not pasta or not os.path.isdir(pasta) or not (chave or nf):
        return ""
    candidatos = []
    limite = time.time() - dias * 86400
    for raiz, _, arquivos in os.walk(pasta):
        for f in arquivos:
            if not f.lower().endswith(".xml"):
                continue
            c = os.path.join(raiz, f)
            if chave and chave in f:
                return c
            try:
                m = os.path.getmtime(c)
            except OSError:
                continue
            if m >= limite:
                candidatos.append((m, c))
        if raiz.count(os.sep) - pasta.count(os.sep) >= 3:      # não desce demais
            break
    padrao_nf = re.compile(rb"<nNF>\s*0*" + re.escape(str(int(nf)).encode()) + rb"\s*</nNF>") if str(nf).isdigit() else None
    emit = [e.encode() for e in emitentes if e]
    for _, c in sorted(candidatos, reverse=True)[:max_arquivos]:
        try:
            with open(c, "rb") as f:
                topo = f.read(60000)                                # chave e dados principais ficam no começo
        except OSError:
            continue
        if chave and chave.encode() in topo:
            return c
        if padrao_nf and padrao_nf.search(topo) and (not emit or any(e in topo for e in emit)):
            return c
    return ""


def le_xml(caminho):
    """Dados da NF-e no mesmo formato que o leitor do DANFE usa."""
    arvore = ET.parse(caminho)
    inf = next((e for e in arvore.iter() if _t(e) == "infNFe"), None)
    if inf is None:
        raise ValueError("XML sem infNFe (não é uma NF-e).")
    chave = re.sub(r"\D", "", inf.get("Id", ""))
    ide, emit, dest = _filho(inf, "ide"), _filho(inf, "emit"), _filho(inf, "dest")
    transp = _filho(inf, "transp")
    ed = _filho(dest, "enderDest")
    nf = str(int(_txt(ide, "nNF") or 0))
    doc = _txt(dest, "CNPJ") or _txt(dest, "CPF")
    cnpj_fmt = (f"{doc[:2]}.{doc[2:5]}.{doc[5:8]}/{doc[8:12]}-{doc[12:]}" if len(doc) == 14 else
                f"{doc[:3]}.{doc[3:6]}.{doc[6:9]}-{doc[9:]}" if len(doc) == 11 else doc)
    lgr, nro = _txt(ed, "xLgr"), _txt(ed, "nro")
    endereco = lgr + (f", {nro}" if nro and nro.upper() not in ("S/N", "SN", "0") else "")
    transporta = _filho(transp, "transporta")
    nome_t = _txt(transporta, "xNome")
    volumes = 0
    for v in (transp if transp is not None else []):
        if _t(v) == "vol":
            q = _txt(v, "qVol")
            if q.isdigit():
                volumes += int(q)
    inf_cpl = _txt(_filho(inf, "infAdic"), "infCpl")
    cod = re.search(r"Cliente\s+Fantasia\s*:\s*(\d+)", inf_cpl, re.I)
    emissao = _txt(ide, "dhEmi") or _txt(ide, "dEmi")
    return {
        "arquivo": caminho, "nf": nf, "chave": chave, "emitente_cnpj": re.sub(r"\D", "", _txt(emit, "CNPJ")),
        "cod_cliente": cod.group(1) if cod else "",
        "dest": {"nome": _txt(dest, "xNome"), "cnpj": cnpj_fmt, "endereco": endereco, "bairro": _txt(ed, "xBairro"),
                 "cep": fmt_cep(_txt(ed, "CEP")), "cidade": _txt(ed, "xMun"), "uf": _txt(ed, "UF"),
                 "fone": _txt(ed, "fone")},
        "bloco_transportador": nome_t, "volumes": volumes or None, "emissao": emissao[:10],
        "lido_por_ocr": False, "lido_por_xml": True,
    }
