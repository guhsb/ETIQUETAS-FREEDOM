# -*- coding: utf-8 -*-
"""Lê o(s) DANFE(s) de um PDF: várias páginas e várias NFs. Só entram as NFs com transportadora cadastrada."""
import re

from etiqueta_core import _sem_acento, fmt_cep, limpa

RE_CNPJ = re.compile(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b|\b\d{3}\.\d{3}\.\d{3}-\d{2}\b")
MAX_PAGINAS = 40


class SemTexto(ValueError):
    pass


# ------------------------------------------------------------------ utilidades
def _norm(t):
    return _sem_acento(t).upper().strip()


def _linhas(texto):
    return [l.strip() for l in texto.splitlines()]


def _valor_apos(linhas, rotulo, ocorrencia=0):
    """Texto da linha seguinte ao rótulo (ex.: 'NOME / RAZÃO SOCIAL')."""
    alvo, achadas = _norm(rotulo), 0
    for i in range(len(linhas)):
        if _norm(linhas[i]) == alvo:
            if achadas == ocorrencia:
                for j in range(i + 1, min(i + 4, len(linhas))):
                    if linhas[j]:
                        return linhas[j]
                return ""
            achadas += 1
    return ""


def _dv_chave(c43):
    pesos, soma = 2, 0
    for d in reversed(c43):
        soma += int(d) * pesos
        pesos = 2 if pesos == 9 else pesos + 1
    resto = soma % 11
    return 0 if resto in (0, 1) else 11 - resto


def acha_chave(texto):
    """Chave de acesso de 44 dígitos (validada pelo dígito verificador)."""
    compacto = re.sub(r"\s+", "", texto)
    for m in re.finditer(r"(?=(\d{44}))", compacto):
        c = m.group(1)
        if c[20:22] in ("55", "65") and _dv_chave(c[:43]) == int(c[43]):
            return c
    return ""


def acha_transportadora(bloco, transportadoras):
    """Transportadora CADASTRADA que aparece no bloco do transportador (palavras inteiras; OCR sem espaços também)."""
    b = " " + re.sub(r"[^A-Z0-9]+", " ", _norm(bloco)).strip() + " "
    compacto = re.sub(r"[^A-Z0-9]", "", _norm(bloco))
    for t in sorted(transportadoras, key=len, reverse=True):
        nome = re.sub(r"[^A-Z0-9]+", " ", _norm(t)).strip()
        if not nome:
            continue
        if f" {nome} " in b:
            return t
        junto = nome.replace(" ", "")
        if len(junto) >= 7 and junto in compacto:
            return t
    return ""


# ------------------------------------------------------------------ texto por página
def _n_paginas(caminho):
    try:
        from pypdf import PdfReader
        return len(PdfReader(caminho).pages)
    except Exception:
        import pypdfium2 as pdfium
        return len(pdfium.PdfDocument(caminho))


def textos_das_paginas(caminho):
    """Texto de cada página (até MAX_PAGINAS). Página sem texto vem como ''. Tenta pypdf e, se preciso, pdfium."""
    n = min(_n_paginas(caminho), MAX_PAGINAS)
    textos = [""] * n
    try:
        from pypdf import PdfReader
        paginas = PdfReader(caminho).pages
        for i in range(n):
            try:
                textos[i] = paginas[i].extract_text() or ""
            except Exception:
                pass
    except Exception:
        pass
    if any(len(t.strip()) < 50 or not acha_chave(t) for t in textos):
        try:
            import pypdfium2 as pdfium
            pdf = pdfium.PdfDocument(caminho)
            for i in range(n):
                if len(textos[i].strip()) < 50 or not acha_chave(textos[i]):
                    t2 = pdf[i].get_textpage().get_text_range()
                    if acha_chave(t2) or len(t2.strip()) > len(textos[i].strip()):
                        textos[i] = t2
        except Exception:
            pass
    return textos


# ------------------------------------------------------------------ página como imagem (barras + OCR)
def _decodifica_barras(img):
    import zxingcpp
    for r in zxingcpp.read_barcodes(img):
        t = re.sub(r"\D", "", r.text or "")
        if len(t) == 44 and t[20:22] in ("55", "65") and _dv_chave(t[:43]) == int(t[43]):
            return t
    return ""


def chave_por_barras(caminho, pagina=0):
    """Página impressa como imagem: lê a chave de acesso (44 dígitos) pelo código de barras."""
    try:
        from pypdf import PdfReader
        for im in PdfReader(caminho).pages[pagina].images:      # 1) imagem embutida, na resolução original
            c = _decodifica_barras(im.image.convert("L"))
            if c:
                return c
    except Exception:
        pass
    try:
        import pypdfium2 as pdfium
        pag = pdfium.PdfDocument(caminho)[pagina]
        for escala in (4, 6, 3, 8):                              # 2) página desenhada em alta resolução
            c = _decodifica_barras(pag.render(scale=escala).to_pil().convert("L"))
            if c:
                return c
    except Exception:
        pass
    return ""


def _imagem_da_pagina(caminho, dpi=150, pagina=0):
    import pypdfium2 as pdfium
    return pdfium.PdfDocument(caminho)[pagina].render(scale=dpi / 72).to_pil().convert("RGB")


def _quantidade_perto_do_rotulo(caixas):
    """'QUANTIDADE' do bloco do transportador (rótulo à esquerda) e o número logo abaixo dele."""
    rotulos = [c for c in caixas if "QUANTIDADE" in _norm(c["texto"])]
    if not rotulos:
        return None
    meio = max(c["x1"] for c in caixas) * 0.40
    esquerda = [c for c in rotulos if c["x0"] < meio]       # o da tabela de produtos fica à direita: não serve
    if not esquerda:
        return None
    r = min(esquerda, key=lambda c: c["y0"])
    alt = max(1, r["y1"] - r["y0"])
    candidatos = []
    for c in caixas:
        if c is r or not re.fullmatch(r"\d{1,6}", c["texto"].replace(" ", "")):
            continue
        if c["y0"] >= r["y0"] - alt * 0.3 and c["y0"] - r["y0"] < alt * 5 and c["x0"] < r["x1"] + alt * 4 \
                and c["x1"] > r["x0"] - alt:
            candidatos.append((abs(c["y0"] - r["y1"]) + abs(c["x0"] - r["x0"]) * 0.2, c))
    if not candidatos:
        return None
    return int(min(candidatos, key=lambda t: t[0])[1]["texto"].replace(" ", ""))


def ocr_topo(caminho, pagina):
    """Texto da parte de cima da página (onde ficam a chave de acesso, o Nº da NF e o CNPJ do emitente)."""
    import ocr
    if not ocr.disponivel():
        return ""
    img = _imagem_da_pagina(caminho, 150, pagina)
    return "\n".join(ocr.linhas_de_texto(ocr.ler_faixa(img, 0.0, 0.30)))


def info_pagina_imagem(caminho, pagina):
    """Página sem texto: acha a chave pelo código de barras; se não der, pelo texto da chave (OCR); se não der,
    só o número da NF (OCR). Devolve {'chave','nf','topo'} ou None."""
    chave = chave_por_barras(caminho, pagina)
    if chave:
        return {"chave": chave, "nf": str(int(chave[25:34])), "topo": ""}
    topo = ocr_topo(caminho, pagina)
    chave = acha_chave(topo)
    if chave:
        return {"chave": chave, "nf": str(int(chave[25:34])), "topo": topo}
    m = re.search(r"N\s*[º°o*]?\s*[:.]?\s*(\d{3}[.\s]?\d{3}[.\s]?\d{3})\b", topo)
    if m:
        return {"chave": "", "nf": str(int(re.sub(r"\D", "", m.group(1)))), "topo": topo}
    return None


def ocr_transporte(caminho, pagina):
    """Bloco do transportador (1ª página da NF): texto da faixa e QUANTIDADE (= volumes)."""
    import ocr
    if not ocr.disponivel():
        return {"bloco": "", "volumes": None}
    img = _imagem_da_pagina(caminho, 150, pagina)
    caixas = ocr.ler_faixa(img, 0.45, 0.68)
    q = _quantidade_perto_do_rotulo(caixas)
    if q is None:
        caixas = ocr.ler_faixa(img, 0.3, 0.75)
        q = _quantidade_perto_do_rotulo(caixas)
    return {"bloco": " ".join(ocr.linhas_de_texto(caixas)), "volumes": q}


def ocr_cliente_fantasia(caminho, paginas):
    """'Cliente Fantasia: NNNN' nos dados adicionais (normalmente na última página da NF)."""
    import ocr
    if not ocr.disponivel():
        return {}
    for p in reversed(paginas):
        img = _imagem_da_pagina(caminho, 150, p)
        for faixa in ((0.72, 0.99), (0.5, 0.99)):
            linhas = "\n".join(ocr.linhas_de_texto(ocr.ler_faixa(img, *faixa)))
            m = re.search(r"Cliente\s*Fan\w*\s*[:;.]?\s*(\d{1,7})", linhas, re.I)
            if m:
                nome = re.search(r"Cliente\s*Fan\w*\s*[:;.]?\s*\d+\s*-\s*(.+)", linhas, re.I)
                return {"cod_cliente": m.group(1), "nome": limpa(nome.group(1)) if nome else ""}
    return {}


def diagnostico(caminho):
    """Resumo do que existe dentro do PDF (ajuda a entender por que não foi lido)."""
    partes = []
    try:
        from pypdf import PdfReader
        leitor = PdfReader(caminho)
        pag = leitor.pages[0]
        try:
            n_img = len(pag.images)
        except Exception:
            n_img = "?"
        try:
            n_txt = len((pag.extract_text() or "").strip())
        except Exception:
            n_txt = "?"
        extra = ""
        try:
            im = pag.images[0].image
            extra = f" · imagem {im.width}x{im.height} px"
        except Exception:
            pass
        try:
            extra += f" · página {float(pag.mediabox.width) / 72 * 25.4:.0f}x{float(pag.mediabox.height) / 72 * 25.4:.0f} mm"
        except Exception:
            pass
        partes.append(f"{len(leitor.pages)} página(s) · {n_txt} caracteres de texto na 1ª · {n_img} imagem(ns){extra}")
    except Exception as ex:
        partes.append(f"não abriu como PDF ({str(ex)[:60]})")
    return " · ".join(partes)


# ------------------------------------------------------------------ páginas -> NFs
def agrupa_paginas(caminho, textos):
    """Junta as páginas por NF (chave de acesso ou número). Página sem identificação continua a NF anterior."""
    grupos, por_chave, ultimo = [], {}, None
    for i, t in enumerate(textos):
        tem_texto = len(t.strip()) >= 50
        if tem_texto:
            chave = acha_chave(t)
            info = {"chave": chave, "nf": str(int(chave[25:34])), "topo": ""} if chave else None
        else:
            info = info_pagina_imagem(caminho, i)
        if info:
            k = info["chave"] or "NF" + info["nf"]
            g = por_chave.get(k)
            if g is None:
                g = {"chave": info["chave"], "nf": info["nf"], "topo": info["topo"], "paginas": [], "imagem": not tem_texto}
                grupos.append(g)
                por_chave[k] = g
            g["paginas"].append(i)
            ultimo = g
        elif ultimo is not None:
            ultimo["paginas"].append(i)               # continuação (2/3, 3/3...) sem identificação legível
        elif tem_texto and "DANFE" in t.upper():
            g = {"chave": "", "nf": "", "topo": "", "paginas": [i], "imagem": False}
            grupos.append(g)
            ultimo = g
    return grupos


def _dados_do_texto(texto, chave, arquivo):
    linhas = _linhas(texto)
    nf = str(int(chave[25:34])) if chave else ""
    if not nf:
        m = re.search(r"N[ºo°]\s*([\d.]{5,11})", texto)
        nf = str(int(re.sub(r"\D", "", m.group(1)))) if m else ""
    if not nf:
        raise ValueError("Não encontrei o número da NF.")
    emit_cnpj = chave[6:20] if chave else re.sub(r"\D", "", _valor_apos(linhas, "CNPJ"))
    cod = re.search(r"Cliente\s+Fantasia\s*:\s*(\d+)", texto, re.I)
    cnpj_dest = _valor_apos(linhas, "CNPJ / CPF")
    dest = {
        "nome": limpa(_valor_apos(linhas, "NOME / RAZÃO SOCIAL")),
        "cnpj": cnpj_dest if RE_CNPJ.fullmatch(cnpj_dest or "") else "",
        "endereco": limpa(_valor_apos(linhas, "ENDEREÇO")),
        "bairro": limpa(_valor_apos(linhas, "BAIRRO / DISTRITO")),
        "cep": fmt_cep(_valor_apos(linhas, "CEP")),
        "cidade": limpa(_valor_apos(linhas, "MUNICÍPIO")),
        "uf": limpa(_valor_apos(linhas, "UF"))[:2].upper(),
        "fone": limpa(_valor_apos(linhas, "TELEFONE / FAX")),
    }
    if not dest["cnpj"]:
        outros = [c for c in RE_CNPJ.findall(texto) if re.sub(r"\D", "", c) != emit_cnpj]
        dest["cnpj"] = outros[0] if outros else ""
    i = texto.upper().find("TRANSPORTADOR / VOLUMES")
    j = texto.upper().find("DADOS DOS PRODUTOS", i + 1) if i >= 0 else -1
    bloco = texto[i:j] if i >= 0 and j > i else ""
    qtd = None
    q = _valor_apos(_linhas(bloco), "QUANTIDADE")
    if re.fullmatch(r"\d{1,6}", q or ""):
        qtd = int(q)
    emissao = re.search(r"Emiss[ãa]o:\s*(\d{2}/\d{2}/\d{4})", texto)
    return {"arquivo": arquivo, "nf": nf, "chave": chave, "emitente_cnpj": re.sub(r"\D", "", emit_cnpj),
            "cod_cliente": cod.group(1) if cod else "", "dest": dest, "bloco_transportador": bloco, "volumes": qtd,
            "emissao": emissao.group(1) if emissao else "", "lido_por_ocr": False}


def _dados_da_imagem(caminho, grupo, plan):
    """NF impressa como imagem: chave pelo código de barras; transportador, quantidade e cliente por OCR.
    Se a transportadora não for cadastrada, nem lê o resto (poupa tempo)."""
    chave = grupo["chave"]
    emitente = chave[6:20] if chave else ""
    if not emitente and plan is not None:                  # sem chave: acha, no topo, um CNPJ de remetente cadastrado
        for c in RE_CNPJ.findall(grupo.get("topo", "")):
            d = re.sub(r"\D", "", c)
            if any(re.sub(r"\D", "", r.get("cnpj", "")) == d for r in plan.remetentes.values()):
                emitente = d
                break
    dados = {"arquivo": caminho, "nf": grupo["nf"], "chave": chave, "emitente_cnpj": emitente,
             "cod_cliente": "", "dest": {"nome": "", "cnpj": "", "endereco": "", "bairro": "", "cep": "",
                                          "cidade": "", "uf": "", "fone": ""},
             "bloco_transportador": "", "volumes": None, "emissao": "", "lido_por_ocr": True}
    t = ocr_transporte(caminho, grupo["paginas"][0])
    dados["bloco_transportador"], dados["volumes"] = t["bloco"], t["volumes"]
    if plan is not None and not acha_transportadora(dados["bloco_transportador"], plan.transportadoras):
        return dados
    c = ocr_cliente_fantasia(caminho, grupo["paginas"])
    dados["cod_cliente"] = c.get("cod_cliente", "")
    dados["dest"]["nome"] = c.get("nome", "")
    return dados


# ------------------------------------------------------------------ cruzamento com os cadastros
def identifica(dados, plan):
    """Cruza o DANFE com os cadastros: cliente, remetente, transportadora e volumes (sugestões para conferir)."""
    avisos = []
    dest = dados["dest"]
    doc = re.sub(r"\D", "", dest["cnpj"])
    rem = next((n for n, r in plan.remetentes.items()
                if re.sub(r"\D", "", r.get("cnpj", "")) == dados["emitente_cnpj"]), "")
    if not rem:
        avisos.append("O CNPJ do emitente não está nos remetentes cadastrados: escolha o remetente.")
    transp = acha_transportadora(dados["bloco_transportador"], plan.transportadoras)
    vol = 1
    if dados["volumes"]:
        if 1 <= dados["volumes"] <= 99:
            vol = dados["volumes"]
        else:
            avisos.append(f"A quantidade do DANFE ({dados['volumes']}) é grande demais para volumes: ajuste.")
    cli, origem = None, ""
    if dados["cod_cliente"]:
        cli = plan.cliente(dados["cod_cliente"])
        if cli and doc and re.sub(r"\D", "", cli["cnpj"]) not in ("", doc):
            avisos.append(f"O código {dados['cod_cliente']} tem outro CNPJ no cadastro; procurei pelo CNPJ do DANFE.")
            cli = None
        if cli:
            origem = "código"
    if not cli and doc:
        cli = plan.cliente_por_cnpj(dest["cnpj"])
        origem = "CNPJ" if cli else ""
    if not cli and (dest["nome"] and (doc or dest["endereco"])):
        avisos.append("Cliente fora da base: usei os dados escritos no DANFE.")
        cli = {"cod": dados["cod_cliente"] or "-", "nome": dest["nome"], "fantasia": "", "endereco": dest["endereco"],
               "cep": dest["cep"], "uf": dest["uf"], "cidade": dest["cidade"], "fone": dest["fone"],
               "cnpj": dest["cnpj"], "situacao": ""}
        origem = "DANFE"
    if not cli:
        avisos.append("Não consegui identificar o cliente no DANFE: digite o código ou o nome.")
    return {"nf": dados["nf"], "cliente": cli, "origem_cliente": origem, "remetente": rem, "transportadora": transp,
            "volumes": vol, "volumes_no_danfe": dados["volumes"], "avisos": avisos, "arquivo": dados["arquivo"]}


def le_danfes(caminho, plan, ao_progredir=None):
    """Lê todas as NFs do PDF. Devolve (aceitas, ignoradas): só entram as NFs cuja transportadora está cadastrada."""
    textos = textos_das_paginas(caminho)
    if not any(t.strip() for t in textos) and not textos:
        raise ValueError("O PDF está vazio.")
    grupos = agrupa_paginas(caminho, textos)
    if not grupos:
        raise ValueError("Não encontrei nenhum DANFE nesse PDF.\n"
                         f"Dentro do PDF: {diagnostico(caminho)}\n"
                         "Se não há texto, o ERP está imprimindo como imagem e o código de barras não ficou legível.")
    aceitas, ignoradas, vistas = [], [], set()
    for n, g in enumerate(grupos, 1):
        if ao_progredir:
            ao_progredir(n, len(grupos))
        try:
            dados = (_dados_da_imagem(caminho, g, plan) if g["imagem"]
                     else _dados_do_texto("\n".join(textos[i] for i in g["paginas"]), g["chave"], caminho))
        except ValueError as ex:
            ignoradas.append({"nf": "?", "motivo": str(ex)})
            continue
        if dados["nf"] in vistas:
            continue
        vistas.add(dados["nf"])
        res = identifica(dados, plan)
        if res["transportadora"]:
            aceitas.append(res)
        elif dados.get("lido_por_ocr") and not dados["bloco_transportador"].strip():
            ignoradas.append({"nf": res["nf"], "curto": "imagem ilegível",
                              "motivo": "não consegui ler o bloco do transportador na imagem (qualidade baixa ou leitor indisponível)"})
        else:
            ignoradas.append({"nf": res["nf"], "curto": "transportadora não cadastrada",
                              "motivo": "transportadora não cadastrada ou não informada"})
    return aceitas, ignoradas


def le_danfe(caminho, plan=None):
    """Primeira NF do PDF (sem filtrar pela transportadora). Usado em testes e diagnósticos."""
    textos = textos_das_paginas(caminho)
    grupos = agrupa_paginas(caminho, textos)
    if not grupos:
        raise ValueError("Não encontrei nenhum DANFE nesse PDF.\n" f"Dentro do PDF: {diagnostico(caminho)}")
    g = grupos[0]
    return (_dados_da_imagem(caminho, g, plan) if g["imagem"]
            else _dados_do_texto("\n".join(textos[i] for i in g["paginas"]), g["chave"], caminho))
