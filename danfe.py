# -*- coding: utf-8 -*-
"""Lê o DANFE (PDF) e descobre NF, cliente, remetente, transportadora e volumes."""
import re

from etiqueta_core import _sem_acento, fmt_cep, limpa

RE_CNPJ = re.compile(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b|\b\d{3}\.\d{3}\.\d{3}-\d{2}\b")


def texto_do_pdf(caminho):
    from pypdf import PdfReader
    leitor = PdfReader(caminho)
    texto = "\n".join((p.extract_text() or "") for p in leitor.pages[:3])
    if len(texto.strip()) < 50:
        raise ValueError("O PDF não tem texto legível (parece uma imagem). Gere o DANFE em PDF com texto.")
    return texto


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


def _linhas(texto):
    return [l.strip() for l in texto.splitlines()]


def _norm(t):
    return _sem_acento(t).upper().strip()


def _valor_apos(linhas, rotulo, ocorrencia=0, inicio=0):
    """Texto da linha seguinte ao rótulo (ex.: 'NOME / RAZÃO SOCIAL')."""
    alvo, achadas = _norm(rotulo), 0
    for i in range(inicio, len(linhas)):
        if _norm(linhas[i]) == alvo:
            if achadas == ocorrencia:
                for j in range(i + 1, min(i + 4, len(linhas))):
                    if linhas[j]:
                        return linhas[j]
                return ""
            achadas += 1
    return ""


def le_danfe(caminho):
    """Devolve um dict com o que dá para ler do DANFE. Levanta ValueError se não for um DANFE."""
    texto = texto_do_pdf(caminho)
    chave = acha_chave(texto)
    if not chave and "DANFE" not in texto.upper():
        raise ValueError("Esse PDF não parece ser um DANFE.")
    linhas = _linhas(texto)
    nf = str(int(chave[25:34])) if chave else ""
    if not nf:
        m = re.search(r"N[ºo°]\s*([\d.]{5,11})", texto)
        nf = str(int(re.sub(r"\D", "", m.group(1)))) if m else ""
    if not nf:
        raise ValueError("Não encontrei o número da NF no DANFE.")
    emit_cnpj = chave[6:20] if chave else re.sub(r"\D", "", _valor_apos(linhas, "CNPJ"))
    cod = re.search(r"Cliente\s+Fantasia\s*:\s*(\d+)", texto, re.I)
    # destinatário (campos com rótulo)
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
    # bloco do transportador / volumes
    i = texto.upper().find("TRANSPORTADOR / VOLUMES")
    j = texto.upper().find("DADOS DOS PRODUTOS", i + 1) if i >= 0 else -1
    bloco = texto[i:j] if i >= 0 and j > i else ""
    qtd = None
    bl = _linhas(bloco)
    q = _valor_apos(bl, "QUANTIDADE")
    if re.fullmatch(r"\d{1,6}", q or ""):
        qtd = int(q)
    emissao = re.search(r"Emiss[ãa]o:\s*(\d{2}/\d{2}/\d{4})", texto)
    return {
        "arquivo": caminho, "nf": nf, "serie": str(int(chave[22:25])) if chave else "", "chave": chave,
        "emitente_cnpj": re.sub(r"\D", "", emit_cnpj), "cod_cliente": cod.group(1) if cod else "",
        "dest": dest, "bloco_transportador": bloco, "volumes": qtd, "emissao": emissao.group(1) if emissao else "",
    }


def identifica(dados, plan):
    """Cruza o DANFE com os cadastros: cliente, remetente, transportadora e volumes (sugestões para o usuário conferir)."""
    avisos = []
    dest = dados["dest"]
    doc = re.sub(r"\D", "", dest["cnpj"])
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
    if not cli:
        avisos.append("Cliente fora da base: usei os dados escritos no DANFE.")
        end = dest["endereco"]
        cli = {"cod": dados["cod_cliente"] or "-", "nome": dest["nome"], "fantasia": "", "endereco": end,
               "cep": dest["cep"], "uf": dest["uf"], "cidade": dest["cidade"], "fone": dest["fone"],
               "cnpj": dest["cnpj"], "situacao": ""}
        origem = "DANFE"
    rem = next((n for n, r in plan.remetentes.items()
                if re.sub(r"\D", "", r.get("cnpj", "")) == dados["emitente_cnpj"]), "")
    if not rem:
        avisos.append("O CNPJ do emitente não está nos remetentes cadastrados: escolha o remetente.")
    bloco = _norm(dados["bloco_transportador"])
    transp = next((t for t in sorted(plan.transportadoras, key=len, reverse=True) if _norm(t) in bloco), "")
    vol = dados["volumes"] if dados["volumes"] and 1 <= dados["volumes"] <= 99 else 1
    return {"nf": dados["nf"], "cliente": cli, "origem_cliente": origem, "remetente": rem, "transportadora": transp,
            "volumes": vol, "volumes_no_danfe": dados["volumes"], "avisos": avisos, "arquivo": dados["arquivo"]}
