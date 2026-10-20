# -*- coding: utf-8 -*-
"""Leitura da planilha + desenho da etiqueta (preto e branco, pronta p/ térmica)."""
import json
import os
import re
import sys

import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# ------------------------------------------------------------------ recursos
def recurso(*partes):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *partes)


FONTES = {
    "reg": "Roboto-Regular.ttf",
    "med": "Roboto-Medium.ttf",
    "bold": "Roboto-Bold.ttf",
    "black": "Roboto-Black.ttf",
    "anton": "Anton-Regular.ttf",
}
_cache_fontes = {}


def fonte(estilo, tam):
    chave = (estilo, int(tam))
    if chave not in _cache_fontes:
        try:
            _cache_fontes[chave] = ImageFont.truetype(recurso("fonts", FONTES[estilo]), int(tam))
        except OSError:
            for alt in ("arialbd.ttf", "arial.ttf", "DejaVuSans-Bold.ttf"):
                try:
                    _cache_fontes[chave] = ImageFont.truetype(alt, int(tam))
                    break
                except OSError:
                    continue
            else:
                _cache_fontes[chave] = ImageFont.load_default()
    return _cache_fontes[chave]


# ------------------------------------------------------------------ dados
def fmt_cep(v):
    d = re.sub(r"\D", "", str(v or ""))
    if not d:
        return ""
    d = d.zfill(8)[-8:]
    return f"{d[:5]}-{d[5:]}"


def limpa(v):
    v = "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)
    return re.sub(r"\s+", " ", v).strip()


CLI_COLS = ["CODCLIENTE", "NOME", "FANTASIA", "ENDERECO", "NUMERO", "CEP", "CIDADE", "UF", "FONE1", "CNPJ", "SITUACAO"]
ALIAS = {"CODIGO": "CODCLIENTE", "COD": "CODCLIENTE", "CODIGO CLIENTE": "CODCLIENTE", "COD CLIENTE": "CODCLIENTE",
         "RAZAO SOCIAL": "NOME", "RAZAO": "NOME", "END": "ENDERECO", "LOGRADOURO": "ENDERECO",
         "TELEFONE": "FONE1", "FONE": "FONE1", "MUNICIPIO": "CIDADE", "ESTADO": "UF"}


def pasta_exe():
    """Pasta onde está o .exe (ou o código, durante o desenvolvimento)."""
    return os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))


def _gravavel(pasta):
    try:
        os.makedirs(pasta, exist_ok=True)
        teste = os.path.join(pasta, ".teste_gravacao")
        with open(teste, "w") as f:
            f.write("ok")
        os.remove(teste)
        return True
    except OSError:
        return False


_cache_pastas = {}


def pasta_dados():
    """Dados do programa (clientes, cadastros, histórico, configuração): pasta 'dados' ao lado do .exe.
    Se a pasta do .exe não aceitar gravação (ex.: Arquivos de Programas), usa AppData."""
    if "dados" not in _cache_pastas:
        p = os.path.join(pasta_exe(), "dados")
        if not _gravavel(p):
            p = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "EtiquetasFreedom")
            os.makedirs(p, exist_ok=True)
        _cache_pastas["dados"] = p
    return _cache_pastas["dados"]


def pasta_saida():
    """Onde fica o Excel do histórico: ao lado do .exe (ou na pasta de dados se não der para gravar)."""
    return pasta_exe() if _gravavel(pasta_exe()) else pasta_dados()


def _sem_acento(t):
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFD", str(t)) if unicodedata.category(c) != "Mn")


def _le_bruto(caminho):
    """Lê CSV (; ou , — UTF-8 ou Latin-1) ou Excel exportado do sistema."""
    if caminho.lower().endswith((".xlsx", ".xlsm")):
        df = pd.read_excel(caminho, dtype=str)
    else:
        raw = open(caminho, "rb").read()
        try:
            texto = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            texto = raw.decode("latin-1")
        import io
        primeira = texto.split("\n", 1)[0]
        sep = ";" if primeira.count(";") >= primeira.count(",") else ","
        df = pd.read_csv(io.StringIO(texto), sep=sep, dtype=str, low_memory=False)
    df = df.fillna("")
    df.columns = [ALIAS.get(_sem_acento(c).strip().upper(), _sem_acento(c).strip().upper()) for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]
    falta = [c for c in ("CODCLIENTE", "NOME") if c not in df.columns]
    if falta:
        raise ValueError("Não achei a coluna " + " e ".join(falta) + " na planilha de clientes.")
    for c in CLI_COLS:
        if c not in df.columns:
            df[c] = ""
    df = df[CLI_COLS].copy()
    for c in CLI_COLS:
        df[c] = df[c].map(limpa)
    df["CODCLIENTE"] = df["CODCLIENTE"].str.replace(r"\.0$", "", regex=True)
    return df[df["CODCLIENTE"] != ""]


def ler_tabela(caminho):
    """Lê o arquivo do sistema. Se o mesmo código aparecer mais de uma vez, vale a última linha."""
    bruto = _le_bruto(caminho)
    return bruto.drop_duplicates("CODCLIENTE", keep="last")[CLI_COLS].copy()


def monta_endereco(end, num):
    end, num = limpa(end), limpa(num)
    if num.upper() in ("", "0", "S/N", "SN", "S/Nº", "SEM NUMERO"):
        return end
    if re.search(r"(?<!\d)" + re.escape(num) + r"(?!\d)", end):
        return end  # o número já está escrito no endereço
    return f"{end}, {num}"


CADASTROS = "cadastros.json"


def fmt_cnpj(v):
    d = re.sub(r"\D", "", str(v or ""))
    if len(d) == 14:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
    return limpa(v)


def _chave(t):
    return _sem_acento(limpa(t)).upper()


class Planilha:
    """Tudo que o programa guarda: clientes (base vigiada) + remetentes e transportadoras (cadastros)."""

    def __init__(self):
        self.base = os.path.join(pasta_dados(), "clientes.csv")
        self.arq_cad = os.path.join(pasta_dados(), CADASTROS)
        self._carrega_cadastros()
        cli = ler_tabela(self.base) if os.path.exists(self.base) else pd.DataFrame(columns=CLI_COLS)
        self._define_clientes(cli)
        if len(cli) == 0:
            # primeira vez: usa o clientes.csv ao lado do .exe; se não houver, a base que vem dentro do programa
            for c in (os.path.join(pasta_exe(), "clientes.csv"), os.path.join(pasta_exe(), "clientes.xlsx"),
                      recurso("clientes.csv")):
                if os.path.exists(c):
                    try:
                        self.importar_clientes(c)
                        break
                    except Exception:
                        continue

    # ---------- cadastros (remetentes / transportadoras)
    def _carrega_cadastros(self):
        dados = None
        for arq in (self.arq_cad, recurso("cadastros_padrao.json")):
            try:
                with open(arq, encoding="utf-8") as f:
                    dados = json.load(f)
                break
            except (OSError, ValueError):
                continue
        dados = dados or {}
        self.remetentes = {}
        for r in dados.get("remetentes", []):
            nome = limpa(r.get("nome"))
            if nome:
                self.remetentes[nome] = dict(
                    nome=nome, endereco=limpa(r.get("endereco")), cep=fmt_cep(r.get("cep")),
                    uf=limpa(r.get("uf")).upper()[:2], cidade=limpa(r.get("cidade")),
                    fone=limpa(r.get("fone")), cnpj=fmt_cnpj(r.get("cnpj")))
        self.transportadoras = []
        for t in dados.get("transportadoras", []):
            t = limpa(t)
            if t and _chave(t) not in {_chave(x) for x in self.transportadoras}:
                self.transportadoras.append(t)
        if not os.path.exists(self.arq_cad):
            self.salva_cadastros()

    def salva_cadastros(self):
        tmp = self.arq_cad + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"remetentes": list(self.remetentes.values()), "transportadoras": self.transportadoras},
                      f, ensure_ascii=False, indent=1)
        os.replace(tmp, self.arq_cad)

    def add_transportadora(self, nome):
        """Cria a transportadora (se ainda não existir). Devolve o nome oficial."""
        nome = limpa(nome).upper()
        if not nome:
            return ""
        for t in self.transportadoras:
            if _chave(t) == _chave(nome):
                return t
        self.transportadoras.append(nome)
        self.salva_cadastros()
        return nome

    def renomear_transportadora(self, antigo, novo):
        novo = limpa(novo).upper()
        if not novo or antigo not in self.transportadoras:
            return ""
        if any(_chave(t) == _chave(novo) and t != antigo for t in self.transportadoras):
            raise ValueError(f"Já existe a transportadora {novo}.")
        self.transportadoras[self.transportadoras.index(antigo)] = novo
        self.salva_cadastros()
        return novo

    def remove_transportadora(self, nome):
        if nome in self.transportadoras:
            self.transportadoras.remove(nome)
            self.salva_cadastros()

    def salva_remetente(self, dados, original=None):
        """Cria ou altera um remetente. 'original' = nome antes da edição (para renomear)."""
        nome = limpa(dados.get("nome")).upper()
        if not nome:
            raise ValueError("Informe o nome da empresa.")
        if any(_chave(n) == _chave(nome) and n != original for n in self.remetentes):
            raise ValueError(f"Já existe o remetente {nome}.")
        reg = dict(nome=nome, endereco=limpa(dados.get("endereco")).upper(), cep=fmt_cep(dados.get("cep")),
                   uf=limpa(dados.get("uf")).upper()[:2], cidade=limpa(dados.get("cidade")).upper(),
                   fone=limpa(dados.get("fone")), cnpj=fmt_cnpj(dados.get("cnpj")))
        if original and original in self.remetentes:  # mantém a posição na lista
            self.remetentes = {(nome if k == original else k): (reg if k == original else v)
                               for k, v in self.remetentes.items()}
        else:
            self.remetentes[nome] = reg
        self.salva_cadastros()
        return nome

    def remove_remetente(self, nome):
        if nome in self.remetentes:
            del self.remetentes[nome]
            self.salva_cadastros()

    # ---------- clientes
    def _define_clientes(self, cli):
        cli = cli.copy()
        cli["_busca"] = (cli["CODCLIENTE"] + " " + cli["NOME"] + " " + cli["FANTASIA"] + " " + cli["CIDADE"]).map(
            lambda t: _sem_acento(t).upper())
        self.clientes = cli.set_index("CODCLIENTE", drop=False)

    def importar_clientes(self, caminho, atualizar=False):
        return self.importar_df(ler_tabela(caminho), atualizar)

    def importar_df(self, novo, atualizar=False):
        """Arquivo cheio do sistema: quem já está cadastrado é IGNORADO e só os clientes novos entram.
        (atualizar=True também altera os dados de quem já existe.) Devolve o resumo."""
        atual = self.clientes[CLI_COLS]
        ja = novo["CODCLIENTE"].isin(set(atual.index))
        so_novos = novo[~ja]
        alterados = 0
        if atualizar and ja.any():
            comuns = novo[ja]
            a = atual.loc[list(comuns["CODCLIENTE"])].reset_index(drop=True)
            alterados = int((a.values != comuns.reset_index(drop=True)[CLI_COLS].values).any(axis=1).sum())
        if len(so_novos) or alterados or not os.path.exists(self.base):
            junto = pd.concat([atual.reset_index(drop=True), novo if atualizar else so_novos])
            junto = junto.drop_duplicates("CODCLIENTE", keep="last")
            junto = junto.sort_values("CODCLIENTE", key=lambda s: pd.to_numeric(s, errors="coerce"))
            junto.to_csv(self.base, sep=";", index=False, encoding="utf-8-sig")
            self._define_clientes(junto)
        return {"novos": len(so_novos), "ignorados": int(ja.sum()), "alterados": alterados, "total": len(self.clientes)}

    def cliente(self, cod):
        cod = limpa(cod)
        cod = cod.lstrip("0") or cod
        if cod in self.clientes.index:
            r = self.clientes.loc[cod]
            return dict(
                cod=cod, nome=limpa(r["NOME"]), fantasia=limpa(r["FANTASIA"]),
                endereco=monta_endereco(r["ENDERECO"], r["NUMERO"]),
                cep=fmt_cep(r["CEP"]), uf=limpa(r["UF"]), cidade=limpa(r["CIDADE"]),
                fone=limpa(r["FONE1"]), cnpj=limpa(r["CNPJ"]), situacao=limpa(r["SITUACAO"]).upper(),
            )
        return None

    def buscar(self, texto, limite=8):
        t = _sem_acento(limpa(texto)).upper()
        if not t:
            return []
        c = self.clientes
        exato = c[c["CODCLIENTE"] == t.lstrip("0")]
        resto = c[c["_busca"].str.contains(re.escape(t), regex=True)]
        res = pd.concat([exato, resto]).drop_duplicates("CODCLIENTE").head(limite)
        def nome(r):
            n = r.NOME if not r.FANTASIA or r.FANTASIA == r.NOME else f"{r.NOME} ({r.FANTASIA})"
            return n + ("  [INATIVO]" if r.SITUACAO.upper() == "I" else "")
        return [(r.CODCLIENTE, nome(r), r.CIDADE, r.UF) for r in res.itertuples()]


# ------------------------------------------------------------------ desenho
PRETO, BRANCO = 0, 255
INFO_ADICIONAL = [
    "Nota fiscal e boleto estão dentro da caixa",
    "Não receber a mercadoria se a caixa estiver violada",
    "Telefone para contato para possíveis dúvidas/problemas",
    "(41) 9 9258-7982 (WhatsApp e Telefone)",
]


def _largura(d, txt, f):
    return d.textlength(txt, font=f)


def _ajusta(d, txt, estilo, largura, tam_max, tam_min):
    tam = tam_max
    while tam > tam_min and _largura(d, txt, fonte(estilo, tam)) > largura:
        tam -= 2
    return fonte(estilo, tam)


def _quebra(d, txt, f, largura, max_linhas):
    palavras, linhas, atual = txt.split(), [], ""
    for p in palavras:
        teste = (atual + " " + p).strip()
        if _largura(d, teste, f) <= largura or not atual:
            atual = teste
        else:
            linhas.append(atual)
            atual = p
    if atual:
        linhas.append(atual)
    if len(linhas) > max_linhas:
        linhas = linhas[:max_linhas]
        while _largura(d, linhas[-1] + "…", f) > largura and len(linhas[-1]) > 1:
            linhas[-1] = linhas[-1][:-1]
        linhas[-1] += "…"
    return linhas


def _rotulo(d, x, y, txt, k):
    """Pílula preta com texto branco (rótulo de seção)."""
    f = fonte("bold", 20 * k)
    w = _largura(d, txt, f) + 28 * k
    h = 34 * k
    d.rounded_rectangle([x, y, x + w, y + h], radius=h / 2, fill=PRETO)
    d.text((x + 14 * k, y + h / 2), txt, font=f, fill=BRANCO, anchor="lm")
    return h


def _logo(altura):
    try:
        im = Image.open(recurso("logo.png")).convert("RGBA")
    except OSError:
        return None
    fundo = Image.new("RGBA", im.size, (255, 255, 255, 255))
    fundo.alpha_composite(im)
    g = fundo.convert("L")
    # realça o contraste: o rosa vira cinza escuro -> preto na térmica
    g = g.point(lambda p: 0 if p < 170 else 255)
    w = int(g.width * altura / g.height)
    return g.resize((w, int(altura)), Image.LANCZOS).point(lambda p: 0 if p < 200 else 255)


def desenha_etiqueta(rem, cli, transportadora, nf, volume, total,
                     w_mm=100, h_mm=150, dpi=203):
    px = dpi / 25.4
    W, H = int(round(w_mm * px)), int(round(h_mm * px))
    k = W / 800.0  # tudo é projetado em 800 px de largura
    esp = max(3, int(6 * k))  # espessura das linhas
    img = Image.new("L", (W, H), BRANCO)
    d = ImageDraw.Draw(img)
    m = 22 * k

    # moldura externa grossa
    d.rounded_rectangle([m / 2, m / 2, W - m / 2, H - m / 2], radius=24 * k, outline=PRETO, width=esp)
    x0, x1 = m + 14 * k, W - m - 14 * k
    larg = x1 - x0
    y = m + 12 * k

    # ---- cabeçalho: logo grande + remetente (3 linhas maiores)
    alt_cab = 138 * k
    logo = _logo(124 * k)
    if logo:
        img.paste(logo, (int(x0), int(y)))
        xr = x0 + logo.width + 22 * k
    else:
        xr = x0
    disp = x1 - xr
    d.text((xr, y), "REMETENTE", font=fonte("bold", 19 * k), fill=PRETO)
    f = _ajusta(d, rem["nome"], "black", disp, 34 * k, 20 * k)
    d.text((xr, y + 20 * k), rem["nome"], font=f, fill=PRETO)
    for i, t in enumerate((rem["endereco"], f'{rem["cidade"]}/{rem["uf"]} · CEP {rem["cep"]}', f'CNPJ {rem["cnpj"]}')):
        d.text((xr, y + (58 + i * 27) * k), t, font=_ajusta(d, t, "bold", disp, 24 * k, 14 * k), fill=PRETO)
    y += alt_cab + 12 * k
    d.line([x0, y, x1, y], fill=PRETO, width=esp)
    y += 16 * k

    # ---- destinatário
    h = _rotulo(d, x0, y, "DESTINATÁRIO", k)
    d.text((x1, y + h / 2), f'COD {cli["cod"]}', font=fonte("black", 36 * k), fill=PRETO, anchor="rm")
    y += h + 12 * k

    nome = cli["nome"].upper()
    for tam in (60, 52, 44, 38):
        fn = fonte("black", tam * k)
        linhas = _quebra(d, nome, fn, larg, 2)
        if not linhas[-1].endswith("…"):
            break
    for ln in linhas:
        d.text((x0, y), ln, font=fn, fill=PRETO)
        y += fn.size * 1.1
    y += 6 * k
    fe = fonte("bold", 35 * k)
    for ln in _quebra(d, cli["endereco"].upper(), fe, larg, 2):
        d.text((x0, y), ln, font=fe, fill=PRETO)
        y += 42 * k
    y += 4 * k
    cidade = f'{cli["cidade"].upper()} / {cli["uf"]}'
    fcid = _ajusta(d, cidade, "black", larg, 62 * k, 32 * k)
    d.text((x0, y), cidade, font=fcid, fill=PRETO)
    y += fcid.size * 1.12 + 2 * k
    d.text((x0, y), f'CEP {cli["cep"]}', font=fonte("black", 56 * k), fill=PRETO)
    y += 66 * k
    for t in (f'Fone: {cli["fone"]}' if cli["fone"] else "", f'CNPJ: {cli["cnpj"]}'):
        if t:
            d.text((x0, y), t, font=_ajusta(d, t, "bold", larg, 32 * k, 18 * k), fill=PRETO)
            y += 40 * k
    y += 2 * k

    # ---- volume / NF (bem grandes)
    d.line([x0, y, x1, y], fill=PRETO, width=esp)
    y += 14 * k
    bh = 214 * k
    meio = x0 + larg * 0.50
    xv1 = meio - 8 * k
    d.rounded_rectangle([x0, y, xv1, y + bh], radius=18 * k, fill=PRETO)
    cxv = (x0 + xv1) / 2
    d.text((cxv, y + 29 * k), "VOLUME", font=fonte("bold", 27 * k), fill=BRANCO, anchor="mm")
    txt_vol = f"{volume}/{total}"
    fv = _ajusta(d, txt_vol, "black", (xv1 - x0) - 36 * k, 150 * k, 70 * k)
    d.text((cxv, y + 124 * k), txt_vol, font=fv, fill=BRANCO, anchor="mm")
    d.rounded_rectangle([meio + 8 * k, y, x1, y + bh], radius=18 * k, outline=PRETO, width=esp)
    cx = (meio + 8 * k + x1) / 2
    d.text((cx, y + 29 * k), "NOTA FISCAL", font=fonte("bold", 27 * k), fill=PRETO, anchor="mm")
    nf_txt = str(nf)
    fnf = _ajusta(d, nf_txt, "black", (x1 - meio) - 44 * k, 100 * k, 40 * k)
    d.text((cx, y + 124 * k), nf_txt, font=fnf, fill=PRETO, anchor="mm")
    y += bh + 10 * k

    # ---- aviso no rodapé (grande), ancorado na base
    base = H - m - 14 * k
    linhas_aviso, fone_aviso = INFO_ADICIONAL[:-1], INFO_ADICIONAL[-1]
    tam = min(_ajusta(d, t, "bold", larg, 30 * k, 14 * k).size for t in linhas_aviso)
    fa = fonte("bold", tam)
    ffone = _ajusta(d, fone_aviso, "black", larg, 40 * k, 20 * k)
    alt_aviso = len(linhas_aviso) * fa.size * 1.28 + ffone.size * 1.3
    y_info = base - alt_aviso
    d.line([x0, y_info - 12 * k, x1, y_info - 12 * k], fill=PRETO, width=esp)
    yy = y_info
    for t in linhas_aviso:
        d.text((W / 2, yy), t, font=fa, fill=PRETO, anchor="ma")
        yy += fa.size * 1.28
    d.text((W / 2, yy + 2 * k), fone_aviso, font=ffone, fill=PRETO, anchor="ma")

    # ---- transportadora centralizada, no meio do espaço que sobrar
    ft = _ajusta(d, transportadora.upper(), "black", larg, 50 * k, 30 * k)
    alt_t = 28 * k + ft.size * 1.15
    topo, fim = y, y_info - 20 * k
    yt = topo + max(0, (fim - topo - alt_t) / 2)
    d.text((W / 2, yt), "TRANSPORTADORA", font=fonte("bold", 24 * k), fill=PRETO, anchor="ma")
    d.text((W / 2, yt + 28 * k), transportadora.upper(), font=ft, fill=PRETO, anchor="ma")

    return img.point(lambda p: 0 if p < 150 else 255).convert("1")
