# -*- coding: utf-8 -*-
"""Comunicado de avaria: mesmo desenho do PDF, redesenhado no tamanho exato da etiqueta (preto e branco)."""
from PIL import Image, ImageDraw

from etiqueta_core import PRETO, BRANCO, _ajusta, _logo, _quebra, fonte, recurso

TEXTOS = (
    ("caixa", "Abra e confira o produto na presença do motorista."),
    ("proibido", "Constatada qualquer avaria, recuse o recebimento e registre ressalva no verso do CT-e e/ou NF-e."),
    ("balanca", "Conforme arts. 749, 750 e 751 do Código Civil Brasileiro, a transportadora é responsável por "
                "quaisquer danos, inclusive avaria oculta, ocorridos durante o transporte."),
)


def _icone(nome, altura_max, largura_max):
    im = Image.open(recurso("assets", f"avaria_{nome}.png")).convert("L")
    fator = min(altura_max / im.height, largura_max / im.width)
    return im.resize((max(1, int(im.width * fator)), max(1, int(im.height * fator))), Image.LANCZOS)


def desenha_comunicado(w_mm=100, h_mm=150, dpi=203):
    px = dpi / 25.4
    W, H = int(round(w_mm * px)), int(round(h_mm * px))
    k = W / 800.0
    img = Image.new("L", (W, H), BRANCO)
    d = ImageDraw.Draw(img)
    m = 26 * k
    esp = max(3, int(7 * k))
    x0, x1 = m, W - m

    # logo no canto superior direito
    logo = _logo(66 * k)
    y = m
    if logo:
        img.paste(logo, (int(x1 - logo.width), int(y)))
        y += logo.height + 14 * k

    # caixa ATENÇÃO + triângulo
    alt_aten = 168 * k
    d.rounded_rectangle([x0, y, x1, y + alt_aten], radius=36 * k, outline=PRETO, width=esp)
    tri = _icone("triangulo", 112 * k, 150 * k)
    f_aten = _ajusta(d, "ATENÇÃO", "anton", (x1 - x0) * 0.55, 150 * k, 60 * k)
    tw = d.textlength("ATENÇÃO", font=f_aten)
    gap = 44 * k
    gx = (x0 + x1) / 2 - (tw + gap + tri.width) / 2
    d.text((gx, y + alt_aten / 2), "ATENÇÃO", font=f_aten, fill=PRETO, anchor="lm")
    img.paste(tri, (int(gx + tw + gap), int(y + (alt_aten - tri.height) / 2)))
    y += alt_aten + 16 * k

    # subtítulo
    sub = "CONFERÊNCIA OBRIGATÓRIA NO RECEBIMENTO"
    f_sub = _ajusta(d, sub, "anton", x1 - x0, 60 * k, 24 * k)
    d.text(((x0 + x1) / 2, y), sub, font=f_sub, fill=PRETO, anchor="ma")
    y += f_sub.size * 1.25 + 14 * k

    # três avisos: maior letra que couber, sobra de altura distribuída entre as caixas
    gap_c = 16 * k
    pad = 20 * k
    col_icone = 168 * k
    x_txt = x0 + pad + col_icone + 20 * k
    larg_txt = (x1 - pad) - x_txt
    disp = (H - m) - y
    escolha = None
    for tam in range(int(42 * k), int(14 * k) - 1, -1):
        f = fonte("med", tam)
        linhas = [_quebra(d, t, f, larg_txt, 40) for _, t in TEXTOS]
        alt = [max(112 * k, len(l) * tam * 1.3) + 2 * pad for l in linhas]
        if sum(alt) + 2 * gap_c <= disp:
            escolha = (f, linhas, alt)
            break
    if escolha is None:
        f = fonte("med", int(14 * k))
        linhas = [_quebra(d, t, f, larg_txt, 40) for _, t in TEXTOS]
        escolha = (f, linhas, [max(112 * k, len(l) * f.size * 1.3) + 2 * pad for l in linhas])
    f, linhas, alt = escolha
    sobra = max(0, disp - sum(alt) - 2 * gap_c)
    alt = [a + sobra / 3 for a in alt]
    for (nome, _), ls, a in zip(TEXTOS, linhas, alt):
        d.rounded_rectangle([x0, y, x1, y + a], radius=30 * k, outline=PRETO, width=esp)
        ic = _icone(nome, min(a - 2 * pad, 150 * k), col_icone)
        img.paste(ic, (int(x0 + pad + (col_icone - ic.width) / 2), int(y + (a - ic.height) / 2)))
        alt_txt = len(ls) * f.size * 1.3
        ty = y + (a - alt_txt) / 2
        for ln in ls:
            d.text((x_txt, ty), ln, font=f, fill=PRETO)
            ty += f.size * 1.3
        y += a + gap_c
    return img.point(lambda p: 0 if p < 150 else 255).convert("1")
