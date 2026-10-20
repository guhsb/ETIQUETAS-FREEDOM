# -*- coding: utf-8 -*-
"""Modo rápido para impressoras Zebra (ou compatíveis com ZPL): manda a etiqueta pronta direto, sem passar pelo driver gráfico."""
from PIL import ImageOps


def etiqueta_zpl(img, copias=1):
    """Imagem 1 bit (preto = 0) -> comando ZPL (^GFA, hexadecimal) no tamanho exato da imagem."""
    W, H = img.size
    pb = ImageOps.invert(img.convert("L")).point(lambda p: 255 if p > 127 else 0).convert("1")   # bit 1 = ponto preto
    linha = (W + 7) // 8
    dados = pb.tobytes().hex().upper()
    total = linha * H
    return (f"^XA^CI28^PW{W}^LL{H}^LH0,0^PQ{max(1, int(copias))}"
            f"^FO0,0^GFA,{total},{total},{linha},{dados}^FS^XZ\n").encode("ascii")


def trabalho_zpl(imagens):
    return b"".join(etiqueta_zpl(i) for i in imagens)


def decodifica(zpl_bytes):
    """Só para testes: volta do ZPL (^GFA) para a imagem."""
    from PIL import Image
    import re
    out = []
    for m in re.finditer(rb"\^PW(\d+)\^LL(\d+).*?\^GFA,(\d+),\d+,(\d+),([0-9A-F]+)\^FS", zpl_bytes, re.S):
        W, H, total, linha, dados = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4)), m.group(5)
        raw = bytes.fromhex(dados.decode())
        assert len(raw) == total == linha * H
        im = Image.frombytes("1", (linha * 8, H), raw).crop((0, 0, W, H))
        out.append(ImageOps.invert(im.convert("L")).point(lambda p: 255 if p > 127 else 0).convert("1"))
    return out
