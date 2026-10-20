# -*- coding: utf-8 -*-
"""Leitor de texto de imagens (OCR) para DANFE impresso como imagem. Usa o RapidOCR (offline)."""
_motor = None


def disponivel():
    try:
        import rapidocr_onnxruntime  # noqa: F401
        return True
    except Exception:
        return False


def _motor_ocr():
    global _motor
    if _motor is None:
        from rapidocr_onnxruntime import RapidOCR
        _motor = RapidOCR()
    return _motor


def ler_faixa(img, y0, y1):
    """OCR de uma faixa horizontal da página (frações da altura). Devolve caixas com coordenadas da página."""
    import numpy as np
    W, H = img.size
    topo, base = int(H * y0), int(H * y1)
    recorte = img.crop((0, topo, W, base)).convert("RGB")
    resultado, _ = _motor_ocr()(np.array(recorte))
    caixas = []
    for box, texto, score in (resultado or []):
        xs, ys = [p[0] for p in box], [p[1] for p in box]
        caixas.append({"x0": min(xs), "x1": max(xs), "y0": min(ys) + topo, "y1": max(ys) + topo,
                       "texto": texto.strip(), "score": float(score)})
    return caixas


def linhas_de_texto(caixas):
    """Junta caixas na mesma altura em linhas (esquerda -> direita, cima -> baixo)."""
    linhas = []
    for c in sorted(caixas, key=lambda c: (c["y0"] + c["y1"]) / 2):
        cy, alt = (c["y0"] + c["y1"]) / 2, max(1, c["y1"] - c["y0"])
        for l in linhas:
            if abs(l["cy"] - cy) <= alt * 0.55:
                l["itens"].append(c)
                break
        else:
            linhas.append({"cy": cy, "itens": [c]})
    return [" ".join(i["texto"] for i in sorted(l["itens"], key=lambda i: i["x0"])) for l in linhas]
