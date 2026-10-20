# -*- coding: utf-8 -*-
"""Leitor de texto de imagens (OCR) para DANFE impresso como imagem. Usa o RapidOCR (offline).

Cuidados: imagens pequenas são ampliadas (texto miúdo) e as grandes reduzidas, para não estourar a memória;
qualquer erro do OCR vira 'não achei' (nunca derruba o programa).
"""
_motor = None
LARGURA_ALVO = 1300     # largura (px) a que cada faixa da página é levada antes de ler


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
        # det_limit_type=max: nunca amplia dentro do motor (a ampliação é nossa, controlada); sem classificador de ângulo;
        # poucas threads = menos memória.
        _motor = RapidOCR(det_limit_type="max", det_limit_side_len=1400, intra_op_num_threads=2,
                          inter_op_num_threads=1, use_cls=False)
    return _motor


def _executa(recorte, fator):
    import numpy as np
    from PIL import Image
    im = recorte
    if abs(fator - 1) > 0.02:
        im = recorte.resize((max(32, int(recorte.width * fator)), max(32, int(recorte.height * fator))), Image.LANCZOS)
    resultado, _ = _motor_ocr()(np.array(im))
    return resultado or []


def ler_faixa(img, y0, y1):
    """OCR de uma faixa horizontal da página (frações da altura). Devolve caixas com coordenadas da página.
    Em caso de erro (ex.: memória), tenta de novo menor; se falhar, devolve lista vazia."""
    W, H = img.size
    topo, base = int(H * y0), int(H * y1)
    recorte = img.crop((0, topo, W, base)).convert("RGB")
    fator = min(2.0, max(0.6, LARGURA_ALVO / max(1, recorte.width)))
    resultado = None
    for tentativa in (fator, fator * 0.65, fator * 0.45):
        try:
            resultado = _executa(recorte, tentativa)
            fator = tentativa
            break
        except Exception:
            continue
    if resultado is None:
        return []
    caixas = []
    for box, texto, score in resultado:
        xs, ys = [p[0] / fator for p in box], [p[1] / fator for p in box]
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


def aquece():
    """Carrega o motor e faz uma leitura de mentira, para o primeiro DANFE não pagar a partida."""
    try:
        from PIL import Image
        _executa(Image.new("RGB", (320, 96), "white"), 1.0)
    except Exception:
        pass
