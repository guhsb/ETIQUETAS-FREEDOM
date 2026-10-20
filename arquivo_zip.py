# -*- coding: utf-8 -*-
"""Guarda os DANFEs lidos em um ZIP por mês/ano (ex.: 10-2026.zip), na pasta do programa, para ocupar menos espaço."""
import datetime
import os
import zipfile


def caminho_zip(pasta, quando=None):
    q = quando or datetime.datetime.now()
    return os.path.join(pasta, f"{q:%m-%Y}.zip")


def guarda(pdf, nome, pasta, quando=None, apagar=False):
    """Acrescenta o PDF ao ZIP do mês. Devolve o caminho do ZIP. Com apagar=True remove o PDF solto depois de guardar.
    Levanta OSError se o ZIP estiver aberto/travado (quem chamou decide o que fazer)."""
    zp = caminho_zip(pasta, quando)
    with zipfile.ZipFile(zp, "a", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        existentes = set(z.namelist())
        base, ext = os.path.splitext(nome)
        n, i = nome, 2
        while n in existentes:
            n, i = f"{base}_{i}{ext}", i + 1
        z.write(pdf, n)
    with zipfile.ZipFile(zp) as z:                       # confere que entrou antes de apagar o original
        z.getinfo(n)
    if apagar:
        os.remove(pdf)
    return zp


def guarda_pendentes(pastas, destino):
    """PDFs soltos que ficaram para trás (ex.: ZIP estava aberto): vão para o ZIP do mês do próprio arquivo."""
    guardados = 0
    for pasta in pastas:
        if not os.path.isdir(pasta):
            continue
        for f in sorted(os.listdir(pasta)):
            if not f.lower().endswith(".pdf"):
                continue
            c = os.path.join(pasta, f)
            try:
                guarda(c, f, destino, datetime.datetime.fromtimestamp(os.path.getmtime(c)), apagar=True)
                guardados += 1
            except OSError:
                break          # ZIP travado agora: tenta de novo na próxima vez
    return guardados
