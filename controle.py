# -*- coding: utf-8 -*-
"""Controle remoto: permite desativar o programa lendo um JSON publicado por você.

Formato do JSON (todos os campos são opcionais):
{
  "bloqueado": false,        // true = desativa o programa para todo mundo
  "mensagem": "",            // texto mostrado ao usuário (no bloqueio ou como aviso)
  "bloquear": [],            // nomes de computador ou usuário do Windows a bloquear. Ex.: ["PC-EXPEDICAO", "maria"]
  "versaoMinima": ""         // ex.: "1.2.0" -> versões mais antigas são bloqueadas
}
"""
import getpass
import json
import os
import socket
import time
import urllib.request

from etiqueta_core import pasta_dados

APP_VERSAO = "1.0.0"

# >>> COLOQUE AQUI o endereço (https) do seu JSON. Vazio = controle desligado.
# Exemplo: "https://raw.githubusercontent.com/SEU_USUARIO/SEU_REPO/main/etiquetas.json"
URL_CONTROLE = ""

MSG_PADRAO = "Este programa foi desativado. Fale com o responsável."
_CACHE = "controle_cache.json"


def _versao(v):
    partes = []
    for p in str(v).strip().lower().lstrip("v").split("."):
        digitos = "".join(c for c in p if c.isdigit())
        partes.append(int(digitos) if digitos else 0)
    return tuple(partes)


def _verdadeiro(v):
    return v is True or str(v).strip().lower() in ("true", "sim", "1")


def decide(dados, host=None, usuario=None, versao=APP_VERSAO):
    """Devolve (bloqueado, mensagem) a partir do JSON."""
    if not isinstance(dados, dict):
        return False, ""
    msg = str(dados.get("mensagem") or "").strip()
    host = (host or socket.gethostname()).strip().lower()
    usuario = (usuario or _usuario()).strip().lower()
    lista = [str(x).strip().lower() for x in (dados.get("bloquear") or []) if str(x).strip()]
    if _verdadeiro(dados.get("bloqueado")):
        return True, msg or MSG_PADRAO
    if host in lista or usuario in lista or f"{usuario}@{host}" in lista:
        return True, msg or MSG_PADRAO
    minima = str(dados.get("versaoMinima") or "").strip()
    if minima and _versao(versao) < _versao(minima):
        return True, msg or f"Esta versão ({versao}) está desatualizada. Instale a versão {minima} ou mais nova."
    return False, msg


def _usuario():
    try:
        return getpass.getuser()
    except Exception:
        return ""


def _le_cache():
    try:
        with open(os.path.join(pasta_dados(), _CACHE), encoding="utf-8") as f:
            return json.load(f).get("dados")
    except Exception:
        return None


def _grava_cache(dados):
    try:
        with open(os.path.join(pasta_dados(), _CACHE), "w", encoding="utf-8") as f:
            json.dump({"dados": dados, "quando": time.time()}, f, ensure_ascii=False)
    except Exception:
        pass


def verifica(url=None, timeout=3):
    """Consulta o JSON. Nunca dá erro: sem internet usa a última resposta guardada
    (se ela dizia 'bloqueado', continua bloqueado); sem nenhuma resposta, libera."""
    url = URL_CONTROLE if url is None else url
    if not url:
        return False, ""
    try:
        sep = "&" if "?" in url else "?"
        req = urllib.request.Request(f"{url}{sep}t={int(time.time())}",
                                     headers={"Cache-Control": "no-cache", "User-Agent": "EtiquetasFreedom"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            dados = json.loads(r.read().decode("utf-8-sig"))
        if not isinstance(dados, dict):
            raise ValueError("JSON inválido")
        _grava_cache(dados)
        return decide(dados)
    except Exception:
        dados = _le_cache()
        return decide(dados) if dados else (False, "")
