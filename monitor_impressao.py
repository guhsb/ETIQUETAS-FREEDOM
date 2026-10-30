# -*- coding: utf-8 -*-
"""Acompanha o trabalho de impressão no Windows: percebe falta de etiqueta/ribbon, tampa aberta, offline, cancelamento,
e diz quantas páginas saíram (para reimprimir só o que faltou).

Observação importante: o Windows mantém o trabalho na fila quando a impressora para; ao resolver o problema ela continua de
onde parou. O que o programa faz aqui é AVISAR e, se o trabalho for cancelado, oferecer a reimpressão do que faltou.
Só funciona se o driver da impressora informar o estado (comunicação bidirecional)."""
import time

# status do TRABALHO (JOB_STATUS_*)
J_PAUSADO, J_ERRO, J_DELETANDO, J_OFFLINE, J_SEM_PAPEL = 0x1, 0x2, 0x4, 0x20, 0x40
J_DELETADO, J_BLOQUEADO, J_INTERVENCAO = 0x100, 0x200, 0x400
# status da IMPRESSORA (PRINTER_STATUS_*)
P_PAUSADA, P_ERRO, P_ATOLADO, P_SEM_PAPEL, P_PROBLEMA_PAPEL, P_OFFLINE = 0x1, 0x2, 0x8, 0x10, 0x40, 0x80
P_NAO_DISPONIVEL, P_INTERVENCAO, P_SEM_TONER, P_PORTA_ABERTA = 0x1000, 0x100000, 0x40000, 0x400000


def problema(job_status, printer_status):
    """Texto amigável do que está impedindo a impressão ('' se está tudo bem)."""
    m = []
    if job_status & J_SEM_PAPEL or printer_status & (P_SEM_PAPEL | P_PROBLEMA_PAPEL):
        m.append("sem etiqueta ou ribbon")
    if printer_status & P_ATOLADO:
        m.append("etiqueta atolada")
    if printer_status & P_PORTA_ABERTA:
        m.append("tampa aberta")
    if job_status & J_OFFLINE or printer_status & (P_OFFLINE | P_NAO_DISPONIVEL):
        m.append("impressora offline ou desligada")
    if job_status & J_PAUSADO or printer_status & P_PAUSADA:
        m.append("impressão pausada")
    if not m and (job_status & (J_ERRO | J_INTERVENCAO | J_BLOQUEADO) or printer_status & (P_ERRO | P_INTERVENCAO | P_SEM_TONER)):
        m.append("erro na impressora (confira etiqueta, ribbon e tampa)")
    return " · ".join(m)


API_PADRAO = None     # (testes) uma API de mentira no lugar do Windows
INTERVALO, OBSERVA_DEPOIS = 1.0, 15   # segundos entre olhadas / quanto observar a impressora depois que o trabalho saiu da fila


class ApiWindows:
    """Acesso à fila de impressão do Windows (win32print)."""

    def abre(self, nome):
        import win32print
        return win32print.OpenPrinter(nome)

    def fecha(self, h):
        import win32print
        win32print.ClosePrinter(h)

    def trabalho(self, h, job_id):
        """dict(Status, PagesPrinted, TotalPages) ou None se o trabalho saiu da fila."""
        import win32print
        try:
            j = win32print.GetJob(h, job_id, 1)
        except Exception:
            return None
        return {"Status": j.get("Status", 0), "PagesPrinted": j.get("PagesPrinted", 0), "TotalPages": j.get("TotalPages", 0)}

    def status_impressora(self, h):
        import win32print
        try:
            return win32print.GetPrinter(h, 2).get("Status", 0)
        except Exception:
            return 0


def acompanha(nome, job_id, total_paginas, evento, api=None, intervalo=None, limite=900, observa_depois=None,
              relogio=time.time, dorme=time.sleep):
    """Acompanha o trabalho e chama evento(dict). Tipos: progresso, problema, retomou, ok, interrompida, sem_acompanhamento."""
    api = api or API_PADRAO or ApiWindows()
    intervalo = INTERVALO if intervalo is None else intervalo
    observa_depois = OBSERVA_DEPOIS if observa_depois is None else observa_depois
    try:
        h = api.abre(nome)
    except Exception:
        evento({"tipo": "sem_acompanhamento"})
        evento({"tipo": "fim"})
        return
    try:
        inicio, visto = relogio(), False
        ult = {"impressas": 0, "total": total_paginas, "deletando": False, "problema": ""}
        problema_atual, anterior = "", -1
        while True:
            if relogio() - inicio > limite:
                evento({"tipo": "tempo_esgotado", "impressas": ult["impressas"], "total": ult["total"]})
                return
            j = api.trabalho(h, job_id)
            if j is None:
                break
            visto = True
            ult["impressas"] = j["PagesPrinted"]
            ult["total"] = j["TotalPages"] or total_paginas
            if j["Status"] & (J_DELETANDO | J_DELETADO):
                ult["deletando"] = True
            prob = problema(j["Status"], api.status_impressora(h))
            ult["problema"] = prob
            if prob != problema_atual:
                evento({"tipo": "problema" if prob else "retomou", "texto": prob, "impressas": ult["impressas"],
                        "total": ult["total"]})
                problema_atual = prob
            elif not prob and ult["impressas"] != anterior:
                evento({"tipo": "progresso", "impressas": ult["impressas"], "total": ult["total"]})
            anterior = ult["impressas"]
            dorme(intervalo)
        if not visto:                    # já tinha saído da fila antes da primeira olhada: normalmente é sinal de que imprimiu
            evento({"tipo": "sem_acompanhamento", "motivo": "saiu_rapido"})
        elif ult["deletando"] or (ult["problema"] and 0 < ult["total"] and ult["impressas"] < ult["total"]):
            evento({"tipo": "interrompida", "impressas": ult["impressas"], "total": ult["total"]})
            return
        else:
            evento({"tipo": "ok", "impressas": ult["total"] or ult["impressas"], "total": ult["total"]})
        # depois de sair da fila: a impressora ainda pode parar (ex.: Zebra sem etiqueta/ribbon no modo ZPL)
        fim = relogio() + observa_depois
        while relogio() < fim and relogio() - inicio <= limite:
            prob = problema(0, api.status_impressora(h))
            if prob != problema_atual:
                evento({"tipo": "problema" if prob else "retomou", "texto": prob, "impressas": ult["impressas"],
                        "total": ult["total"], "depois": True})
                problema_atual = prob
            if prob:
                fim = max(fim, relogio() + observa_depois)
            dorme(intervalo)
    finally:
        try:
            api.fecha(h)
        except Exception:
            pass
        evento({"tipo": "fim"})
