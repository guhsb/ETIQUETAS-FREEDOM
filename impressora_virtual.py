# -*- coding: utf-8 -*-
"""Impressora virtual 'Etiqueta Freedom': o ERP imprime o DANFE nela e o programa lê o PDF gerado.

Como funciona: cria uma impressora do Windows (driver "Microsoft Print To PDF") ligada a uma porta que é um
ARQUIVO (dados\\entrada\\danfe.pdf). Ao imprimir, o Windows grava o PDF ali. Uma tarefa agendada abre o programa
quando algum documento é impresso (evento 307 do log de impressão), e o programa lê o DANFE.
Precisa de permissão de administrador uma única vez (instalação).
"""
import os
import subprocess
import sys
import tempfile

from etiqueta_core import pasta_dados

NOME = "Etiqueta Freedom"
TAREFA = "EtiquetasFreedom-DANFE"
SEM_JANELA = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def pasta_entrada():
    p = os.path.join(pasta_dados(), "entrada")
    os.makedirs(p, exist_ok=True)
    return p


def porta_arquivo():
    return os.path.join(pasta_entrada(), "danfe.pdf")


def log_instalacao():
    return os.path.join(pasta_dados(), "instalacao_impressora.log")


def eh_exe():
    return bool(getattr(sys, "frozen", False))


def impressora_instalada():
    try:
        import win32print
        flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        return any(p[2].lower() == NOME.lower() for p in win32print.EnumPrinters(flags))
    except Exception:
        return False


def tarefa_instalada():
    try:
        r = subprocess.run(["schtasks", "/Query", "/TN", TAREFA], capture_output=True, creationflags=SEM_JANELA)
        return r.returncode == 0
    except Exception:
        return False


def _ps(t):
    """Texto seguro dentro de aspas simples do PowerShell."""
    return str(t).replace("'", "''")


def script_instalar(exe, usuario):
    return f"""$ErrorActionPreference = 'Stop'
$log = '{_ps(log_instalacao())}'
function Escreve($m) {{ Add-Content -LiteralPath $log -Value ((Get-Date -Format 's') + ' ' + $m) -Encoding UTF8 }}
Set-Content -LiteralPath $log -Value '' -Encoding UTF8
try {{
  $nome = '{_ps(NOME)}'
  $entrada = '{_ps(pasta_entrada())}'
  $porta = '{_ps(porta_arquivo())}'
  New-Item -ItemType Directory -Force -Path $entrada | Out-Null
  if (-not (Get-PrinterPort -Name $porta -ErrorAction SilentlyContinue)) {{ Add-PrinterPort -Name $porta; Escreve 'porta criada' }}
  if (-not (Get-Printer -Name $nome -ErrorAction SilentlyContinue)) {{
    Add-Printer -Name $nome -DriverName 'Microsoft Print To PDF' -PortName $porta
    Escreve 'impressora criada'
  }} else {{ Escreve 'impressora ja existia' }}
  & wevtutil.exe sl Microsoft-Windows-PrintService/Operational /e:true
  Escreve 'log de impressao ativado'
  $acao = New-ScheduledTaskAction -Execute '{_ps(exe)}' -Argument '--danfe'
  $cls = Get-CimClass -Namespace 'ROOT\\Microsoft\\Windows\\TaskScheduler' -ClassName 'MSFT_TaskEventTrigger'
  $gatilho = New-CimInstance -CimClass $cls -ClientOnly
  $gatilho.Enabled = $true
  $gatilho.Subscription = '<QueryList><Query Id="0" Path="Microsoft-Windows-PrintService/Operational"><Select Path="Microsoft-Windows-PrintService/Operational">*[System[(EventID=307)]]</Select></Query></QueryList>'
  $principal = New-ScheduledTaskPrincipal -UserId '{_ps(usuario)}' -LogonType Interactive -RunLevel Limited
  $config = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances Parallel -ExecutionTimeLimit (New-TimeSpan -Hours 12)
  Register-ScheduledTask -TaskName '{TAREFA}' -Action $acao -Trigger $gatilho -Principal $principal -Settings $config -Force | Out-Null
  Escreve 'tarefa criada'
  Escreve 'OK'
}} catch {{
  Escreve ('ERRO: ' + $_.Exception.Message)
}}
"""


def script_remover():
    return f"""$ErrorActionPreference = 'SilentlyContinue'
$log = '{_ps(log_instalacao())}'
function Escreve($m) {{ Add-Content -LiteralPath $log -Value ((Get-Date -Format 's') + ' ' + $m) -Encoding UTF8 }}
Set-Content -LiteralPath $log -Value '' -Encoding UTF8
Unregister-ScheduledTask -TaskName '{TAREFA}' -Confirm:$false
Remove-Printer -Name '{_ps(NOME)}'
Remove-PrinterPort -Name '{_ps(porta_arquivo())}'
Escreve 'removido'
Escreve 'OK'
"""


def executa_como_admin(script, espera=True):
    """Roda o script PowerShell pedindo permissão de administrador (UAC). Devolve o texto do log."""
    ps1 = os.path.join(tempfile.gettempdir(), "etiquetas_freedom_admin.ps1")
    with open(ps1, "w", encoding="utf-8-sig") as f:
        f.write(script)
    cmd = ("Start-Process powershell -Verb RunAs -WindowStyle Hidden " + ("-Wait " if espera else "") +
           "-ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File','\"" + ps1.replace("'", "''") + "\"'")
    subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", cmd],
                   creationflags=SEM_JANELA, capture_output=True)
    try:
        with open(log_instalacao(), encoding="utf-8-sig") as f:
            return f.read().strip()
    except OSError:
        return ""


def instalar():
    if not eh_exe():
        raise RuntimeError("A instalação só funciona no programa .exe (não no código-fonte).")
    usuario = (os.environ.get("USERDOMAIN", "") + "\\" + os.environ.get("USERNAME", "")).strip("\\")
    return executa_como_admin(script_instalar(sys.executable, usuario))


def remover():
    return executa_como_admin(script_remover())


# ------------------------------------------------------------------ abrir junto com o Windows (sem administrador)
CHAVE_RUN = r"Software\Microsoft\Windows\CurrentVersion\Run"
NOME_RUN = "EtiquetasFreedom"


def inicio_automatico_ativo():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, CHAVE_RUN) as k:
            winreg.QueryValueEx(k, NOME_RUN)
        return True
    except Exception:
        return False


def define_inicio_automatico(ligar):
    """Liga/desliga a abertura do programa (minimizado) quando o Windows inicia. Só vale para o .exe."""
    import winreg
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, CHAVE_RUN, 0, winreg.KEY_SET_VALUE) as k:
        if ligar:
            if not eh_exe():
                raise RuntimeError("Isso só funciona no programa .exe (não no código-fonte).")
            winreg.SetValueEx(k, NOME_RUN, 0, winreg.REG_SZ, f'"{sys.executable}" --minimizado')
        else:
            try:
                winreg.DeleteValue(k, NOME_RUN)
            except FileNotFoundError:
                pass
