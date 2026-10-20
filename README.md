# Etiquetas Freedom

Programa para Windows que imprime as etiquetas de volume (e o comunicado de avaria) em qualquer impressora ativa,
lê o DANFE direto do ERP e guarda tudo na pasta do `.exe`.

## Gerar o .exe no GitHub (sem instalar nada)

1. Crie um repositório **privado** no GitHub e envie todos estes arquivos (pasta inteira, incluindo `.github`).
2. Abra a aba **Actions**. O fluxo **Gerar EtiquetasFreedom.exe** roda sozinho a cada envio para `main`
   (ou clique em **Run workflow**). Leva uns 3 a 5 minutos.
3. Quando terminar, abra a execução e baixe **EtiquetasFreedom** em **Artifacts**: é um `.zip` com o `.exe` pronto.
4. Para publicar uma versão fixa: crie uma tag `v1.0.0` (Releases > Draft a new release > tag `v1.0.0`).
   O `.zip` aparece em **Releases** e o número da versão entra no programa.

Quem usa só precisa baixar o `.zip`, extrair numa pasta (ex.: `C:\Etiquetas`) e abrir o `EtiquetasFreedom.exe`.

## Dados de clientes

O `clientes.csv` (CNPJ, telefones, endereços) **não é enviado** ao GitHub (está no `.gitignore`).
Coloque o arquivo exportado do sistema ao lado do `.exe`: o programa lê e, a cada arquivo novo, adiciona só os clientes novos.
Em repositório **privado**, se quiser embutir uma base inicial no `.exe`: `git add -f clientes.csv`.

## Desativar o programa de longe

O arquivo `controle_etiquetas.json` deste repositório é lido pelo programa ao abrir e a cada 5 minutos:

```json
{ "bloqueado": false, "mensagem": "", "bloquear": [], "versaoMinima": "" }
```

- `"bloqueado": true` desativa para todos (mostra a `mensagem` e fecha).
- `"bloquear": ["PC-EXPEDICAO", "maria"]` bloqueia só esses computadores/usuários do Windows.
- `"versaoMinima": "1.2.0"` bloqueia versões mais antigas.

O endereço lido é `https://raw.githubusercontent.com/<usuario>/<repo>/main/controle_etiquetas.json`, que **só abre se o
repositório for público**. Com repositório privado, publique o JSON em outro lugar (ex.: um Gist) e informe o endereço em
**Settings > Secrets and variables > Actions > Variables > `CONTROLE_URL`**. Sem internet, vale a última resposta guardada.

Mais detalhes de uso no `LEIA-ME.txt`.
