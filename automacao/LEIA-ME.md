# Mapa de Transportadoras na TV: atualização automática

Funciona igual ao painel de Entregas CIF. A TV abre o mapa direto de uma pasta
do SharePoint/OneDrive. Um script copia o relatório de parceiros para o arquivo
de dados ao lado do mapa, e o mapa confere a cada 5 minutos se há versão nova.
Ninguém precisa carregar arquivo na TV nem mexer na planilha.

```
SAP ──► relatorios\Lista de Parceiros*.xlsx ──► atualizar_dados_mapa.ps1 ──► dados\mapa_dados.js ──► mapa_tv.html (TV)
        (do jeito que sai, sem filtro/PROCX)       (agendado)                   (pasta sincronizada)     (confere a cada 5 min)
```

As regras continuam automáticas, aplicadas pelo próprio mapa:

- coluna "Ativo" = Y;
- "Código do PN" começando com C;
- só as transportadoras autorizadas: F00141, F01357, F01785, F01844, F02800,
  F03089, F04527, F04628 e F04958. O nome vem do próprio código, sem PROCX.

## 1. Montar a pasta no SharePoint

Crie uma pasta em qualquer biblioteca do SharePoint (ex.: `Logística/Mapa Transportadoras`).
Ela é independente do painel de entregas e pode ficar em outro lugar. A estrutura é esta:

```
Mapa Transportadoras\
  mapa_tv.html              ← cópia do mapa-transportadoras.html (renomeada)
  relatorios\               ← onde o relatório "Lista de Parceiros de negócios" é salvo
  dados\                    ← criada pelo script (mapa_dados.js e o log)
  automacao\
    atualizar_dados_mapa.ps1
    agendar_atualizacao.ps1
    abrir_mapa_tv.bat
```

Sincronize a pasta pelo OneDrive (botão **Sincronizar** na biblioteca) **no
computador da TV** e **no computador que vai rodar o script** (pode ser o mesmo
das Entregas CIF).

## 2. Primeiro teste, à mão

1. Salve o relatório do SAP em `relatorios\` com o nome de sempre
   (ex.: `Lista_de_Parceiros_de_negócios.xlsx`). Pode substituir o anterior:
   o script sempre usa o mais recente.
2. Clique com o botão direito em `automacao\atualizar_dados_mapa.ps1` → **Executar com o PowerShell**.
3. Abra `mapa_tv.html`: o mapa já aparece com os dados, sem escolher arquivo.

Se o relatório tiver outro nome (sem "Parceiros"), rode com `-PadraoRelatorio "*meu_nome*.xlsx"`.

## 3. Agendar

No computador que roda o script, abra o PowerShell na pasta `automacao` e rode:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File agendar_atualizacao.ps1
```

Isso cria a tarefa **"Mapa Transportadoras - atualizar dados"**, que roda a cada
15 minutos, de segunda a sábado, das 6h às 20h. O script só regrava os dados
quando o relatório mudou. O histórico fica em `dados\atualizacao.log`.

## 4. Deixar a TV ligada no mapa

No computador da TV, dê dois cliques em `automacao\abrir_mapa_tv.bat`: o Edge
abre em tela cheia, já no Modo TV. Para abrir sozinho ao ligar o computador,
coloque um atalho do `.bat` na pasta `shell:startup` (Win+R → `shell:startup`).
Para sair da tela cheia: **Alt+F4**.

Se o Edge já estiver aberto, feche-o antes de rodar o `.bat`, senão ele abre uma
aba comum em vez da tela cheia.

O computador da TV precisa de internet: o mapa baixa o leitor de Excel, o 3D e
as fontes da web.

## Histórico de mudanças

O botão **Mudanças**, no alto do mapa, abre um painel separado com tudo o que
mudou de um relatório para o outro, por data:

- **Entraram**: cliente novo, reativado ou que passou para uma transportadora autorizada.
- **Saíram**: cliente removido, inativado ou que passou para uma transportadora fora da lista.
- **Trocaram de transportadora**: antes e depois.
- **Mudaram de cidade**: cidade/UF anterior e atual.
- **Atenção**: cliente ativo novo que não aparece no mapa (sem transportadora ou com
  transportadora fora da lista).

O número verde no botão mostra quantos clientes mudaram na última atualização.
O painel tem filtros por período, transportadora e busca, e baixa a lista em .csv.

O script guarda uma foto dos clientes em `dados\clientes_base.tsv` e as mudanças em
`dados\historico.js`. Não apague esses arquivos: sem eles o histórico recomeça do zero.
Se um relatório vier incompleto (menos da metade dos clientes do anterior), o
histórico não é alterado e o log mostra um AVISO.

## Fora do guia de transportadoras

O botão **Fora do guia** (número vermelho) abre a aba com os clientes do mapa
cadastrados com uma transportadora diferente da indicada no **Guia de
Transportadoras por Estado** (cadastro SAP). O estado considerado é o da cidade
do cliente. A aba mostra os totais, o resumo por estado e por transportadora
cadastrada, a lista com "desde quando" (pelo histórico) e baixa um .csv.

No **Histórico de mudanças**, cada cliente que entra, troca de transportadora ou
muda de cidade fora do guia ganha a marca **⚠ Fora do guia**, e a opção
**Só fora do guia** filtra só esses casos.

O guia está gravado no próprio mapa. Se ele mudar, é preciso atualizar o
`mapa_tv.html` (tabela `GUIA_TRANSP`).

## Conferência

- No Modo TV, ao lado do botão de tela cheia, aparece "Dados de …", com a data
  em que o relatório foi salvo. Se o relatório tiver mais de 8 dias, o texto fica
  amarelo com um aviso.
- No Modo Individual, o rodapé mostra o nome do arquivo e as regras aplicadas.
- Para incluir ou tirar uma transportadora, edite a lista de códigos na tela de
  importação do mapa (Modo Individual). A lista fica salva naquele computador.
