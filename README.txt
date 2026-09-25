AUTOMACAO_PREV

Automacao dos relatorios SAP, preparacao das pastas e tratamento dos arquivos Excel.

PRE-REQUISITOS

- Acesso a unidade de rede X: com permissao de leitura e gravacao.
- Acesso ao SAP.
- Python 3.11 ou mais recente, com o Python Launcher (py).
- Google Chrome instalado.

CONFIGURACAO EM OUTRO NOTEBOOK

1. Copie a pasta automacao_prev para o notebook ou para um local de rede acessivel.
2. Garanta que a unidade X: esteja conectada com a mesma estrutura de pastas.
3. Abra o PowerShell dentro da pasta automacao_prev.
4. Instale as dependencias com o comando:

   py -m pip install -r requirements.txt

5. Execute executar_automacao.bat com duplo clique.

O arquivo BAT procura o Python instalado localmente e nao depende do nome do usuario do Windows.

COMO EXECUTAR

1. Abra executar_automacao.bat.
2. Informe o usuario do SAP antes do sufixo @icatuseguros.com.br.
3. Informe a senha.
4. Informe as datas no formato DD.MM.AAAA.
5. Clique em Iniciar automacao.

Datas futuras sao bloqueadas antes do acesso ao SAP. Se o e-mail, a senha ou uma data estiverem invalidos, a aplicacao informa o problema, mantem a janela aberta e permite corrigir e tentar novamente.

DURANTE A EXECUCAO

E permitido usar o computador normalmente: abrir outros programas, ler e-mails, trabalhar em documentos e usar o navegador para outras tarefas.

Nao faca estas acoes ate aparecer a mensagem de conclusao:

- Nao feche a janela da automacao.
- Nao feche o Chrome aberto pela automacao.
- Nao clique ou digite na tela do SAP usada pela automacao.
- Nao abra outra automacao ao mesmo tempo.
- Nao coloque o computador em suspensao, hibernacao ou desligamento.
- Nao mova, renomeie ou abra os arquivos que estao sendo gerados.

Para interromper uma execucao, use o botao Cancelar execucao e aguarde a confirmacao.

ARQUIVOS E REGRAS DOS RELATORIOS

Os arquivos sao salvos em:

X:\Comum-PagamentosDiarios

Os relatorios gerais RELATORIO SEG 8H.xlsx, RELATORIO SEG 11H.xlsx, RELATORIO RG 8H.xlsx e RELATORIO RG 11H.xlsx registram o historico dos lotes tratados. Nao apague esses arquivos, pois eles sao usados para evitar reprocessamento.

Em DEVOLUCAO IS, entram documentos 65 e documentos 79 somente quando a coluna Origem contem PGBL. Documentos 79 com qualquer outra origem sao descartados automaticamente antes da geracao dos arquivos.

Nao e necessario apagar manualmente arquivos ou pastas de DEVOLUCAO IS.

Um log sem acentos e salvo dentro da pasta MANHA da execucao.

GERAR EXECUTAVEL

Em uma maquina com Python instalado, execute gerar_executavel.bat. O arquivo sera criado em dist\AutomacaoRelatorios.exe.
