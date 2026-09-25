# automacao_prev

Automação dos relatórios SAP, preparação das pastas e tratamento dos arquivos Excel.

## Pré-requisitos

- Acesso à unidade de rede `X:` com permissão de leitura e gravação.
- Acesso ao SAP.
- Python 3.11 ou mais recente, com o Python Launcher (`py`).
- Google Chrome instalado.

## Configuração em outro notebook

1. Copie a pasta `automacao_prev` para o notebook ou para um local de rede acessível.
2. Garanta que a unidade `X:` esteja conectada com a mesma estrutura de pastas.
3. Abra o PowerShell dentro da pasta `automacao_prev`.
4. Instale as dependências:

	```powershell
	py -m pip install -r requirements.txt
	```

5. Execute `executar_automacao.bat` com duplo clique.

O arquivo BAT procura o Python instalado localmente e não depende do nome do usuário do Windows.

## Como executar

1. Abra `executar_automacao.bat`.
2. Informe o usuário do SAP antes do sufixo `@icatuseguros.com.br`.
3. Informe a senha.
4. Informe as datas no formato `DD.MM.AAAA`.
5. Clique em `Iniciar automação`.

Datas futuras são bloqueadas antes do acesso ao SAP. Se o e-mail, a senha ou uma data estiverem inválidos, a aplicação informa o problema, mantém a janela aberta e permite corrigir e tentar novamente.

## Durante a execução

É permitido usar o computador normalmente: abrir outros programas, ler e-mails, trabalhar em documentos e usar o navegador para outras tarefas.

Não faça estas ações até aparecer a mensagem de conclusão:

- Não feche a janela da automação.
- Não feche o Chrome aberto pela automação.
- Não clique ou digite na tela do SAP usada pela automação.
- Não abra outra automação ao mesmo tempo.
- Não coloque o computador em suspensão, hibernação ou desligamento.
- Não mova, renomeie ou abra os arquivos que estão sendo gerados.

Para interromper uma execução, use o botão `Cancelar execução` e aguarde a confirmação.

## Arquivos e regras dos relatórios

Os arquivos são salvos em:

```text
X:\Comum-PagamentosDiarios
```

Os relatórios gerais `RELATORIO SEG 8H.xlsx`, `RELATORIO SEG 11H.xlsx`, `RELATORIO RG 8H.xlsx` e `RELATORIO RG 11H.xlsx` registram o histórico dos lotes tratados. Não apague esses arquivos, pois eles são usados para evitar reprocessamento.

Em `DEVOLUÇÃO IS`, entram documentos `65` e documentos `79` somente quando a coluna `Origem` contém `PGBL`. Documentos `79` com qualquer outra origem são descartados automaticamente antes da geração dos arquivos.

Não é necessário apagar manualmente arquivos ou pastas de `DEVOLUÇÃO IS`.

Um log sem acentos é salvo dentro da pasta `MANHÃ` da execução.

## Gerar executável

Em uma máquina com Python instalado, execute `gerar_executavel.bat`. O arquivo será criado em `dist\AutomacaoRelatorios.exe`.
