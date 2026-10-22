# LaTeX IHR — suplemento do Word

Painel lateral para o Word desktop (Windows) e o Word Online. Você escreve em LaTeX e o conteúdo entra no documento aberto já com os estilos do template de relatório do IHR (*Template Relatório v4*; documentos da v3 também funcionam). Como as edições passam pelo próprio Word, o documento continua sendo um `.docx` comum, editável em coautoria por quem não usa o suplemento.

## O que faz

- **Leitura automática**: ao selecionar um trecho no Word, ele aparece em LaTeX no painel e fica *vinculado*.
- **Ctrl+Enter**: insere no cursor (modo inserção) ou atualiza o trecho vinculado. Depois de inserir ou atualizar, o trecho continua vinculado e selecionado, então editar de novo não duplica o conteúdo.
- **Proteções**: a leitura automática não sobrescreve edições ainda não aplicadas no painel, e o painel avisa se o trecho foi alterado no Word (por você ou por um colega) depois da leitura.
- Clicar num ponto fora do trecho, sem edições pendentes, desvincula e limpa o editor para um conteúdo novo.

| LaTeX | Word |
|---|---|
| `\chapter` … `\subparagraph` (com `*` = sem número) | Título 1 … Título 6 |
| `\textbf`, `\textit`, `\emph`, `\underline`, `\texttt`, `\textsuperscript`, `\textsubscript` | formatação de caractere |
| `$…$`, `\(…\)`, `\[…\]`, `$$…$$`, `equation`, `align`, `gather` | equações nativas do Word |
| `figure` / `table` / `quadro` + `\caption`, `\label`, `\fonte` | Legenda (campo SEQ), estilos Imagem, Fonte, Tabela padrão IHR e Quadro com grade |
| `tabular` (com `\multicolumn`), `itemize`, `enumerate`, `description` | tabelas e listas |
| `\ref`, `\eqref`, `\pageref` | referências cruzadas do Word (indicadores ocultos `_RefS_`, `_RefN_` e `_RefE_`) |
| `\footnote`, `\url`, `\href`, `\newpage`, `\tableofcontents` | nota de rodapé, hiperlink, quebra de página, sumário |
| `\title`, `\subtitulo`, `\local`, `\date`, `\maketitle` | textos da capa (estilos *Capa - …* da v4) |
| `\cite`, `\textcite`/`\citeonline`, `\citeauthor`, `\citeyear` | citação ABNT autor-data, num campo `ADDIN IHRCITE` (como fazem Zotero e Mendeley) |
| `\printbibliography`, `\refmanual{…}` | capítulo Referências no estilo *Bibliografia*; entrada manual |

## Referências

- O `.bib` carregado no painel fica guardado **dentro do .docx**, numa parte XML personalizada (`urn:ihr-latex-word:bib`), e vai junto com o arquivo.
- Cada citação é um campo que guarda as chaves e mostra o texto "(SILVA, 2020)". Quem não usa o suplemento vê texto comum.
- A lista fica no capítulo "Referências" (Título 1 sem número), com parágrafos no estilo nativo *Bibliografia*:
  - entradas geradas do `.bib` levam um indicador oculto `_RefB_<chave>`;
  - entradas digitadas à mão (por exemplo, por colegas sem o suplemento) não têm indicador. Elas nunca são apagadas, só reordenadas em ordem alfabética junto com as demais.
- **Atualizar lista no documento** (ou qualquer inserção com `\cite`) refaz a lista: entram todas as chaves citadas no documento, mais as entradas manuais.

As imagens de `\includegraphics{nome}` precisam ser carregadas no painel, na seção *Imagens*. As imagens lidas do documento entram nessa lista automaticamente.

## Instalação (Word desktop, Windows)

1. Salve o `manifest.xml` numa pasta do computador.
2. Registre o manifesto: execute o arquivo `.reg` que acompanha o pacote, ou crie, em
   `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer`,
   um valor de cadeia com nome `9bda6c45-f03e-4100-8913-ccba270fcbbc` e o caminho completo do `manifest.xml` como dado.
3. Feche e abra o Word. O botão **Editor LaTeX** aparece na guia *Página Inicial*, no grupo *LaTeX*.

Para desinstalar, apague esse valor do registro.

## Estrutura

- `taskpane.html`, `taskpane.js`, `taskpane.css`: o painel
- `lib/latex2ooxml.js`: conversão de LaTeX para OOXML
- `lib/ooxml2latex.js`: conversão de OOXML para LaTeX
- `lib/math.js`: conversão de equações (LaTeX ↔ OMML), usando o KaTeX para interpretar o LaTeX
- `lib/template-parts.js`: estilos e numeração extraídos do template IHR
- `vendor/katex.min.js`: KaTeX 0.16 (licença MIT)
- `test/`: scripts de teste em Node
