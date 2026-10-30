# PowerPoint library tab and cleaner import menu
Objetivo: guardar os PowerPoints usados numa aba própria da gaveta, prontos para abrir ou levar ao projeto, e deixar o menu Importar só com as opções usadas.

## Critérios de aceite
- **PPTLIB-01**: QUANDO o usuário abre a gaveta ENTÃO existe a aba "PowerPoints", que pode ser desativada como as outras abas (exceto Shows e Mídia)
- **PPTLIB-02**: QUANDO a aba PowerPoints é aberta pela primeira vez ENTÃO a pasta "PowerPoint" da pasta de dados do FreeShow (Documentos\FreeShow\PowerPoint) aparece como pasta padrão, sem opção de excluir
- **PPTLIB-03**: QUANDO o usuário clica em "Adicionar PowerPoint" e escolhe arquivos .ppt/.pptx/.key ENTÃO os arquivos são copiados para a pasta padrão e aparecem na lista; SE já existe um arquivo com o mesmo nome ENTÃO a cópia recebe o sufixo " (2)", " (3)"…
- **PPTLIB-04**: QUANDO o usuário adiciona uma pasta do Windows ENTÃO a aba lista os .ppt/.pptx/.key dessa pasta, sem copiar nada
- **PPTLIB-05**: QUANDO o usuário clica num PowerPoint da lista ENTÃO ele abre na área principal com o botão Iniciar (como um item de projeto)
- **PPTLIB-06**: QUANDO o usuário arrasta um PowerPoint da lista para o projeto, ou dá duplo clique, ENTÃO ele entra no projeto como item PowerPoint (com animações), não convertido em slides
- **PPTLIB-07**: QUANDO o usuário abre Importar ENTÃO aparecem só: show e projeto do FreeShow (e "Mais opções" do FreeShow), área de transferência, PDF, PowerPoint, Arquivo de texto, CSV e Word
- **PPTLIB-08**: QUANDO o usuário importa um PowerPoint pelo menu Importar ENTÃO o arquivo é copiado para a pasta padrão de PowerPoints e entra no projeto atual como item PowerPoint, sem passar pela tela de opções

## Casos de borda
- Pasta adicionada que não existe mais: a lista fica vazia, sem erro.
- Cancelar a escolha de arquivos não muda nada.
- Arquivos que não são .ppt/.pptx/.key nas pastas não aparecem.

## Fora de escopo
- Miniatura dos slides na lista (aparece ao iniciar, ou pelo pré-carregamento).
- Converter PowerPoint em slides de texto ou em PDF (saem do Importar; o código fica, sem atalho na interface).
- Formatos de outros programas no Importar (ProPresenter, EasyWorship, OpenLP, ChordPro, Lessons.church etc.): saem do menu; o código de importação continua no projeto, para voltar se precisar.

## Decisões críticas
- **Posição da aba:** no fim da barra, depois de Funções — não muda os atalhos Ctrl+1…9 das abas existentes.
- **Pasta padrão:** Documentos\FreeShow\PowerPoint (pasta de dados do FreeShow), criada pelo Electron sob demanda.
