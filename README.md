# JuntaPDF

Programa de emenda local de arquivos PDF e documentos Word (.doc e .docx) em lote.

![Ícone - JuntaPDF](assets/icons/juntapdf.ico)

## 1. Requisitos

Para uso adequado do programa, o usuário deve possuir:
- **Sistema Operacional:** Windows 10 ou 11

Para uso da funcionalidade de conversão DOC/DOCX->PDF, o usuário deve ter o *Microsoft Word* instalado na versão mais recente.

## 2. Uso

### Baixando e Instalando o Programa

Para usar o `JuntaPDF`, primeiro, você deve baixar o arquivo .exe disponível [aqui](link). Procure pela versão mais recente (*Latest*) e clique no arquivo .exe para fazer o download.

> [!Warning](Aviso)
> Caso você ainda tenha o executável de uma versão antiga do programa, recomenda-se excluí-lo.

> **🚨 NOME DO AVISO CUSTOMIZADO**
> Conteúdo do aviso aqui.


Baixado o programa, você pode colocar o arquivo .exe onde achar melhor.

### Abrindo o Programa

Feito isso, clique no arquivo .exe para abrir o programa.

![Tela inicial](assets/images/01_starting_screen.png)

Junte os arquivos que você deseja emendar em uma pasta. O programa utiliza uma convenção de nomenclatura de arquivos para saber quais arquivos emendar e em que ordem fazê-lo.

### Organizando os Arquivos de Entrada

Todos os arquivos colocados na pasta devem ser nomeados como:

`nome_numero.pdf`

Sendo que:
- `nome` pode ser qualquer texto (incluindo letras maiúsculas, minúsculas, acentos, sinais e números)
- `_` entre `nome` e `numero` é **INDISPENSÁVEL**
- `numero` deve ser expresso em números inteiros, na ordem desejada para emenda dos arquivos do conjunto `nome`
- O conjunto de arquivos de entrada com o mesmo elemento `nome` será mesclado em um único arquivo `nome.pdf`, seguindo a ordem dos índices `numero`

Para que o programa seja capaz de processá-los e emendá-los adequadamente.

Observe o exemplo:

![Pasta "test"](assets/images/02_test_files.png)

Aqui, tenho 5 diferentes grupos de arquivos (disponíveis na pasta [test/test_0](test/test_0)):
- Grupo "batch": `batch_1.pdf`, `batch_2.pdf`
- Grupo "copy_test": `copy_test_1.pdf`, `copy_test_2.pdf`, `copy_test_3.doc`, `copy_test_4.pdf`, `copy_test_5.pdf`
- Grupo "exodia": `exodia_0.pdf`, `exodia_1.pdf`, `exodia_2.pdf`, `exodia_3.pdf`, `exodia_4.pdf`
- Grupo "lorem ipsum": `lorem ipsum_1.pdf`, `lorem ipsum_2.pdf`, `lorem ipsum_3.pdf`, `lorem ipsum_4.pdf`, `lorem ipsum_5.pdf`, `lorem ipsum_6.pdf`, `lorem ipsum_7.pdf`, `lorem ipsum_8.pdf`, `lorem ipsum_9.pdf`, `lorem ipsum_10.pdf`, `lorem ipsum_11.docx`
- Grupo "something 04-05-26": `something 04-05-26_1.pdf`, `something 04-05-26_2.pdf`, `something 04-05-26_3.pdf`, `something 04-05-26_4.pdf`, `something 04-05-26_5.pdf`, `something 04-05-26_6.pdf`

O programa reconhece o grupo do arquivo e a ordem em que ele deve ser emendado ao arquivo de saída através da nomenclatura do documento. Tomemos o arquivo `lorem ipsum_5.pdf`, por exemplo:
- Nome do arquivo (sem a extensão): `lorem ipsum_5`
- Texto **antes** do underscore ("_"): `lorem ipsum`
- Texto **depois** do underscore ("_"): `5`
Logo, este arquivo pertence ao grupo "lorem ipsum", e é o arquivo de índice 5.

Para o conjunto de arquivos apresentados acima, o programa processará os seguintes arquivos de saída:
- `batch.pdf`
- `copy_test.pdf`
- `exodia.pdf`
- `lorem ipsum.pdf`
- `something 04-05-26.pdf`

### Selecionando a Pasta

Na janela do programa, clique no botão `Selecionar Pasta` para escolher a pasta onde seus arquivos de entrada estão salvos.

![Escolhendo pasta de entrada](assets/images/03_input_folder.png)

### Emendando os Arquivos

Para emendar os PDFs, clique no botão `Juntar PDFs`.

![Emendando PDFs](assets/images/04_compiling.png)

Após o processamento dos arquivos, o programa abrirá uma janela para que você escolha a pasta onde os arquivos de saída serão salvos.

![Escolhendo pasta de saída](assets/images/05_output_folder.png)

![PDFs emendados](assets/images/06_success.png)

## 3. Releases

### JuntaPDF v0.1.0 (*alpha release*)

> ![Warning]
> O lançamento alfa (*alpha release*) foi desenvolvido para **testes internos**, visando identificar e corrigir bugs antes do lançamento de uma versão estável.

Data de lançamento: `13/05/2026`

Para fazer o download dessa versão, clique [aqui](link).

*Release* inicial do programa de emenda local de arquivos PDF e documentos Word (.doc e .docx) em lote.

**Features:**

- Recebe arquivos .pdf, juntando-os em PDFs "costurados" de acordo com a nomenclatura e numeração dos arquivos. Por exemplo:
  - Arquivos de entrada: `string1_1.pdf`, `string1_2.pdf`, `...`, `string1_10.pdf`, `string2_1.pdf`, `string2_2.pdf`, `...`, `string2_10.pdf`, `string3_1.pdf`, `string3_2.pdf`, `...`, `string3_10.pdf`;
  - Arquivos de saída: `string1.pdf`, `string2.pdf`, `string3.pdf`
  - Onde "string1", "string2" e "string3" podem ser quaisquer strings de texto (incluindo letras maiúsculas, minúsculas, acentos, sinais e números).
- Compatível com arquivos .doc e .docx, convertendo-os em .pdf antes da mescla.
- Permite que o usuário escolha o diretório de salvamento para o(s) arquivo(s) de saída.

Clique [aqui](link) para acessar o **changelog completo**.

## 4. Desenvolvimento

#### Autor:

Timóteo Altoé (*handle:* [imbaTIMvel](https://github.com/imbaTIMvel))

#### Datas:

`04/05/2026` Início do projeto

`05/05/2026` Lançamento da versão para testes internos

`13/05/2026` 

`06/05/2026` Lançamento da primeira versão oficial
