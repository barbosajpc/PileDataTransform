# Pile Plant Data Transformation

Este projeto fornece uma aplicação web interativa usando Streamlit para transformação de dados de arquivos Excel e CSV relacionados a fundações de plantas fotovoltaicas. A aplicação permite a transformação e o formato dos dados, incluindo a mesclagem e alinhamento das células em um arquivo Excel resultante.

Acesse a aplicação no link fornecido via Streamlit: [Clique Aqui](piledatatransformation.streamlit.app)

## Funcionalidades

- **Transformação de Dados**: Formata e reordena os dados de entrada.
- **Mesclagem e Alinhamento de Células**: Mescla células baseadas em valores comuns e alinha vertical e horizontalmente o conteúdo.
- **Interface Interativa**: Usa Streamlit para uma interface web amigável para o upload de arquivos e visualização dos dados.

## Ferramentas Utilizadas

- **Python**: Linguagem de programação principal.
- **pandas**: Biblioteca para manipulação de dados.
- **openpyxl**: Biblioteca para leitura e escrita de arquivos Excel.
- **Streamlit**: Biblioteca para criação de aplicativos web interativos.
- **Git**: Para organização e versionamento das etapas da elaboração da aplicação.

## Instalação

Certifique-se de ter o Python instalado. Em seguida, instale as dependências necessárias usando `pip`:

```bash
pip install pandas openpyxl streamlit
```

## Uso

1. **Inicie o Servidor Streamlit**:
   Navegue até o diretório do projeto e execute o seguinte comando para iniciar o aplicativo:

   ```bash
   streamlit run pile_data.py
   ```

2. **Interaja com a Aplicação**:
   - Acesse a aplicação através do navegador na URL fornecida pelo Streamlit (geralmente `http://localhost:8501`).
   - Faça o upload de um arquivo Excel, CSV ou XLS.
   - Visualize os dados e clique no botão "Transform Pile Plant Data" para aplicar as transformações.
   - Baixe o arquivo Excel transformado.

## Funções Principais

### `transform_pile_plant(data)`

Transforma e formata o DataFrame de entrada, cria uma nova coluna 'ID_PILE', e salva o resultado em um arquivo Excel.

- **Parâmetros**:
  - `data` (pd.DataFrame): DataFrame contendo os dados a serem transformados.

- **Retorna**:
  - `out_path` (str): Caminho do arquivo Excel resultante.

### `merge_and_align_cells(out_path, transformed_data)`

Mescla células baseadas em valores comuns e alinha o conteúdo vertical e horizontalmente.

- **Parâmetros**:
  - `out_path` (str): Caminho do arquivo Excel a ser formatado.
  - `transformed_data` (pd.DataFrame): DataFrame contendo os dados transformados.

- **Retorna**:
  - Nenhum. Salva o arquivo Excel modificado no caminho especificado.


## Contato

Para mais informações ou dúvidas, entre em contato com [jpedrobarbosa.jpb@gmail.com].

Linkedin: [Clique Aqui](https://www.linkedin.com/in/jo%C3%A3o-pedro-barbosa-697678254/)
