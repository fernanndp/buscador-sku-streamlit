# Buscador de SKU

App simples em Streamlit para comparar uma planilha de referência de SKUs com outra planilha de busca.

## O que o app faz

- verifica se os EANs da planilha 1 existem na planilha 2
- compara descrições com tolerância a ordem diferente, abreviações e equivalências
- usa exclusivamente o dicionário enviado pelo usuário para normalização e similaridade
- gera um Excel final com resumo e resultados para revisão

## Formatos aceitos

### Planilha 1
Exemplo de colunas:
- `Produtos`
- `EAN 13 - UND`

### Planilha 2
Exemplo de colunas:
- `Descricao`
- `Codigo Barras`

### Dicionário
O dicionário é obrigatório.

Você pode enviar:

#### Opção 1 — CSV ou Excel simples
Com duas colunas:
- `Padrao`
- `Substituto`

Nesse formato, o app usará apenas regras de substituição.

#### Opção 2 — Excel com abas de configuração
Abas recomendadas:

- `Regras` → colunas `Padrao` e `Substituto`
- `Stopwords` → coluna `Palavra`
- `CategoryNoise` → coluna `Palavra`
- `BrandHints` → coluna `Palavra`

## Como rodar

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Observações

- o app tenta detectar automaticamente as colunas principais
- não existe fallback interno de dicionário no código
- somente os termos e regras enviados no arquivo de dicionário serão usados
- se uma aba não existir no Excel do dicionário, aquela configuração ficará vazia
