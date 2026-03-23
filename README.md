# Validador MIMC

Aplicativo desktop em Python para validar 1 planilha `.xlsx` por vez.

## O que faz nesta versão

Analisa as abas:
- `TALHAO`
- `INVENTARIO`
- `PRODUCAO`
- `VENDAS`
- `DESPESAS`

Regras implementadas:
- `INV-001` valor menor que 100 ou maior que 500000
- `INV-002` data de fabricação maior que data de aquisição
- `PRO-001` rateio sem todos os talhões em produção
- `PRO-002` valores diferentes em produção com rateio
- `VEN-001` preço de venda acima de 100
- `DES-001` falha de recorrência administrativa
- `DES-002` atividade com insumo sem mão de obra associada
- `DES-004` rateio sem todos os talhões cadastrados
- `DES-005` valores diferentes em despesa com rateio

## Como gerar o EXE sem instalar nada no computador

### 1. Criar repositório no GitHub
Crie um repositório novo com qualquer nome, por exemplo:

`validador-mimc`

### 2. Enviar estes arquivos
No repositório, clique em **Add file** > **Upload files** e envie:
- `app.py`
- `requirements.txt`
- a pasta `.github/workflows/` com o arquivo `build.yml`

### 3. Rodar a compilação
Depois que os arquivos forem enviados:
- abra a aba **Actions**
- clique no workflow **Build Windows EXE**
- clique em **Run workflow**

### 4. Baixar o executável
Quando terminar:
- abra a execução concluída
- em **Artifacts**, baixe `validador-mimc-exe`

Dentro do arquivo baixado estará o `validador-mimc.exe`.

## Como usar o programa

1. Abra o `validador-mimc.exe`
2. Clique em **Selecionar planilha**
3. Escolha uma planilha `.xlsx`
4. Clique em **Analisar**
5. Revise as inconsistências
6. Clique em **Exportar relatório** se quiser salvar em Excel

## Observações

- Esta V1 foi desenhada para processar 1 planilha por vez.
- O reconhecimento de colunas tenta localizar nomes próximos, mas quanto mais padronizada a planilha, melhor o resultado.
- A lógica pode ser refinada nas próximas versões com base nas suas bases reais.
