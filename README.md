# Manutenção SQL Script

Este projeto foi desenvolvido para automatizar a geração de scripts SQL para manutenção de registros em tabelas do banco de dados. Ele permite:
- Geração de arquivos SQL a partir de uma planilha de entrada.
- Personalização de tabelas, campos e valores a serem alterados.
- Salvamento dos scripts para execução manual.

## Funcionalidades
- Geração de scripts SQL com `BEGIN TRAN` e `ROLLBACK`.
- Suporte para múltiplos campos e valores de atualização.
- Organização de comandos para execução manual no banco de dados.

## Como Usar
1. Clone este repositório:
   ```bash
   git clone https://github.com/seu-usuario/manutencao-sql-script.git
   ```
2. Certifique-se de ter o Python instalado na sua máquina.
3. Instale as dependências necessárias:
   ```bash
   pip install pandas pyodbc
   ```
4. Execute o script principal:
   ```bash
   python GER_SCRIPT_SQL.py
   ```

## Estrutura do Projeto
- **GER_SCRIPT_SQL.py**: Script principal que realiza a geração de arquivos SQL.
- **Planilhas de entrada**: Exemplo de planilha com os campos `ID`, `Campo`, `Valor`.
- **Diretório de saída**: Os scripts SQL gerados serão salvos neste diretório.

## Exemplo de Entrada
Planilha de entrada no formato `.xlsx` ou `.csv`:
