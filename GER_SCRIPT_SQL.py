def executar_comandos_no_banco(conn, sql_file):
    """
    Executa os comandos SQL em um banco de dados conectado (REMOVIDO NESTA VERSÃO).

    Args:
        conn (pyodbc.Connection): Objeto de conexão com o banco de dados.
        sql_file (str): Caminho para o arquivo SQL com os comandos.
    """
    print("Execução no banco foi desativada. O script apenas gera o arquivo SQL para execução manual.")


def gerar_script_sql(tabela, coligada, campo_coligada, filial, campo_filial, campo_id, directory1, directory2, nome_arquivo):
    """
    Gera um script SQL com base em uma planilha de entrada.

    Args:
        tabela (str): Nome da tabela que será alterada.
        coligada (int): Valor do código da coligada.
        campo_coligada (str): Nome do campo correspondente ao código da coligada.
        filial (int ou None): Valor do código da filial (None se não for necessário).
        campo_filial (str ou None): Nome do campo correspondente ao código da filial (None se não for necessário).
        campo_id (str): Nome do campo correspondente ao ID.
        directory1 (str): Diretório para salvar o arquivo SQL.
        directory2 (str): Diretório onde está o arquivo de entrada.
        nome_arquivo (str): Nome do arquivo de entrada com os dados.

    Returns:
        str: Caminho do arquivo SQL gerado.
    """
    import pandas as pd
    import os

    input_file = os.path.join(directory2, nome_arquivo)

    if not os.path.exists(input_file):
        raise FileNotFoundError(f"O arquivo {input_file} não foi encontrado.")
    
    if input_file.endswith(".xlsx"):
        df = pd.read_excel(input_file)
    elif input_file.endswith(".csv"):
        df = pd.read_csv(input_file)
    else:
        raise ValueError("O arquivo deve ser no formato .xlsx ou .csv")
    
    # Transformar os nomes das colunas para minúsculas
    df.columns = map(str.lower, df.columns)
    
    # Verificar se as colunas esperadas estão presentes
    if not {"id", "campo", "valor"}.issubset(df.columns):
        raise ValueError("A planilha deve conter as colunas: ID, Campo e Valor.")
    
    updates = ["BEGIN TRAN;"]
    
    for id_registro, group in df.groupby("id"):
        set_clause = []
        for _, row in group.iterrows():
            valor = str(row['valor'])  # Garante que o valor seja tratado como string
            if ',' in valor:  # Corrigir separadores decimais
                valor = valor.replace(',', '.')
            
            # Determinar se o valor é numérico ou texto
            if valor.replace('.', '', 1).isdigit():  # Valor numérico
                set_clause.append(f"{row['campo']} = {valor}")
            else:  # Valor textual
                set_clause.append(f"{row['campo']} = '{valor}'")
        
        where_clause = f"WHERE {campo_id} = {id_registro} AND {campo_coligada} = {coligada}"
        if filial and campo_filial:
            where_clause += f" AND {campo_filial} = {filial}"
        update = f"UPDATE {tabela} SET {', '.join(set_clause)} {where_clause};"
        updates.append(update)
    
    updates.append("ROLLBACK;  -- Verifique se os comandos estão corretos antes de COMMIT")
    updates.append("COMMIT;")
    
    output_file = os.path.join(directory1, f"updates_{nome_arquivo.split('.')[0]}.sql")
    
    with open(output_file, "w") as file:
        file.write("\n".join(updates))
    
    print(f"Script SQL gerado com sucesso: {output_file}")
    return output_file


def main():
    tabela = input("Informe a tabela que será alterada: ")
    coligada = input("Informe o código da Coligada: ")
    campo_coligada = input("Informe o nome do campo para o código da Coligada: ")
    filial = input("Informe o código da Filial (pressione Enter se não for necessário): ") or None
    campo_filial = input("Informe o nome do campo para o código da Filial (pressione Enter se não for necessário): ") or None
    campo_id = input("Informe o nome do campo para o ID: ")
    directory1 = input("Informe o caminho do diretório para salvar dos updates (Ex: C:/caminho/): ")
    directory2 = input("Informe o caminho do diretório onde está salvo o arquivo com os dados (Ex: C:/caminho/): ")
    nome_arquivo = input("Informe o nome do arquivo que precisa ser processado (com extensão .xlsx ou .csv): ")
    
    try:
        # Gerar o SQL
        sql_file = gerar_script_sql(tabela, coligada, campo_coligada, filial, campo_filial, campo_id, directory1, directory2, nome_arquivo)
        
        # Perguntar se deseja processar no banco
        decisao = input("Deseja salvar apenas o arquivo (s) ou salvar e processar no banco (p)? ").lower()
        if decisao == "p":
            print("Execução no banco desativada. Por favor, processe o arquivo gerado manualmente.")
        else:
            print(f"O script foi salvo em {sql_file}.")
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    main()
