import pandas as pd
import mysql.connector
from mysql.connector import Error
import os
import re
from datetime import datetime

# Caminho do arquivo Excel
arquivo_excel = r"H:\Meu Drive\Documentos_One\ACOMPANHAMENTO_OS_2025-DESKTOP-QJ34J63.xlsx"

# Verificar existência
if not os.path.exists(arquivo_excel):
    raise FileNotFoundError(f"Arquivo não encontrado: {arquivo_excel}")
print("📂 Arquivo encontrado!")

try:
    # Ler Excel
    df = pd.read_excel(arquivo_excel)
    
    # Limpar nomes de colunas
    df.columns = [
        re.sub(r'[^a-zA-Z0-9_]', '_', str(col).lower().strip()).replace('__', '_')
        for col in df.columns
    ]

    # Remover colunas "Unnamed"
    df = df.loc[:, ~df.columns.str.contains('unnamed')]
    
    # Garantir coluna 'os' como int
    if 'os' not in df.columns:
        raise ValueError("Coluna 'os' não encontrada no Excel.")
    df['os'] = pd.to_numeric(df['os'], errors='coerce').fillna(0).astype(int)

    # Gerar numero_unico: ano corrente (2 dígitos) + OS 4 dígitos + sufixo se duplicado
    ano_corrente = datetime.now().year % 100
    existentes = set()

    def gerar_numero_unico(valor):
        if pd.isna(valor) or valor == 0:
            return None
        base = f"{ano_corrente}{valor:04d}"
        if base not in existentes:
            existentes.add(base)
            return base
        sufixo = 1
        while f"{base}{sufixo}" in existentes:
            sufixo += 1
        novo_valor = f"{base}{sufixo}"
        existentes.add(novo_valor)
        return novo_valor

    df['numero_unico'] = df['os'].apply(gerar_numero_unico)

    # Reordenar: numero_unico logo após os
    colunas = df.columns.tolist()
    idx_os = colunas.index('os')
    colunas.insert(idx_os + 1, colunas.pop(colunas.index('numero_unico')))
    df = df[colunas]

    # Conectar ao MySQL
    conn = mysql.connector.connect(
        host="localhost",
        user="root",
        password="",
        database="oficina_db",
        charset="utf8mb4"
    )
    cursor = conn.cursor()

    # Criar tabela com tipos apropriados
    col_defs = []
    for col in df.columns:
        if col == 'os':
            col_defs.append(f"`{col}` INT")
        elif col == 'numero_unico':
            col_defs.append(f"`{col}` VARCHAR(10) UNIQUE COMMENT 'Número único com prefixo do ano'")
        elif 'data' in col:
            col_defs.append(f"`{col}` DATE")
        elif 'valor' in col or 'total' in col:
            col_defs.append(f"`{col}` DECIMAL(15,2)")
        else:
            col_defs.append(f"`{col}` VARCHAR(255)")

    create_table_sql = f"""
    CREATE TABLE IF NOT EXISTS acompanhamento_os_2025 (
        id INT AUTO_INCREMENT PRIMARY KEY,
        {', '.join(col_defs)}
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    """
    cursor.execute(create_table_sql)
    print("✅ Tabela criada/verificada!")

    # Inserção ou atualização automática
    cont_inseridos, cont_atualizados, cont_ignorados = 0, 0, 0
    colunas_sql = ', '.join([f"`{col}`" for col in df.columns])
    placeholders = ', '.join(['%s'] * len(df.columns))
    update_sql = ', '.join([f"`{col}`=VALUES(`{col}`)" for col in df.columns if col not in ('id', 'numero_unico')])

    sql = f"INSERT INTO acompanhamento_os_2025 ({colunas_sql}) VALUES ({placeholders}) ON DUPLICATE KEY UPDATE {update_sql}"

    for idx, row in df.iterrows():
        valores = []
        for col in df.columns:
            valor = row[col]
            if isinstance(valor, pd.Timestamp):
                valor = valor.date()
            valores.append(None if pd.isna(valor) else valor)
        try:
            cursor.execute(sql, tuple(valores))
            if cursor.rowcount == 1:
                cont_inseridos += 1
            elif cursor.rowcount == 2:  # ON DUPLICATE UPDATE counts as 2
                cont_atualizados += 1
        except Error as e:
            print(f"⚠ Linha {idx + 2} erro: {e}")
            cont_ignorados += 1
            continue

    conn.commit()

    print("\n📊 Importação concluída!")
    print(f"   - Inseridos: {cont_inseridos}")
    print(f"   - Atualizados: {cont_atualizados}")
    print(f"   - Ignorados: {cont_ignorados}")

except Exception as e:
    print(f"\n❌ Erro geral: {e}")

finally:
    if 'cursor' in locals() and cursor:
        cursor.close()
    if 'conn' in locals() and conn.is_connected():
        conn.close()
        print("🔒 Conexão encerrada.")
