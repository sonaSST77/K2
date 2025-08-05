from struktura_Dat.db_connect import get_db_connection
import pandas as pd
import os

# Zadej názvy schémat, která chceš prohledat
schema1 = "K2_MIGUSER1"

try:
    connection = get_db_connection()
    cursor = connection.cursor()

    vystupy = {}
    for schema in [schema1]:
        cursor.execute("""
            SELECT t.table_name, c.comments
            FROM all_tables t
            LEFT JOIN all_tab_comments c
              ON t.owner = c.owner AND t.table_name = c.table_name
            WHERE t.owner = :schema
            ORDER BY t.table_name
        """, {"schema": schema.upper()})
        tables = cursor.fetchall()
        data = []
        for table_name, comment in tables:
            # Zjisti počet záznamů v tabulce
            try:
                cursor.execute(f'SELECT COUNT(*) FROM "{schema}"."{table_name}"')
                count = cursor.fetchone()[0]
            except Exception as e:
                count = None
            data.append({"TABLE_NAME": table_name, "COMMENT": comment, "ROW_COUNT": count})
        df = pd.DataFrame(data, columns=["TABLE_NAME", "COMMENT", "ROW_COUNT"])
        vystupy[schema] = df

    cursor.close()
    connection.close()

    # Uložení do Excelu
    output_dir = "vystupy"
    os.makedirs(output_dir, exist_ok=True)
    excel_path = os.path.join(output_dir, "struktura_Tabulek.xlsx")
    with pd.ExcelWriter(excel_path) as writer:
        for schema, df in vystupy.items():
            df.to_excel(writer, sheet_name=schema, index=False)
    print(f"Výstup uložen do {excel_path}")

except Exception as e:
    print(f"Chyba: {e}")
