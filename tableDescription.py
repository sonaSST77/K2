from struktura_Dat.db_connect import get_db_connection
import pandas as pd
import os

# Zadej názvy schémat
schemas = ["K2_MIGUSER1", "K2_MIGUSER3"]

try:
    connection = get_db_connection()
    cursor = connection.cursor()

    vystupy = {}
    for schema in schemas:
        # Najdi tabulky, které mají aspoň jeden komentář
        cursor.execute("""
            SELECT DISTINCT table_name
            FROM all_col_comments
            WHERE owner = :schema_name AND comments IS NOT NULL
        """, {"schema_name": schema.upper()})
        tables_with_comment = [row[0] for row in cursor.fetchall()]
        data = []
        for table_name in tables_with_comment:
            # Vypiš všechny sloupce této tabulky, i bez komentáře
            cursor.execute("""
                SELECT column_name, comments
                FROM all_col_comments
                WHERE owner = :schema_name AND table_name = :table_name
                ORDER BY column_name
            """, {"schema_name": schema.upper(), "table_name": table_name})
            columns = cursor.fetchall()
            for column_name, comment in columns:
                data.append({"TABLE_NAME": table_name, "COLUMN_NAME": column_name, "COMMENT": comment})
        df = pd.DataFrame(data, columns=["TABLE_NAME", "COLUMN_NAME", "COMMENT"])
        vystupy[schema] = df

    cursor.close()
    connection.close()

    # Uložení do Excelu
    output_dir = "vystupy"
    os.makedirs(output_dir, exist_ok=True)
    excel_path = os.path.join(output_dir, "tableDescription.xlsx")
    with pd.ExcelWriter(excel_path) as writer:
        for schema, df in vystupy.items():
            df.to_excel(writer, sheet_name=schema, index=False)
    print(f"Výstup uložen do {excel_path}")

except Exception as e:
    print(f"Chyba: {e}")
