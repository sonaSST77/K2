import pandas as pd
from struktura_Dat.db_connect import get_db_connection
from datetime import datetime
import warnings

warnings.filterwarnings("ignore", category=UserWarning)

def main():
    try:
        connection = get_db_connection()
        # SELECT pro validace s daty
        query_full = '''
        select 
        D.RESPONSIBLE, 
        D.VAL_ID, 
        D.SEVERITY, 
        (select key from K2_MIGUSER1.l0_crmt_customer where id = F.entity_id) as KEY,
        F.DESCRIPTION, 
        F.DETAILS
        from K2_MIGUSER1.L1_CUSTOMER_CHECK_VALIDATION_ARE F
        join K2_MIGUSER1.msa_val_Design D on F.val_id = D.val_id
        where D.final_query is not null and D.STATUS = 'IMPLEMENTED'
        order by D.responsible DESC, VAL_ID, severity
        '''
        df_full = pd.read_sql(query_full, connection)
        print(f"Načteno {len(df_full)} záznamů s daty.")

        # SELECT pro prázdné validace
        query_empty = '''
        select D.RESPONSIBLE, D.VAL_ID, D.name, D.SEVERITY, 'No errors reported' as DETAILS
        from K2_MIGUSER1.msa_val_Design D 
        where final_query like '%rowcount = 0'
        '''
        df_empty = pd.read_sql(query_empty, connection)
        print(f"Načteno {len(df_empty)} prázdných validací.")

        cursor = connection.cursor()
        today = datetime.now().date()
        # Kontrola, zda už pro dnešní datum existují záznamy
        cursor.execute("SELECT COUNT(*) FROM SO081267.K2_VALIDACE_ZAKAZNIK WHERE DATUMREPORTU = :1", [today])
        count_validace = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM SO081267.K2_VALIDACE_ZAKAZNIK_EMPTY WHERE DATUMREPORTU = :1", [today])
        count_empty = cursor.fetchone()[0]
        if count_validace > 0 or count_empty > 0:
            print(f"Záznamy pro dnešní datum ({today}) již existují. Vkládání bylo přeskočeno.")
        else:
            # Vložení validací s daty
            for _, row in df_full.iterrows():
                cursor.execute(
                    "INSERT INTO SO081267.K2_VALIDACE_ZAKAZNIK (DATUMREPORTU, RESPONSIBLE, VAL_ID, SEVERITY, KEY, DESCRIPTION, DETAILS) VALUES (:1, :2, :3, :4, :5, :6, :7)",
                    [today,
                     str(row['RESPONSIBLE']) if pd.notnull(row['RESPONSIBLE']) else None,
                     str(row['VAL_ID']) if pd.notnull(row['VAL_ID']) else None,
                     str(row['SEVERITY']) if pd.notnull(row['SEVERITY']) else None,
                     str(row['KEY']) if pd.notnull(row['KEY']) else None,
                     str(row['DESCRIPTION']) if pd.notnull(row['DESCRIPTION']) else None,
                     str(row['DETAILS']) if pd.notnull(row['DETAILS']) else None]
                )
            # Vložení prázdných validací
            for _, row in df_empty.iterrows():
                cursor.execute(
                    "INSERT INTO SO081267.K2_VALIDACE_ZAKAZNIK_EMPTY (DATUMREPORTU, RESPONSIBLE, VAL_ID, SEVERITY, DESCRIPTION, NOERRORSREPORTED) VALUES (:1, :2, :3, :4, :5, :6)",
                    [today,
                     str(row['RESPONSIBLE']) if pd.notnull(row['RESPONSIBLE']) else None,
                     str(row['VAL_ID']) if pd.notnull(row['VAL_ID']) else None,
                     str(row['SEVERITY']) if pd.notnull(row['SEVERITY']) else None,
                     str(row['NAME']) if pd.notnull(row['NAME']) else None,
                     'No errors reported']
                )
            connection.commit()
            print("Data byla zapsána do tabulek SO081267.K2_VALIDACE_ZAKAZNIK a SO081267.K2_VALIDACE_ZAKAZNIK_EMPTY.")
        cursor.close()
        connection.close()
    except Exception as e:
        print(f"Chyba: {e}")

if __name__ == "__main__":
    main()
