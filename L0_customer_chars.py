# Skript pro načtení pole CHARS pro zadané ID zákazníka z tabulky L0_CRMT_CUSTOMER a napojení na L0_CRMT_PERSON
# Použití: python L0_customer_chars.py <ID_zakaznika>
import sys
import oracledb
from struktura_Dat.db_connect import get_db_connection
import json

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Použití: python L0_customer_chars.py <ID_zakaznika>")
        sys.exit(1)
    customer_id = sys.argv[1]
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT c.CHARS as CUSTOMER_CHARS, p.*, p.CHARS as PERSON_CHARS
        FROM K2_MIGUSER1.L0_CRMT_CUSTOMER c
        JOIN K2_MIGUSER1.L0_CRMT_PERSON p ON c.ID = p.CUSTOMER_ID
        WHERE c.ID = :id
    """, {"id": customer_id})
    rows = cursor.fetchall()
    if rows:
        # Získání názvů sloupců
        columns = [desc[0] for desc in cursor.description]
        for idx, row in enumerate(rows, 1):
            print(f"\n--- Záznam osoby {idx} ---")
            customer_chars = row[0]
            person_data = row[1:]
            person_columns = columns[1:]
            # Výpis CUSTOMER_CHARS
            print("CUSTOMER_CHARS:")
            try:
                parsed = json.loads(customer_chars) if isinstance(customer_chars, str) else customer_chars
                if isinstance(parsed, dict):
                    for k, v in parsed.items():
                        print(f"  {k}: {v}")
                elif isinstance(parsed, list):
                    for i, item in enumerate(parsed, 1):
                        print(f"  {i}. {item}")
                else:
                    print(parsed)
            except Exception as e:
                print(f"  (nelze dekódovat jako JSON): {customer_chars}")
                print(f"  Chyba: {e}")
            # Výpis údajů z tabulky L0_CRMT_PERSON
            print("Údaje z L0_CRMT_PERSON:")
            for col, val in zip(person_columns, person_data):
                if col == "CHARS" or col == "PERSON_CHARS":
                    print(f"  {col}:")
                    try:
                        parsed_person = json.loads(val) if isinstance(val, str) else val
                        if isinstance(parsed_person, dict):
                            for k, v in parsed_person.items():
                                print(f"    {k}: {v}")
                        elif isinstance(parsed_person, list):
                            for i, item in enumerate(parsed_person, 1):
                                print(f"    {i}. {item}")
                        else:
                            print(f"    {parsed_person}")
                    except Exception as e:
                        print(f"    (nelze dekódovat jako JSON): {val}")
                        print(f"    Chyba: {e}")
                else:
                    print(f"  {col}: {val}")
    else:
        print(f"Zákazník s ID {customer_id} nebyl nalezen.")
    cursor.close()
    conn.close()
