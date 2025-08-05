# Skript pro načtení pole CHARS pro zadané ID zákazníka z tabulky L0_CRMT_CUSTOMER
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
    cursor.execute("SELECT CHARS FROM K2_MIGUSER1.L0_CRMT_CUSTOMER WHERE ID = :id", {"id": customer_id})
    row = cursor.fetchone()
    if row:
        chars = row[0]
        try:
            parsed = json.loads(chars) if isinstance(chars, str) else chars
            if isinstance(parsed, dict):
                print("CHARS (slovník):")
                for k, v in parsed.items():
                    print(f"  {k}: {v}")
            elif isinstance(parsed, list):
                print("CHARS (seznam):")
                for i, item in enumerate(parsed, 1):
                    print(f"  {i}. {item}")
            else:
                print("CHARS:", parsed)
        except Exception as e:
            print("CHARS (nelze dekódovat jako JSON):", chars)
            print("Chyba:", e)
    else:
        print(f"Zákazník s ID {customer_id} nebyl nalezen.")
    cursor.close()
    conn.close()
