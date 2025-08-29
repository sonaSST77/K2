"""
K2_GENERATE_ZAKAZNIK_REPORT.py
Základní struktura programu pro generování reportu zákazníků.
"""
"""
K2_GENERATE_ZAKAZNIK_REPORT.py
Základní struktura programu pro generování reportu zákazníků.
"""

from struktura_Dat.db_connect import get_db_connection
from datetime import datetime, timedelta

def main():
    print("Skript spuštěn")
    try:
        print("Začínám vyhodnocení reportu...")
        print("Připojuji se k databázi...")
        connection = get_db_connection()
        cursor = connection.cursor()
        print("Připojeno.")

        # Počet záznamů se SEVERITY = 'ERRORS' dnes
        print("Zjišťuji počet ERRORS dnes...")
        cursor.execute("SELECT COUNT(*) FROM SO081267.K2_VALIDACE_ZAKAZNIK WHERE UPPER(SEVERITY) = 'ERRORS' AND DATUMREPORTU = :1", [datetime.now().date()])
        count_errors_today = cursor.fetchone()[0]
        print(f"Počet ERRORS dnes: {count_errors_today}")

        # Počet záznamů se SEVERITY = 'WARNINGS' dnes
        print("Zjišťuji počet WARNINGS dnes...")
        cursor.execute("SELECT COUNT(*) FROM SO081267.K2_VALIDACE_ZAKAZNIK WHERE UPPER(SEVERITY) = 'WARNINGS' AND DATUMREPORTU = :1", [datetime.now().date()])
        count_warnings_today = cursor.fetchone()[0]
        print(f"Počet WARNINGS dnes: {count_warnings_today}")

        # Počet záznamů se SEVERITY = 'ERRORS' před týdnem
        last_week = datetime.now().date() - timedelta(days=7)
        print(f"Zjišťuji počet ERRORS před týdnem ({last_week})...")
        cursor.execute("SELECT COUNT(*) FROM SO081267.K2_VALIDACE_ZAKAZNIK WHERE UPPER(SEVERITY) = 'ERRORS' AND DATUMREPORTU = :1", [last_week])
        count_errors_lastweek = cursor.fetchone()[0]
        print(f"Počet ERRORS před týdnem: {count_errors_lastweek}")

        # Počet záznamů se SEVERITY = 'WARNINGS' před týdnem
        print(f"Zjišťuji počet WARNINGS před týdnem ({last_week})...")
        cursor.execute("SELECT COUNT(*) FROM SO081267.K2_VALIDACE_ZAKAZNIK WHERE UPPER(SEVERITY) = 'WARNINGS' AND DATUMREPORTU = :1", [last_week])
        count_warnings_lastweek = cursor.fetchone()[0]
        print(f"Počet WARNINGS před týdnem: {count_warnings_lastweek}")

        print("\n--- Změny oproti minulému týdnu ---")
        error_diff = count_errors_today - count_errors_lastweek
        if error_diff < 0:
            print(f"Opraveno ERRORS: {abs(error_diff)} ({abs(error_diff)/count_errors_lastweek*100:.1f} %)")
        elif error_diff > 0:
            print(f"Přibylo ERRORS: {error_diff} (+{error_diff/count_errors_lastweek*100:.1f} %)")
        else:
            print(f"Počet ERRORS beze změny.")

        warning_diff = count_warnings_today - count_warnings_lastweek
        if warning_diff < 0:
            print(f"Opraveno WARNINGS: {abs(warning_diff)} ({abs(warning_diff)/count_warnings_lastweek*100:.1f} %)")
        elif warning_diff > 0:
            print(f"Přibylo WARNINGS: {warning_diff} (+{warning_diff/count_warnings_lastweek*100:.1f} %)")

        else:
            print(f"Počet WARNINGS beze změny.")


        cursor.close()
        connection.close()
    except Exception as e:
        print(f"Chyba při práci s databází: {e}")

if __name__ == "__main__":
    main()
