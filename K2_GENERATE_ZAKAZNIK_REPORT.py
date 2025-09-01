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
    print("Začínám vyhodnocení reportu...")
    print("Připojuji se k databázi...")
    try:
        connection = get_db_connection()
        cursor = connection.cursor()
        print("Připojeno.")

        # Zadání dat od uživatele
        date_compare_str = input("Zadejte porovnávací datum (např. před týdnem) ve formátu YYYY-MM-DD: ")
        try:
            date_today = datetime.now().date()
            date_compare = datetime.strptime(date_compare_str, "%Y-%m-%d").date()
        except Exception as e:
            print(f"Chybný formát data: {e}")
            return

        # Počet záznamů se SEVERITY = 'ERRORS' pro dnešní datum
        print(f"Zjišťuji počet ERRORS pro {date_today}...")
        cursor.execute("SELECT COUNT(*) FROM SO081267.K2_VALIDACE_ZAKAZNIK WHERE UPPER(SEVERITY) = 'ERRORS' AND DATUMREPORTU = :1", [date_today])
        count_errors_today = cursor.fetchone()[0]
        print(f"Počet ERRORS pro {date_today}: {count_errors_today}")

        # Počet záznamů se SEVERITY = 'WARNINGS' pro dnešní datum
        print(f"Zjišťuji počet WARNINGS pro {date_today}...")
        cursor.execute("SELECT COUNT(*) FROM SO081267.K2_VALIDACE_ZAKAZNIK WHERE UPPER(SEVERITY) = 'WARNINGS' AND DATUMREPORTU = :1", [date_today])
        count_warnings_today = cursor.fetchone()[0]
        print(f"Počet WARNINGS pro {date_today}: {count_warnings_today}")

        # Počet záznamů se SEVERITY = 'ERRORS' pro porovnávací datum
        print(f"Zjišťuji počet ERRORS pro {date_compare}...")
        cursor.execute("SELECT COUNT(*) FROM SO081267.K2_VALIDACE_ZAKAZNIK WHERE UPPER(SEVERITY) = 'ERRORS' AND DATUMREPORTU = :1", [date_compare])
        count_errors_lastweek = cursor.fetchone()[0]
        print(f"Počet ERRORS pro {date_compare}: {count_errors_lastweek}")

        # Počet záznamů se SEVERITY = 'WARNINGS' pro porovnávací datum
        print(f"Zjišťuji počet WARNINGS pro {date_compare}...")
        cursor.execute("SELECT COUNT(*) FROM SO081267.K2_VALIDACE_ZAKAZNIK WHERE UPPER(SEVERITY) = 'WARNINGS' AND DATUMREPORTU = :1", [date_compare])
        count_warnings_lastweek = cursor.fetchone()[0]
        print(f"Počet WARNINGS pro {date_compare}: {count_warnings_lastweek}")

        print("\n--- Změny oproti minulému týdnu ---")
        error_diff = count_errors_today - count_errors_lastweek
        if error_diff < 0:
            if count_errors_lastweek == 0:
                print(f"Opraveno ERRORS: {abs(error_diff)} (minulý týden bylo 0)")
            else:
                print(f"Opraveno ERRORS: {abs(error_diff)} ({abs(error_diff)/count_errors_lastweek*100:.1f} %)")
        elif error_diff > 0:
            if count_errors_lastweek == 0:
                print(f"Přibylo ERRORS: {error_diff} (minulý týden bylo 0)")
            else:
                print(f"Přibylo ERRORS: {error_diff} (+{error_diff/count_errors_lastweek*100:.1f} %)")
        else:
            print(f"Počet ERRORS beze změny.")

        warning_diff = count_warnings_today - count_warnings_lastweek
        if warning_diff < 0:
            if count_warnings_lastweek == 0:
                print(f"Opraveno WARNINGS: {abs(warning_diff)} (minulý týden bylo 0)")
            else:
                print(f"Opraveno WARNINGS: {abs(warning_diff)} ({abs(warning_diff)/count_warnings_lastweek*100:.1f} %)")
        elif warning_diff > 0:
            if count_warnings_lastweek == 0:
                print(f"Přibylo WARNINGS: {warning_diff} (minulý týden bylo 0)")
            else:
                print(f"Přibylo WARNINGS: {warning_diff} (+{warning_diff/count_warnings_lastweek*100:.1f} %)")
        else:
            print(f"Počet WARNINGS beze změny.")

        # Export do Excelu
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Status dat"
        ws.append([f"Souhrn stavu čistoty dat za týden: {date_compare} - {date_today}"])
        ws.append([])
        ws.append(["Datum", "Počet ERRORS", "Počet WARNINGS"])
        ws.append([str(date_today), count_errors_today, count_warnings_today])
        ws.append([str(date_compare), count_errors_lastweek, count_warnings_lastweek])
        ws.append([])
        ws.append(["Změny oproti minulému týdnu:"])
        if error_diff < 0:
            if count_errors_lastweek == 0:
                ws.append([f"Opraveno ERRORS: {abs(error_diff)} (minulý týden bylo 0)"])
            else:
                ws.append([f"Opraveno ERRORS: {abs(error_diff)} ({abs(error_diff)/count_errors_lastweek*100:.1f} %)"])
        elif error_diff > 0:
            if count_errors_lastweek == 0:
                ws.append([f"Přibylo ERRORS: {error_diff} (minulý týden bylo 0)"])
            else:
                ws.append([f"Přibylo ERRORS: {error_diff} (+{error_diff/count_errors_lastweek*100:.1f} %)"])
        else:
            ws.append(["Počet ERRORS beze změny."])

        if warning_diff < 0:
            if count_warnings_lastweek == 0:
                ws.append([f"Opraveno WARNINGS: {abs(warning_diff)} (minulý týden bylo 0)"])
            else:
                ws.append([f"Opraveno WARNINGS: {abs(warning_diff)} ({abs(warning_diff)/count_warnings_lastweek*100:.1f} %)"])
        elif warning_diff > 0:
            if count_warnings_lastweek == 0:
                ws.append([f"Přibylo WARNINGS: {warning_diff} (minulý týden bylo 0)"])
            else:
                ws.append([f"Přibylo WARNINGS: {warning_diff} (+{warning_diff/count_warnings_lastweek*100:.1f} %)"])
        else:
            ws.append(["Počet WARNINGS beze změny."])

        wb.save("K2_status_cistota_dat.xls")
        print("Výsledky byly uloženy do K2_status_cistota_dat.xls")

        cursor.close()
        connection.close()
    except Exception as e:
        print(f"Chyba při práci s databází: {e}")
    # ...konec validní funkce main()...

if __name__ == "__main__":
    main()
