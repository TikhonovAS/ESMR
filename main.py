import pandas as pd
import os
from src.logic import calculate_maintenance_balanced, export_to_excel, generate_word_reports

TARGET_YEAR = 2026
INPUT_FILE = f"output/График_ППР_{TARGET_YEAR - 1}_г.xlsx"
OUTPUT_FILE = f"output/График_ППР_{TARGET_YEAR}_г.xlsx"


def run():
    print(f"--- ESMR: Запуск планирования на {TARGET_YEAR} год ---")
    if not os.path.exists(INPUT_FILE):
        print(f"Файл {INPUT_FILE} не найден!")
        return

    try:
        input_df = pd.read_excel(INPUT_FILE, sheet_name='Годовой график')
        schedule = calculate_maintenance_balanced(input_df, TARGET_YEAR)

        # 1. Экспорт Excel (Марка возвращена, Содержание убрано)
        export_to_excel(schedule, OUTPUT_FILE)

        # 2. Генерация Word (Ведомости работ отдельно)
        generate_word_reports(schedule)

        print(f"УСПЕХ: Отчеты в /output")
    except Exception as e:
        print(f"Ошибка: {e}")


if __name__ == "__main__":
    run()
