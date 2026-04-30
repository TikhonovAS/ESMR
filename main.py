import pandas as pd
import sys
import os
# Добавили импорт generate_word_permits
from src.logic import calculate_maintenance_balanced, export_to_excel, generate_word_permits

# --- НАСТРОЙКИ ---
TARGET_YEAR = 2023  # Поставил 2024, так как мы планировали на него
INPUT_FILE = "data/equipment.xlsx"
# Теперь файл будет сохраняться внутри папки output
OUTPUT_FILE = "output/График_ППР_ESMR.xlsx"


def run():
    print("--- Запуск системы ESMR ---", flush=True)

    try:
        # Проверка наличия папки output
        if not os.path.exists('output'):
            os.makedirs('output')

        input_df = pd.read_excel(INPUT_FILE)

        if input_df.empty:
            print(f"ВНИМАНИЕ: Файл {INPUT_FILE} пуст!", flush=True)
            return

        print(f"Загружено строк: {len(input_df)}", flush=True)

        # 1. Расчет
        schedule = calculate_maintenance_balanced(input_df, target_year=TARGET_YEAR)

        # 2. Экспорт в Excel
        export_to_excel(schedule, OUTPUT_FILE)
        print(f"Excel-файл готов: {OUTPUT_FILE}", flush=True)

        # 3. ГЕНЕРАЦИЯ НАРЯДОВ WORD
        # Эта функция создаст файлы .docx в папке output/Наряды_Допуски
        generate_word_permits(schedule)

        print(f"ГОТОВО! Все документы (Excel и Word) в папке /output", flush=True)

    except Exception as e:
        print(f"КРИТИЧЕСКАЯ ОШИБКА: {e}", flush=True)


if __name__ == "__main__":
    run()
