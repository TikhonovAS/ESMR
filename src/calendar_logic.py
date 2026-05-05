import pandas as pd


def calculate_maintenance_schedule(row):
    """
    Вспомогательная функция для расчета базовых интервалов.
    Основная логика распределения теперь находится в logic.py
    """
    start_date = pd.to_datetime(row.get('last_repair_date', pd.Timestamp.now()))
    schedule = []

    # Пример базовой логики, если нужно использовать отдельно
    if row.get('equipment_type') == 'Vessel':
        next_to = start_date + pd.DateOffset(months=6)
        schedule.append({'Date': next_to, 'Type': 'ТО'})

    return schedule
