from datetime import timedelta
import pandas as pd


def calculate_maintenance_schedule(row):
    start_date = pd.to_datetime(row['Last_Repair_Date'])
    schedule = []

    if row['Equipment_Type'] == 'Vessel':  # Емкости и трубы
        # Через 6 месяцев ТО
        next_to = start_date + pd.DateOffset(months=6)
        schedule.append({'Date': next_to, 'Type': 'ТО'})
        # Через год снова ТР
        next_tr = start_date + pd.DateOffset(months=12)
        schedule.append({'Date': next_tr, 'Type': 'ТР'})

    else:  # Другое оборудование
        # 3 ТО через каждые 3 месяца
        for i in range(1, 4):
            next_to = start_date + pd.DateOffset(months=i * 3)
            schedule.append({'Date': next_to, 'Type': f'ТО-{i}'})
        # Через год снова ТР
        next_tr = start_date + pd.DateOffset(months=12)
        schedule.append({'Date': next_tr, 'Type': 'ТР'})

    return schedule
