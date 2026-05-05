import pandas as pd
from datetime import datetime, timedelta
import holidays
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict
import os
import re
import shutil
from docxtpl import DocxTemplate

# --- КОНСТАНТЫ ---
MONTH_NAMES_RU = {
    1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
    5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
    9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
}

WORK_CATALOG = {
    'vessel': {'prio': 1, 'label': 'Емкость/Резервуар',
               'ТР': 'Вскрытие, дегазация, очистка, дефектовка швов, толщинометрия, замена прокладок люков, испытания.',
               'ТО': 'Наружный осмотр, проверка дыхательной арматуры, контроль заземления.',
               'hours': {'ТР': 72, 'ТО': 8}, 'days': 3},
    'heatex': {'prio': 2, 'label': 'Теплообменник',
               'ТР': 'Разборка, механическая очистка трубных пучков, проверка на герметичность, замена уплотнений.',
               'ТО': 'Осмотр, контроль температуры/давления, проверка отсутствия протечек.',
               'hours': {'ТР': 64, 'ТО': 12}, 'days': 3},
    'pump': {'prio': 3, 'label': 'Насос',
             'ТР': 'Центровка валов, разборка, замена подшипников и уплотнений, ревизия муфты.',
             'ТО': 'Проверка масла, контроль вибрации и шума, подтяжка болтов.', 'hours': {'ТР': 18, 'ТО': 4},
             'days': 1},
    'fan': {'prio': 4, 'label': 'Вентилятор',
            'ТР': 'Балансировка лопастей, замена подшипников, проверка натяжения ремней.',
            'ТО': 'Очистка лопастей, контроль вибрации, проверка ограждений.', 'hours': {'ТР': 16, 'ТО': 4}, 'days': 1},
    'pipe': {'prio': 5, 'label': 'Трубопровод',
             'ТР': 'Ревизия опорожнения, замена дефектных участков, проверка швов, восстановление изоляции.',
             'ТО': 'Обход трассы, проверка состояния опор, контроль герметичности.', 'hours': {'ТР': 14, 'ТО': 2},
             'days': 1}
}

HOLIDAYS_SET = set()


def clear_output_folder(folder_path):
    if os.path.exists(folder_path):
        shutil.rmtree(folder_path)
    os.makedirs(folder_path, exist_ok=True)


def get_tatarstan_holidays(year):
    ru_hols = holidays.Russia(years=year)
    all_hols = {d.strftime('%Y-%m-%d') for d in ru_hols}
    all_hols.update([f"{year}-08-30", f"{year}-11-06"])
    return all_hols


def is_workday(date):
    return date.weekday() < 5 and date.strftime('%Y-%m-%d') not in HOLIDAYS_SET


def get_end_date(start_date, duration_days):
    if duration_days <= 1: return start_date
    curr, count = start_date, 1
    while count < duration_days:
        curr += timedelta(days=1)
        if is_workday(curr): count += 1
    return curr


def get_all_workdays_in_month(year, month):
    workdays = []
    for d in range(1, 32):
        try:
            curr = datetime(year, month, d)
            if is_workday(curr): workdays.append(curr)
        except ValueError:
            break
    return workdays


def get_prev_repair_date(row, df_columns, target_year):
    for m_num in range(1, 13):
        m_name = MONTH_NAMES_RU[m_num].lower()
        for col in df_columns:
            if m_name in col:
                val = str(row.get(col, ""))
                if "ТР" in val:
                    match = re.search(r'(\d{2})\.(\d{2})', val)
                    if match: return datetime(target_year - 1, int(match.group(2)), int(match.group(1)))
    return datetime(target_year - 1, 1, 1)


def calculate_maintenance_balanced(df, target_year):
    global HOLIDAYS_SET
    HOLIDAYS_SET = get_tatarstan_holidays(target_year)
    df.columns = [str(c).lower().strip() for c in df.columns]
    slots = defaultdict(int)

    res = []
    for idx, row in df.iterrows():
        name = str(row.get('наименование', row.get('оборудование', '')))
        cat = 'vessel' if 'емкост' in name.lower() or 'сосуд' in name.lower() else 'pump' if 'насос' in name.lower() else 'pipe'
        cfg = WORK_CATALOG.get(cat, WORK_CATALOG['pipe'])

        last_done_date = get_prev_repair_date(row, df.columns, target_year)
        m_tr = (idx % 6) + 4

        w_days = get_all_workdays_in_month(target_year, m_tr)
        if not w_days: continue

        date_tr = w_days[slots[m_tr] % len(w_days)]
        slots[m_tr] += 1
        end_tr = get_end_date(date_tr, cfg['days'])

        jobs_list = []
        jobs_list.append(
            {'date': date_tr, 'end_date': end_tr, 'type': 'ТР', 'desc': cfg['ТР'], 'hours': cfg['hours']['ТР'],
             'day': date_tr.day})

        step = 3 if cat in ['pump', 'fan'] else 6
        for off in [-9, -6, -3, 3, 6, 9]:
            if abs(off) % step != 0: continue
            ideal_m = date_tr.month + off
            if 1 <= ideal_m <= 12 and ideal_m != m_tr:
                to_w_days = get_all_workdays_in_month(target_year, ideal_m)
                if to_w_days:
                    to_d = to_w_days[slots[m_tr] % len(to_w_days)]
                    jobs_list.append(
                        {'date': to_d, 'end_date': to_d, 'type': 'ТО', 'desc': cfg['ТО'], 'hours': cfg['hours']['ТО'],
                         'day': to_d.day})

        jobs_list.sort(key=lambda x: x['date'])
        final_jobs, current_base = {}, last_done_date
        for j in jobs_list:
            j['op_hours'] = (j['date'] - current_base).days * 24
            final_jobs[j['date'].month] = j
            if "ТР" in j['type']: current_base = j['end_date']

        res.append({
            'name': name, 'marka': row.get('марка', '—'),
            'zav_no': row.get('зав. №', row.get('зав', '—')),
            'poz_no': row.get('поз. №', row.get('поз', '—')),
            'jobs': final_jobs, 'cat': cat
        })
    return res


def export_to_excel(schedule, output_path):
    brd = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        yr_rows = []
        for it in schedule:
            r = {'Оборудование': it['name'], 'Марка': it['marka'], 'Зав. №': it['zav_no'], 'Поз. №': it['poz_no']}
            for m in range(1, 13):
                r[MONTH_NAMES_RU[m]] = f"{it['jobs'][m]['type']} ({it['jobs'][m]['date'].strftime('%d.%m')})" if m in \
                                                                                                                 it[
                                                                                                                     'jobs'] else ""
            yr_rows.append(r)
        pd.DataFrame(yr_rows).to_excel(writer, index=False, sheet_name='Годовой график')

        target_year = schedule[0]['jobs'][next(iter(schedule[0]['jobs']))]['date'].year
        for m in range(1, 13):
            m_name = MONTH_NAMES_RU[m]
            p_data = []
            for it in schedule:
                if m in it['jobs']:
                    j = it['jobs'][m]
                    p_data.append({
                        'Начало': j['date'].strftime('%d.%m.%Y'),
                        'Окончание': j['end_date'].strftime('%d.%m.%Y'),
                        'Оборудование': it['name'], 'Марка': it['marka'],
                        'Зав. №': it['zav_no'], 'Поз. №': it['poz_no'],
                        'Наработка (ч)': j['op_hours'], 'Вид': j['type']
                    })
            if p_data:
                pd.DataFrame(p_data).sort_values('Начало').to_excel(writer, index=False, sheet_name=f"План_{m_name}")

                ws_c = writer.book.create_sheet(title=f"Календарь_{m_name}")
                last_d = (pd.Timestamp(target_year, m, 1) + pd.offsets.MonthEnd(0)).day
                ws_c.append(['Оборудование', 'Марка', 'Зав. №', 'Поз. №'] + list(range(1, last_d + 1)))
                for it in schedule:
                    if m in it['jobs']:
                        j = it['jobs'][m]
                        row = [it['name'], it['marka'], it['zav_no'], it['poz_no']] + [""] * last_d
                        row[j['day'] + 3] = j['type']
                        ws_c.append(row)

        for sn in writer.sheets:
            ws = writer.sheets[sn]
            for row in ws.iter_rows():
                for cell in row:
                    cell.border = brd
                    cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
                    if cell.row == 1: cell.font = Font(bold=True)
            for col_cells in ws.columns:
                col_letter = get_column_letter(col_cells[0].column)
                if "Календарь" in sn and col_cells[0].column > 4:
                    ws.column_dimensions[col_letter].width = 4
                else:
                    ws.column_dimensions[col_letter].width = 18


def generate_word_reports(schedule):
    output_dir = 'output/Ведомости_работ'
    clear_output_folder(output_dir)
    template_path = 'templates/template_works.docx'  # Нужно создать этот шаблон
    if not os.path.exists(template_path):
        print("ВНИМАНИЕ: Шаблон templates/template_works.docx не найден.")
        return

    for it in schedule:
        for m, j in it['jobs'].items():
            doc = DocxTemplate(template_path)
            ctx = {
                'date': j['date'].strftime('%d.%m.%Y'),
                'name': it['name'], 'marka': it['marka'],
                'zav_no': it['zav_no'], 'poz_no': it['poz_no'],
                'type': j['type'], 'op_hours': j['op_hours'],
                'works': j['desc']
            }
            doc.render(ctx)
            safe_name = re.sub(r'[\\/*?:"<>|]', "_", f"{it['poz_no']}_{j['type']}_{MONTH_NAMES_RU[m]}")
            doc.save(os.path.join(output_dir, f"{safe_name}.docx"))
