import requests
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# =========================================================
# НАСТРОЙКИ
# =========================================================

START_DATE = input("Стартовая дата в формате гггг-мм-дд")
END_DATE = input("Конечная дата в формате гггг-мм-дд")

INDEXES = {
    "IMOEX": {
        "ticker": "IMOEX",
        "name": "Индекс Мосбиржи IMOEX",
    },
    "RTSI": {
        "ticker": "RTSI",
        "name": "Индекс Мосбиржи РТС (RTSI)",
    },
    "RGBITR": {
        "ticker": "RGBITR",
        "name": "Индекс Мосбиржи государственных облигаций (RGBITR)",
    },
    "RUCBTRNS": {
        "ticker": "RUCBTRNS",
        "name": "Индекс Мосбиржи корпоративных облигаций (RUCBTRNS)",
    },
}

BASE_URL = "https://iss.moex.com/iss/history/engines/stock/markets/index/securities"


# =========================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =========================================================

def format_number(value, digits=2):
    if pd.isna(value):
        return "-"
    return f"{value:,.{digits}f}".replace(",", " ").replace(".", ",")


def format_percent(value):
    if pd.isna(value):
        return "-"

    sign = "+" if value > 0 else ""
    return f"{sign}{value:.2f}%".replace(".", ",")


def short_date(date_str):
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    weekdays = {
        0: "пн",
        1: "вт",
        2: "ср",
        3: "чт",
        4: "пт",
        5: "сб",
        6: "вс",
    }
    return f"{dt.strftime('%d.%m')} ({weekdays[dt.weekday()]})"


# =========================================================
# ЗАГРУЗКА ДАННЫХ
# =========================================================

def load_data(ticker):
    url = f"{BASE_URL}/{ticker}.json"

    params = {
        "from": START_DATE,
        "till": END_DATE,
        "iss.meta": "off",
    }

    response = requests.get(url, params=params)
    response.raise_for_status()

    data = response.json()

    columns = data["history"]["columns"]
    rows = data["history"]["data"]

    df = pd.DataFrame(rows, columns=columns)

    numeric_columns = [
        "OPEN",
        "LOW",
        "HIGH",
        "CLOSE",
        "VALUE",
        "DURATION",
        "YIELD",
    ]

    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


# =========================================================
# ОБРАБОТКА ДАННЫХ
# =========================================================

def prepare_dataframe(df):
    df = df.copy()

    df["CHANGE"] = df["CLOSE"].pct_change() * 100

    return df


def weekly_change(df):
    first_close = df.iloc[0]["CLOSE"]
    last_close = df.iloc[-1]["CLOSE"]

    return ((last_close / first_close) - 1) * 100


def find_max(df):
    idx = df["HIGH"].idxmax()

    return (
        df.loc[idx, "HIGH"],
        df.loc[idx, "TRADEDATE"]
    )


def find_min(df):
    idx = df["LOW"].idxmin()

    return (
        df.loc[idx, "LOW"],
        df.loc[idx, "TRADEDATE"]
    )


# =========================================================
# WORD
# =========================================================

def set_font(run, size=10, bold=False):
    run.font.size = Pt(size)
    run.bold = bold
    run.font.name = "Times New Roman"


def add_title(document, text):
    p = document.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    run = p.add_run(text)
    set_font(run, 14, True)


def add_heading(document, text):
    p = document.add_paragraph()

    run = p.add_run(text)
    set_font(run, 12, True)


def add_table(document, headers, rows):
    table = document.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"

    hdr_cells = table.rows[0].cells

    for i, header in enumerate(headers):
        p = hdr_cells[i].paragraphs[0]
        run = p.add_run(header)

        set_font(run, 10, True)

    for row in rows:
        row_cells = table.add_row().cells

        for i, value in enumerate(row):
            p = row_cells[i].paragraphs[0]
            run = p.add_run(str(value))

            set_font(run, 10)

    document.add_paragraph()


# =========================================================
# ОСНОВНАЯ ЛОГИКА
# =========================================================

all_data = {}

for key, info in INDEXES.items():
    df = load_data(info["ticker"])
    df = prepare_dataframe(df)

    all_data[key] = df

    display_data = {}

    for key, df in all_data.items():
        display_data[key] = df.iloc[1:].reset_index(drop=True)


# =========================================================
# СОЗДАНИЕ WORD
# =========================================================

doc = Document()

add_title(doc, "Отчёт по индексам MOEX")


# =========================================================
# ПОДРОБНЫЕ ТАБЛИЦЫ
# =========================================================

for key, info in INDEXES.items():

    df = display_data[key]

    add_heading(doc, info["name"])

    headers = [
        "Дата",
        "open",
        "high",
        "low",
        "close",
        "Изм. %",
    ]

    if key == "IMOEX":
        headers.append("Объём, млрд ₽")

    elif key == "RTSI":
        headers.append("Объём, млн $")

    else:
        headers.extend([
            "Дюрация, дн.",
            "YIELD, %",
            "Объём, млрд ₽",
        ])

    rows = []

    for _, row in df.iterrows():

        values = [
            short_date(row["TRADEDATE"]),
            format_number(row["OPEN"]),
            format_number(row["HIGH"]),
            format_number(row["LOW"]),
            format_number(row["CLOSE"]),
            format_percent(row["CHANGE"]),
        ]

        # IMOEX
        if key == "IMOEX":

            values.append(
                format_number(
                    row["VALUE"] / 1_000_000_000
                )
            )

        # RTSI
        elif key == "RTSI":

            values.append(
                format_number(
                    row["VALUE"] / 1_000_000
                )
            )

        # RGBITR / RUCBTRNS
        else:

            values.extend([
                format_number(row["DURATION"], 0),
                format_number(row["YIELD"]),
                format_number(
                    row["VALUE"] / 1_000_000_000
                ),
            ])

        rows.append(values)

    add_table(doc, headers, rows)

# =========================================================
# ИТОГОВАЯ СВОДНАЯ ТАБЛИЦА
# =========================================================

add_heading(doc, "Сводная таблица")

summary_headers = [
    "Дата",
    "IMOEX",
    "RTSI",
    "RGBITR",
    "RUCBTRNS",
]

summary_rows = []

for i in range(len(display_data["IMOEX"])):

    imoex = display_data["IMOEX"].iloc[i]
    rtsi = display_data["RTSI"].iloc[i]
    rgbitr = display_data["RGBITR"].iloc[i]
    rucb = display_data["RUCBTRNS"].iloc[i]

    row = [
        short_date(imoex["TRADEDATE"]),

        f"{format_number(imoex['CLOSE'])} ({format_percent(imoex['CHANGE'])})",

        f"{format_number(rtsi['CLOSE'])} ({format_percent(rtsi['CHANGE'])})",

        f"{format_number(rgbitr['CLOSE'])} ({format_percent(rgbitr['CHANGE'])})",

        f"{format_number(rucb['CLOSE'])} ({format_percent(rucb['CHANGE'])})",
    ]

    summary_rows.append(row)

summary_rows.append([
    "Итог недели",
    format_percent(weekly_change(all_data["IMOEX"])),
    format_percent(weekly_change(all_data["RTSI"])),
    format_percent(weekly_change(all_data["RGBITR"])),
    format_percent(weekly_change(all_data["RUCBTRNS"])),
])

imoex_max, imoex_max_date = find_max(display_data["IMOEX"])
rtsi_max, rtsi_max_date = find_max(display_data["RTSI"])
rgbitr_max, rgbitr_max_date = find_max(display_data["RGBITR"])
rucb_max, rucb_max_date = find_max(display_data["RUCBTRNS"])

summary_rows.append([
    "Максимум",
    f"{format_number(imoex_max)} ({short_date(imoex_max_date)})",
    f"{format_number(rtsi_max)} ({short_date(rtsi_max_date)})",
    f"{format_number(rgbitr_max)} ({short_date(rgbitr_max_date)})",
    f"{format_number(rucb_max)} ({short_date(rucb_max_date)})",
])

imoex_min, imoex_min_date = find_min(display_data["IMOEX"])
rtsi_min, rtsi_min_date = find_min(display_data["RTSI"])
rgbitr_min, rgbitr_min_date = find_min(display_data["RGBITR"])
rucb_min, rucb_min_date = find_min(display_data["RUCBTRNS"])

summary_rows.append([
    "Минимум",
    f"{format_number(imoex_min)} ({short_date(imoex_min_date)})",
    f"{format_number(rtsi_min)} ({short_date(rtsi_min_date)})",
    f"{format_number(rgbitr_min)} ({short_date(rgbitr_min_date)})",
    f"{format_number(rucb_min)} ({short_date(rucb_min_date)})",
])

add_table(doc, summary_headers, summary_rows)

# =========================================================
# СОХРАНЕНИЕ
# =========================================================

filename = "moex_report.docx"

doc.save(filename)

print(f"Файл сохранён: {filename}")