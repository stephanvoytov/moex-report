import requests
import pandas as pd
from datetime import datetime, timedelta
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# =========================================================
# НАСТРОЙКИ
# =========================================================

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

START_DATE = ""
END_DATE = ""


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
    if date_str is None:
        return "-"
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

    try:
        response = requests.get(url, params=params, timeout=30)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Ошибка загрузки данных для {ticker}: {e}")
        return pd.DataFrame()

    try:
        data = response.json()
        columns = data["history"]["columns"]
        rows = data["history"]["data"]
        df = pd.DataFrame(rows, columns=columns)
    except (KeyError, ValueError) as e:
        print(f"Ошибка обработки ответа для {ticker}: {e}")
        return pd.DataFrame()

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
    if df.empty:
        return df

    df = df.copy()
    df["CHANGE"] = df["CLOSE"].pct_change() * 100
    return df


def weekly_change(df):
    if df.empty:
        return None

    first_close = df.iloc[0]["CLOSE"]
    last_close = df.iloc[-1]["CLOSE"]
    return ((last_close / first_close) - 1) * 100


def find_max(df):
    if df.empty:
        return None, None

    idx = df["HIGH"].idxmax()
    return df.loc[idx, "HIGH"], df.loc[idx, "TRADEDATE"]


def find_min(df):
    if df.empty:
        return None, None

    idx = df["LOW"].idxmin()
    return df.loc[idx, "LOW"], df.loc[idx, "TRADEDATE"]


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

def main():
    all_data = {}

    for key, info in INDEXES.items():
        df = load_data(info["ticker"])
        df = prepare_dataframe(df)
        all_data[key] = df

    display_data = {key: df.reset_index(drop=True) for key, df in all_data.items()}

    doc = Document()
    add_title(doc, "Отчёт по индексам MOEX")

    # ПОДРОБНЫЕ ТАБЛИЦЫ
    for key, info in INDEXES.items():
        if display_data[key].empty:
            print(f"Нет данных для индекса {info['name']}")
            continue

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

            if key == "IMOEX":
                values.append(format_number(row["VALUE"] / 1_000_000_000))
            elif key == "RTSI":
                values.append(format_number(row["VALUE"] / 1_000_000))
            else:
                values.extend([
                    format_number(row["DURATION"], 0),
                    format_number(row["YIELD"]),
                    format_number(row["VALUE"] / 1_000_000_000),
                ])

            rows.append(values)

        add_table(doc, headers, rows)

    # ИТОГОВАЯ СВОДНАЯ ТАБЛИЦА
    add_heading(doc, "Сводная таблица")

    summary_headers = [
        "Дата",
        "IMOEX",
        "RTSI",
        "RGBITR",
        "RUCBTRNS",
    ]

    summary_rows = []

    min_len = min(len(display_data[key]) for key in INDEXES)

    if min_len > 0:
        for i in range(min_len):
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

    imoex_max, imoex_max_date = find_max(all_data["IMOEX"])
    rtsi_max, rtsi_max_date = find_max(all_data["RTSI"])
    rgbitr_max, rgbitr_max_date = find_max(all_data["RGBITR"])
    rucb_max, rucb_max_date = find_max(all_data["RUCBTRNS"])

    summary_rows.append([
        "Максимум",
        f"{format_number(imoex_max)} ({short_date(imoex_max_date)})",
        f"{format_number(rtsi_max)} ({short_date(rtsi_max_date)})",
        f"{format_number(rgbitr_max)} ({short_date(rgbitr_max_date)})",
        f"{format_number(rucb_max)} ({short_date(rucb_max_date)})",
    ])

    imoex_min, imoex_min_date = find_min(all_data["IMOEX"])
    rtsi_min, rtsi_min_date = find_min(all_data["RTSI"])
    rgbitr_min, rgbitr_min_date = find_min(all_data["RGBITR"])
    rucb_min, rucb_min_date = find_min(all_data["RUCBTRNS"])

    summary_rows.append([
        "Минимум",
        f"{format_number(imoex_min)} ({short_date(imoex_min_date)})",
        f"{format_number(rtsi_min)} ({short_date(rtsi_min_date)})",
        f"{format_number(rgbitr_min)} ({short_date(rgbitr_min_date)})",
        f"{format_number(rucb_min)} ({short_date(rucb_min_date)})",
    ])

    add_table(doc, summary_headers, summary_rows)

    # СОХРАНЕНИЕ
    now = datetime.now()
    filename = f"moex_report_{now:%Y%m%d_%H%M}.docx"

    try:
        doc.save(filename)
        print(f"Файл сохранён: {filename}")
    except Exception as e:
        print(f"Ошибка сохранения файла: {e}")


if __name__ == "__main__":
    start_input = input("Стартовая дата (гггг-мм-дд): ").strip()
    end_input = input("Конечная дата (гггг-мм-дд): ").strip()

    for inp in (start_input, end_input):
        try:
            datetime.strptime(inp, "%Y-%m-%d")
        except ValueError:
            print(f"Ошибка: дата '{inp}' не соответствует формату гггг-мм-дд")
            exit(1)

    if start_input > end_input:
        print("Ошибка: стартовая дата не может быть позже конечной")
        exit(1)

    START_DATE = start_input
    END_DATE = end_input

    main()
