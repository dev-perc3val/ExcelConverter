import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import os

INPUT_FILE = r"C:\Users\Admin\Downloads\260902_gre2p_CH2F5_NADPH_absorbance.xlsx"
OUTPUT_FILE = r"C:\Users\Admin\Desktop\UNI KIT\260902_gre2p_NADPH_340_clean.xlsx"
SHEET_NAME = "Plate 1 - Sheet1"

LAYOUT_HEADER_ROW = 32
LAYOUT_FIRST_ROW = 33

MEASUREMENTS = {
    "Measurement 1": ["B", "E"],
    "Measurement 2": ["C", "F"],
    "Measurement 3": ["D", "G"],
}


def load_workbook(filename):
    return openpyxl.load_workbook(filename, data_only=True)


def read_plate_layout(ws):
    layout = {}
    column_numbers = {
        col: ws.cell(row=LAYOUT_HEADER_ROW, column=col).value
        for col in range(1, ws.max_column + 1)
        if isinstance(ws.cell(row=LAYOUT_HEADER_ROW, column=col).value, int)
    }

    for row in range(LAYOUT_FIRST_ROW, LAYOUT_FIRST_ROW + 8):
        row_letter = ws.cell(row=row, column=2).value
        if row_letter not in list("ABCDEFGH"):
            continue

        for col, plate_number in column_numbers.items():
            value = ws.cell(row=row, column=col).value
            if value is None:
                continue
            layout[f"{row_letter}{plate_number}"] = str(value)

    return layout


def read_340_data(ws):
    """Findet automatisch die 340-nm-Tabelle; 'T° 340' wird ignoriert."""
    header_row = None
    for row in ws.iter_rows():
        values = [cell.value for cell in row]
        if "Time" in values:
            well_count = sum(
                1 for v in values
                if isinstance(v, str) and len(v) >= 2
                and v[0] in "ABCDEFGH" and v[1:].isdigit()
            )
            if well_count >= 5:
                header_row = row[0].row
                break

    if header_row is None:
        raise ValueError("Keine passende 340-Messdatentabelle gefunden.")

    headers = {
        cell.value: cell.column
        for cell in ws[header_row]
        if cell.value is not None
    }

    selected_columns = {"Time": headers["Time"]}
    for header, column in headers.items():
        if (isinstance(header, str) and len(header) >= 2
                and header[0] in "ABCDEFGH" and header[1:].isdigit()):
            selected_columns[header] = column

    data = {key: [] for key in selected_columns}
    for row in ws.iter_rows(min_row=header_row + 1):
        time = row[selected_columns["Time"] - 1].value
        if time is None:
            break
        for well, column in selected_columns.items():
            data[well].append(row[column - 1].value)

    return data


def create_measurement_data(measurement_name, plate_rows, layout, raw_data):
    result = {"Time": raw_data["Time"]}

    for row_letter in plate_rows:
        wells = sorted(
            (w for w in layout if w.startswith(row_letter)),
            key=lambda w: int(w[1:])
        )
        for well in wells:
            if well not in raw_data:
                continue
            result[layout[well]] = raw_data[well]

    return result


def write_measurement_sheet(wb, sheet_name, measurement_data):
    ws = wb.create_sheet(title=sheet_name)
    headers = list(measurement_data.keys())

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    number_of_rows = len(measurement_data["Time"])
    for row in range(number_of_rows):
        excel_row = row + 2
        for col, header in enumerate(headers, start=1):
            ws.cell(row=excel_row, column=col, value=measurement_data[header][row])

    for row in range(2, number_of_rows + 2):
        ws.cell(row=row, column=1).number_format = "hh:mm:ss"

    for column in ws.columns:
        max_length = max(
            (len(str(c.value)) for c in column if c.value is not None),
            default=0
        )
        ws.column_dimensions[column[0].column_letter].width = min(max_length + 2, 20)

    ws.freeze_panes = "A2"


def main():
    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(f"Datei nicht gefunden:\n{INPUT_FILE}")

    input_wb = load_workbook(INPUT_FILE)

    if SHEET_NAME not in input_wb.sheetnames:
        raise ValueError(
            f"Tabellenblatt '{SHEET_NAME}' nicht gefunden.\n"
            f"Vorhanden: {input_wb.sheetnames}"
        )

    ws = input_wb[SHEET_NAME]

    layout = read_plate_layout(ws)
    print(f"{len(layout)} Well-Zuordnungen")

    raw_data = read_340_data(ws)
    print(f"{len(raw_data)} Daten-Spalten")

    output_wb = Workbook()
    output_wb.remove(output_wb.active)

    for measurement_name, plate_rows in MEASUREMENTS.items():
        measurement_data = create_measurement_data(
            measurement_name, plate_rows, layout, raw_data
        )
        write_measurement_sheet(output_wb, measurement_name, measurement_data)
        print(f"{measurement_name} erstellt: {' + '.join(plate_rows)}")

    output_wb.save(OUTPUT_FILE)
    print(f"\nNeue Datei:\n{OUTPUT_FILE}")


if __name__ == "__main__":
    main()
