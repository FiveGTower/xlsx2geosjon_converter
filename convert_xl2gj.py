import os
import openpyxl
import re
import json
import argparse
import logging
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

logging.basicConfig(filename="error.log", level=logging.ERROR, format="%(asctime)s - %(message)s")

def get_excel_files(path: str) -> list[str]:
    """
    Получает список Excel-файлов (.xlsx) по заданному пути.

    Входные параметры:
      path (str): Путь к файлу или каталогу.

    Выходные данные:
      list[str]: Список путей к найденным Excel-файлам (фильтрация временных файлов).
    """
    if os.path.isfile(path):
        return [path] if path.endswith('.xlsx') and not os.path.basename(path).startswith('~$') else []
    elif os.path.isdir(path):
        return [os.path.join(path, f) for f in os.listdir(path) if f.endswith('.xlsx') and not f.startswith('~$')]
    return []

def is_valid_coordinate(coord_str: str) -> bool:
    """
    Проверяет, соответствует ли строка заданному формату координат.

    Входные параметры:
      coord_str (str): Строка, содержащая координаты.

    Выходные данные:
      bool: True, если строка соответствует формату координат, иначе False.
    """
    pattern = r'^(?:[NS]\d+\.\d+\s[EW]\d+\.\d+|[EW]\d+\.\d+\s[NS]\d+\.\d+|\d+\.\d+\s\d+\.\d+)$'
    if not re.match(pattern, coord_str):
        return False
    try:
        parts = coord_str.split()
        values = [float(part[1:]) if part[0] in "NSEW" else float(part) for part in parts]
        if len(values) != 2:
            return False
        # Если первая часть начинается с E или W, то порядок: (долгота, широта), иначе наоборот.
        lon, lat = values if parts[0][0] in "EW" else reversed(values)
        return 0 <= lon <= 180 and 0 <= lat <= 90
    except ValueError:
        return False

def parse_coordinates(coord_str: str) -> tuple[float, float]:
    """
    Преобразует строку с координатами в кортеж чисел (широта, долгота).

    Входные параметры:
      coord_str (str): Строка с координатами (например, "N55.5 E37.5" или "55.5 37.5").

    Выходные данные:
      tuple[float, float]: Кортеж, содержащий (широта, долгота).
    """
    parts = coord_str.split()
    lat = None
    lon = None
    raw_values = []
    for part in parts:
        if part.startswith("N"):
            try:
                lat = float(part[1:])
            except ValueError:
                raise ValueError(f"Неверное значение широты: {part}")
        elif part.startswith("E"):
            try:
                lon = float(part[1:])
            except ValueError:
                raise ValueError(f"Неверное значение долготы: {part}")
        else:
            try:
                raw_values.append(float(part))
            except ValueError:
                raise ValueError(f"Невозможно преобразовать '{part}' в число")
    if lat is None:
        if raw_values:
            lat = raw_values.pop(0)
        else:
            raise ValueError("Не указана широта")
    if lon is None:
        if raw_values:
            lon = raw_values.pop(0)
        else:
            raise ValueError("Не указана долгота")
    return (lat, lon)

def find_first_coordinate(sheet, keyword: str) -> tuple[int, int]:
    """
    Ищет первую ячейку с координатой, рядом с которой находится заголовок, начинающийся с заданного ключевого слова.

    Входные параметры:
      sheet: Объект рабочего листа openpyxl.
      keyword (str): Ключевое слово для поиска (например, "Номер" или "Привязка").

    Выходные данные:
      tuple[int, int]: Кортеж (номер строки, номер столбца) найденной ячейки или (-1, -1), если не найдено.
    """
    for row_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        for col_idx, cell in enumerate(row, start=1):
            if cell and isinstance(cell, str) and is_valid_coordinate(cell):
                prev_cell = sheet.cell(row=max(1, row_idx - 1), column=1).value
                if isinstance(prev_cell, str) and prev_cell.startswith(keyword):
                    return row_idx, col_idx
    return -1, -1

def read_excel_coordinates(file_path: str, start_cell: str = None, cycle_check: bool = True) -> tuple[list[tuple[float, float]] | None, list[tuple[float, float]]]:
    """
    Извлекает координаты полигона и привязки из Excel-файла.

    Входные параметры:
      file_path (str): Путь к Excel-файлу.
      start_cell (str, optional): Адрес ячейки начала парсинга координат полигона (например, "F19").
                                  Если задан, поиск через find_first_coordinate не выполняется.
      cycle_check (bool): Флаг, указывающий, нужно ли выполнять проверку цикличности нумерации.
                          Если False, нумерация не читается и не проверяется.

    Выходные данные:
      tuple:
        - list[tuple[float, float]]: Список координат полигона или None при ошибке.
        - list[tuple[float, float]]: Список координат привязки (MultiPoint) или пустой список.
    """
    polygon_coordinates = []
    anchor_coordinates = []
    numbering_list = []
    try:
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        sheet = workbook.active

        if start_cell:
            col_letter, row_number = coordinate_from_string(start_cell)
            poly_row, poly_col = row_number, column_index_from_string(col_letter)
        else:
            poly_row, poly_col = find_first_coordinate(sheet, "Номер")
        
        if poly_row == -1:
            logging.error(f"Не найдены координаты полигона в {file_path}")
            return None, None

        current_row = poly_row
        while True:
            coord_cell = sheet.cell(row=current_row, column=poly_col).value
            if coord_cell and isinstance(coord_cell, str) and is_valid_coordinate(coord_cell):
                try:
                    coord = parse_coordinates(coord_cell)
                except Exception as e:
                    logging.error(f"Ошибка парсинга координат в {file_path} строка {current_row}: {e}")
                    return None, None
                polygon_coordinates.append(coord)
                if cycle_check:
                    # Чтение нумерации: значения из ячеек, находящихся на 4 и 3 позиции левее
                    num_cell1 = sheet.cell(row=current_row, column=poly_col - 4).value
                    num_cell2 = sheet.cell(row=current_row, column=poly_col - 3).value
                    try:
                        num1 = int(num_cell1) if num_cell1 is not None else None
                        num2 = int(num_cell2) if num_cell2 is not None else None
                    except ValueError:
                        logging.error(f"Невозможно преобразовать нумерацию в {file_path} строка {current_row}")
                        return None, None
                    if num1 is None or num2 is None:
                        logging.error(f"Отсутствует нумерация в {file_path} строка {current_row}")
                        return None, None
                    numbering_list.append((num1, num2))
                current_row += 1
            else:
                break

        # Если координаты полигона отсутствуют, логируем ошибку и возвращаем None.
        if not polygon_coordinates:
            logging.error(f"В файле {file_path} не найдены координаты полигона.")
            return None, None

        if cycle_check:
            # Проверка последовательности нумерации
            for i in range(len(numbering_list) - 1):
                if numbering_list[i][1] != numbering_list[i+1][0]:
                    logging.error(f"Нарушена последовательность нумерации в {file_path} на строке {poly_row + i}")
                    return None, None
            if numbering_list and (numbering_list[-1][1] != numbering_list[0][0]):
                logging.error(f"Нарушена циклическая нумерация в {file_path}: последний элемент {numbering_list[-1][1]} не соответствует первому {numbering_list[0][0]}")
                return None, None

        # Извлечение координат привязки
        anchor_row, anchor_col = find_first_coordinate(sheet, "Привязка")
        if anchor_row != -1:
            for row in sheet.iter_rows(min_row=anchor_row, min_col=anchor_col, values_only=True):
                if row and row[0] and isinstance(row[0], str) and is_valid_coordinate(str(row[0])):
                    try:
                        anchor_coordinates.append(parse_coordinates(str(row[0])))
                    except Exception as e:
                        logging.error(f"Ошибка парсинга координат привязки в {file_path}: {e}")
                        break
                else:
                    break

    except Exception as e:
        logging.error(f"Ошибка чтения {file_path}: {e}")
        return None, None

    return polygon_coordinates, anchor_coordinates

def generate_geojson(file_path: str, polygon_coords: list[tuple[float, float]], anchor_coords: list[tuple[float, float]], output_dir: str, create_anchor: bool = True):
    """
    Генерирует GeoJSON-файл для полигона и, при необходимости, для привязки.

    Входные параметры:
      file_path (str): Путь к исходному Excel-файлу (используется для формирования имени).
      polygon_coords (list[tuple[float, float]]): Список координат полигона.
      anchor_coords (list[tuple[float, float]]): Список координат привязки.
      output_dir (str): Каталог для сохранения сгенерированного GeoJSON.
      create_anchor (bool): Флаг создания GeoJSON для привязки. Если False – файл не создаётся.

    Выходные данные:
      Файлы GeoJSON сохраняются в указанном каталоге.
    """
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    os.makedirs(output_dir, exist_ok=True)
    if polygon_coords:
        geojson = {
            "type": "FeatureCollection",
            "name": f"{file_name}.xlsx",
            "features": [
                {
                    "type": "Feature",
                    "properties": {"name": "0", "buffer": 0},
                    "geometry": {"type": "Polygon", "coordinates": [polygon_coords + [polygon_coords[0]]]}
                }
            ],
            "crs": {"type": "name", "properties": {"name": "urn:ogc:def:crs:OGC:1.3:CRS84"}}
        }
        output_path = os.path.join(output_dir, f"{file_name}.geojson")
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(geojson, f, ensure_ascii=False, indent=4)
        print(f"GeoJSON сохранен: {output_path}")
    if create_anchor and anchor_coords:
        anchor_output_path = os.path.join(output_dir, f"{file_name}_.geojson")
        anchor_geojson = {
            "type": "FeatureCollection",
            "name": f"{file_name}_.geojson",
            "features": [
                {
                    "type": "Feature",
                    "geometry": {"type": "MultiPoint", "coordinates": anchor_coords},
                    "properties": {"name": "0", "buffer": 0}
                }
            ],
            "crs": {"type": "name", "properties": {"name": "urn:ogc:def:crs:OGC:1.3:CRS84"}}
        }
        with open(anchor_output_path, "w", encoding="utf-8") as f:
            json.dump(anchor_geojson, f, ensure_ascii=False, indent=4)
        print(f"GeoJSON сохранен: {anchor_output_path}")

def main():
    parser = argparse.ArgumentParser(description="Конвертер Excel в GeoJSON")
    parser.add_argument("input", type=str, nargs="?", default=os.getcwd(),
                        help="Путь к каталогу с Excel-файлами или к конкретному файлу")
    parser.add_argument("-o", "--output", type=str, default="result",
                        help="Каталог для сохранения GeoJSON")
    parser.add_argument("--no-cycle-check", action="store_true",
                        help="Отключить проверку цикличности нумерации")
    parser.add_argument("--start-cell", type=str, default=None,
                        help="Адрес ячейки начала парсинга координат полигона (например, F19)")
    parser.add_argument("--no-anchor-geojson", action="store_true",
                        help="Отключить создание GeoJSON с координатами привязки")
    args = parser.parse_args()

    excel_files = get_excel_files(args.input)
    if not excel_files:
        print("Нет файлов .xlsx для обработки.")
        return

    for file in excel_files:
        print(f"Обрабатываем: {file}")
        polygon_coords, anchor_coords = read_excel_coordinates(
            file,
            start_cell=args.start_cell,
            cycle_check=not args.no_cycle_check
        )
        if polygon_coords is None:
            print(f"Ошибка в файле {file}. Пропускаем.")
            continue
        generate_geojson(
            file,
            polygon_coords,
            anchor_coords,
            args.output,
            create_anchor=not args.no_anchor_geojson
        )
    print("Обработка завершена!")

if __name__ == "__main__":
    main()
