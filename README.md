# Excel to GeoJSON Converter

Этот скрипт извлекает координаты из Excel-файлов (.xlsx) некоторого документного типа и преобразует их в формат GeoJSON. Результирующие файлы сохраняются в указанном каталоге, а любые ошибки или проблемы логируются в файл `error.log`.

## Установка зависимостей

Для установки зависимостей можно использовать следующую команду:

```bash
pip install -r requirements.txt
```
## Использование

Скрипт принимает аргументы командной строки для гибкой настройки обработки файлов и генерации GeoJSON. Ниже приведены все доступные аргументы и примеры их использования.

### Доступные аргументы

- **input** (позиционный)  
  Путь к каталогу с Excel-файлами или к конкретному файлу.  
  **По умолчанию:** текущая рабочая директория.

- **-o, --output**  
  Каталог для сохранения сгенерированных GeoJSON-файлов.  
  **По умолчанию:** `result`

- **--no-cycle-check**  
  Отключает проверку цикличности нумерации координат полигона.  
  **По умолчанию:** проверка включена.

- **--start-cell**  
  Задает адрес ячейки (например, `"F19"`), с которой начинается парсинг координат полигона. Если указан, автоматический поиск первой координаты не выполняется.  
  **По умолчанию:** не задан.

- **--enable-anchor-geojson**  
  Включает генерацию GeoJSON-файла для координат привязки.  
  **По умолчанию:** GeoJSON для привязки не создается.

### Примеры запуска

- **Запуск с использованием настроек по умолчанию:**  
  Обработка всех Excel-файлов в текущем каталоге и сохранение GeoJSON в папку `result` этого каталога.
  ```bash
  py convert_xl2gj.py
- **Запуск с использованием всех аргументов:**  
  ```bash
  python convert_xl2gj.py example.xlsx -o "geojsonfolder/automatic" --enable-anchor-geojson --start-cell F5 --no-cycle-check
