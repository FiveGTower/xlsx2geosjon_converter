import unittest
from convert_xl2gj import parse_coordinates, is_valid_coordinate

class TestCoordinateParser(unittest.TestCase):
    """Тесты для функций парсинга и валидации координат"""

    ## Тесты для is_valid_coordinate()

    def test_valid_coordinates(self):
        """Проверка корректных координат"""
        self.assertTrue(is_valid_coordinate("N64.062788 E67.503584"))
        self.assertTrue(is_valid_coordinate("E67.503584 N64.062788"))
        self.assertTrue(is_valid_coordinate("64.062788 67.503584"))

    def test_invalid_format(self):
        """Проверка некорректных форматов строк"""
        self.assertFalse(is_valid_coordinate("N64.062788 X67.503584"))  # Некорректный префикс
        self.assertFalse(is_valid_coordinate("64.062788"))  # Только одно число
        self.assertFalse(is_valid_coordinate(""))  # Пустая строка
        self.assertFalse(is_valid_coordinate("N64.062788"))  # Только широта
        self.assertFalse(is_valid_coordinate("64,062788 67,503584"))  # Запятая вместо точки
        self.assertFalse(is_valid_coordinate("N200.000000 E100.000000"))  # Выход за допустимый диапазон
        self.assertFalse(is_valid_coordinate("S64.062788 W67.503584"))  # Другие полушария не обрабатываем

    def test_edge_cases(self):
        """Граничные значения координат"""
        self.assertTrue(is_valid_coordinate("N90.000000 E180.000000"))  # Максимальные допустимые
        self.assertTrue(is_valid_coordinate("0.000000 0.000000"))  # Ноль

    ## Тесты для parse_coordinates()

    def test_parse_valid_coordinates(self):
        """Проверка успешного парсинга корректных данных"""
        self.assertEqual(parse_coordinates("N64.062788 E67.503584"), (64.062788, 67.503584))
        self.assertEqual(parse_coordinates("E67.503584 N64.062788"), (64.062788, 67.503584))
        self.assertEqual(parse_coordinates("64.062788 67.503584"), (64.062788, 67.503584))

    def test_parse_invalid_format(self):
        """Проверка обработки некорректных данных"""
        with self.assertRaises(ValueError):
            parse_coordinates("N64.062788 X67.503584")  # Некорректный префикс
        with self.assertRaises(ValueError):
            parse_coordinates("64.062788")  # Только одно число
        with self.assertRaises(ValueError):
            parse_coordinates("")  # Пустая строка
        with self.assertRaises(ValueError):
            parse_coordinates("N64.062788")  # Только широта
        with self.assertRaises(ValueError):
            parse_coordinates("64,062788 67,503584")  # Запятая вместо точки
        with self.assertRaises(ValueError):
            parse_coordinates("S64.062788 W67.503584")  # Другие полушария не обрабатываем

    def test_parse_edge_cases(self):
        """Граничные значения координат"""
        self.assertEqual(parse_coordinates("N90.000000 E180.000000"), (90.0, 180.0))  # Максимальные
        self.assertEqual(parse_coordinates("0.000000 0.000000"), (0.0, 0.0))  # Ноль

if __name__ == "__main__":
    unittest.main()
