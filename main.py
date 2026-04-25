"""
Аналитический инструмент для обработки экспериментальных данных и построения трендов.
Реализован метод наименьших квадратов для линейной и полиномиальной регрессии.
Входные данные: файл CSV/TXT с парами (x, y).
Вывод: график с аппроксимирующими кривыми, коэффициенты регрессии, R^2.
"""

import sys
import os
import numpy as np
import matplotlib.pyplot as plt


def load_data(filepath):
    """
    Загрузка данных из текстового файла.
    Поддерживаются форматы с разделителями: запятая, точка с запятой, пробелы/табуляции.
    Ожидается, что в файле ровно два столбца чисел (x, y).
    Строки, начинающиеся с '#', считаются комментариями.
    """
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"Файл '{filepath}' не найден.")

    # Пробуем разные разделители
    for delimiter in [',', ';', None]:  # None – автоопределение пробельных символов
        try:
            data = np.loadtxt(filepath, delimiter=delimiter, comments='#')
            if data.ndim == 2 and data.shape[1] == 2:
                break
        except ValueError:
            continue
    else:
        # Если автоматически не получилось, читаем построчно
        x_vals, y_vals = [], []
        with open(filepath, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                # Заменяем запятую на точку и разбиваем по пробелам/запятым
                parts = line.replace(',', ' ').split()
                if len(parts) >= 2:
                    try:
                        x_vals.append(float(parts[0]))
                        y_vals.append(float(parts[1]))
                    except ValueError:
                        continue
        if not x_vals:
            raise ValueError("Не удалось считать данные из файла.")
        return np.array(x_vals), np.array(y_vals)

    x = data[:, 0]
    y = data[:, 1]
    if len(x) == 0:
        raise ValueError("Файл содержит пустые данные.")
    return x, y


def poly_fit(x, y, degree):
    """
    Полиномиальная аппроксимация методом наименьших квадратов.
    Используется построение матрицы Вандермонда и решение нормальных уравнений
    через np.linalg.lstsq.
    Возвращает коэффициенты полинома в порядке возрастания степеней: [c0, c1, ..., cd].
    """
    # Построение матрицы A (столбцы: x^0, x^1, ..., x^degree)
    A = np.vander(x, degree + 1, increasing=True)
    # Решение МНК-задачи ||A*c - y||^2 -> min
    coeffs, residuals, rank, s = np.linalg.lstsq(A, y, rcond=None)
    return coeffs


def poly_val(coeffs, x):
    """
    Вычисление значений полинома, заданного коэффициентами coeffs
    (в порядке возрастания степеней) в точках x.
    """
    y = np.zeros_like(x, dtype=float)
    for i, c in enumerate(coeffs):
        y += c * (x ** i)
    return y


def r_squared(y_true, y_pred):
    """
    Вычисление коэффициента детерминации R^2.
    R^2 = 1 - SS_res / SS_tot, где
    SS_res = sum (y_i - y_pred_i)^2,
    SS_tot = sum (y_i - mean(y))^2.
    """
    ss_res = np.sum((y_true - y_pred) ** 2)
    ss_tot = np.sum((y_true - np.mean(y_true)) ** 2)
    return 1 - ss_res / ss_tot


def print_coefficients(name, coeffs, degree):
    """Красивый вывод коэффициентов полинома."""
    print(f"\n{name}:")
    if degree == 1:
        print(f"  y = {coeffs[0]:.6f} + {coeffs[1]:.6f} * x")
    else:
        terms = []
        for i, c in enumerate(coeffs):
            if i == 0:
                terms.append(f"{c:.6f}")
            else:
                sign = "+" if c >= 0 else "-"
                abs_c = abs(c)
                if i == 1:
                    terms.append(f"{sign} {abs_c:.6f} * x")
                else:
                    terms.append(f"{sign} {abs_c:.6f} * x^{i}")
        print("  y = " + " ".join(terms))


def main():
    # Путь к файлу данных: либо из аргументов командной строки, либо по умолчанию
    if len(sys.argv) > 1:
        filepath = sys.argv[1]
    else:
        filepath = "data.txt"  # Имя файла по умолчанию

    print(f"Загрузка данных из файла: {filepath}")
    try:
        x, y = load_data(filepath)
    except Exception as e:
        print(f"Ошибка загрузки: {e}")
        sys.exit(1)

    print(f"Загружено {len(x)} точек")

    # Проверка, что данные не вырождены
    if len(np.unique(x)) < 2:
        print("Ошибка: недостаточно уникальных значений x для регрессии.")
        sys.exit(1)

    # ---------- Линейная регрессия (степень 1) ----------
    coeffs_lin = poly_fit(x, y, 1)
    y_pred_lin = poly_val(coeffs_lin, x)
    r2_lin = r_squared(y, y_pred_lin)
    print_coefficients("Линейная регрессия", coeffs_lin, 1)
    print(f"  R^2 = {r2_lin:.6f}")

    # ---------- Полином 2-й степени ----------
    if len(x) >= 3:  # достаточно данных
        coeffs_poly2 = poly_fit(x, y, 2)
        y_pred_poly2 = poly_val(coeffs_poly2, x)
        r2_poly2 = r_squared(y, y_pred_poly2)
        print_coefficients("Полином 2-й степени", coeffs_poly2, 2)
        print(f"  R^2 = {r2_poly2:.6f}")
    else:
        coeffs_poly2 = None
        print("Недостаточно данных для полинома 2-й степени (нужно минимум 3 точки).")

    # ---------- Полином 3-й степени ----------
    if len(x) >= 4:
        coeffs_poly3 = poly_fit(x, y, 3)
        y_pred_poly3 = poly_val(coeffs_poly3, x)
        r2_poly3 = r_squared(y, y_pred_poly3)
        print_coefficients("Полином 3-й степени", coeffs_poly3, 3)
        print(f"  R^2 = {r2_poly3:.6f}")
    else:
        coeffs_poly3 = None
        print("Недостаточно данных для полинома 3-й степени (нужно минимум 4 точки).")

    # ---------- Визуализация ----------
    plt.figure(figsize=(10, 6))
    # Исходные данные
    plt.scatter(x, y, color='black', label='Экспериментальные данные', zorder=5)

    # Диапазон для построения кривых
    x_min, x_max = np.min(x), np.max(x)
    x_plot = np.linspace(x_min - 0.05*(x_max-x_min), x_max + 0.05*(x_max-x_min), 300)

    # Линейная регрессия
    y_plot_lin = poly_val(coeffs_lin, x_plot)
    plt.plot(x_plot, y_plot_lin, color='red', linewidth=2,
             label=f'Линейная (R²={r2_lin:.4f})')

    # Полином 2-й степени
    if coeffs_poly2 is not None:
        y_plot_poly2 = poly_val(coeffs_poly2, x_plot)
        plt.plot(x_plot, y_plot_poly2, color='green', linewidth=2,
                 label=f'Полином 2-й ст. (R²={r2_poly2:.4f})')

    # Полином 3-й степени
    if coeffs_poly3 is not None:
        y_plot_poly3 = poly_val(coeffs_poly3, x_plot)
        plt.plot(x_plot, y_plot_poly3, color='blue', linewidth=2,
                 label=f'Полином 3-й ст. (R²={r2_poly3:.4f})')

    plt.xlabel('x')
    plt.ylabel('y')
    plt.title('Аппроксимация данных методом наименьших квадратов')
    plt.legend()
    plt.grid(True, linestyle='--', alpha=0.6)
    plt.tight_layout()
    plt.show()


if __name__ == '__main__':
    main()