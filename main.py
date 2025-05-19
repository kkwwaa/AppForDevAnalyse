import pandas as pd

# Список недостатков (только 6 элементов, без 'NB! Все числа - положительные!')
DEFICIENCIES = [
    'Низкие требования (низкие требования к оцениванию, низкая сложность заданий, отсутствие дедлайнов, много пересдач)',
    'Нет командных работ (в группах)',
    'Давать неактуальные знания, изучать устаревший материал, использовать устаревшее ПО',
    'Преподаватель не имеет опыта по своему предмету, читает лекции монотонно и скучно',
    'Не идти на контакт со студентами, отказывать в объяснении, не отвечать на вопросы',
    'Не поощрять творческий подход, инициативность и  самостоятельность студентов'
]

# Список дисциплин
DISCIPLINES = [
    "Безопасность жизнедеятельности",
    "Математический анализ",
    "Алгебра",
    "Программирование",
    "Дискретная математика",
    "Профориентационный семинар",
    "Проектный семинар",
    "Практикум по основам разработки технической документации",
    "Теоретические основы информатики",
    "История России",
    "Основы российской государственности",
    "Английский язык",
    "Правовая грамотность"
]

def read_excel_file(file_path):
    """Читаем Excel-файл."""
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        print(f"Файл {file_path} успешно прочитан!")
        return df
    except Exception as e:
        print(f"Ошибка при чтении файла {file_path}: {e}")
        return None


def check_table_structure(df):
    """Проверяем, что таблица имеет правильную структуру."""
    # Ожидаемое количество столбцов: 9
    if len(df.columns) != 9:
        print(f"Ошибка: ожидается 9 столбцов, найдено {len(df.columns)}: {list(df.columns)}")
        return False

    # Проверяем названия столбцов
    expected_columns = ['Unnamed: 0', 'Unnamed: 1', 'Недостаток', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5',
                        'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']
    if list(df.columns) != expected_columns:
        print(f"Ошибка: ожидались столбцы {expected_columns}, найдены {list(df.columns)}")
        return False

    # Проверяем количество строк: 16 (заголовки, важность, пустая, 13 дисциплин)
    if len(df) != 16:
        print(f"Ошибка: ожидается 16 строк, найдено {len(df)}")
        return False

    # Проверяем недостатки в первой строке (столбцы Unnamed: 3–Unnamed: 8)
    deficiency_columns = ['Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']
    deficiencies_row = df.iloc[0][deficiency_columns].tolist()
    # Очищаем строки от лишних пробелов и невидимых символов
    deficiencies_row = [str(x).strip() if pd.notna(x) else x for x in deficiencies_row]
    if deficiencies_row != DEFICIENCIES:
        print(f"Ошибка: недостатки в первой строке не совпадают.")
        print(f"Ожидались: {DEFICIENCIES}")
        print(f"Найдены: {deficiencies_row}")
        for i, (expected, found) in enumerate(zip(DEFICIENCIES, deficiencies_row)):
            if expected != found:
                print(f"Различие в элементе {i}:")
                print(f"  Ожидалось: {expected}")
                print(f"  Найдено: {found}")
                print(f"  Ожидалось (коды символов): {[ord(c) for c in expected]}")
                print(f"  Найдено (коды символов): {[ord(c) for c in found if c != ' ']}")
        return False

    # Проверяем, что Unnamed: 1 в первой строке содержит 'NB! Все числа - положительные!'
    nb_text = str(df.iloc[0]['Unnamed: 1']).strip()
    expected_nb = 'NB! Все числа - положительные!'
    if nb_text != expected_nb:
        print(f"Ошибка: в Unnamed: 1 (первая строка) ожидалось '{expected_nb}', найдено: '{nb_text}'")
        return False

    # Проверяем важность во второй строке (столбцы Unnamed: 3–Unnamed: 8)
    importance = df.iloc[1][deficiency_columns]
    # Проверяем, что все значения можно преобразовать в числа
    try:
        importance = pd.to_numeric(importance, errors='coerce')
    except Exception as e:
        print(f"Ошибка: значения важности (вторая строка) не могут быть преобразованы в числа: {importance.tolist()}")
        print(f"Подробности: {e}")
        return False

    # Проверяем пропуски после преобразования
    if importance.isna().any():
        print(f"Ошибка: в значениях важности есть пропуски или некорректные данные: {importance.tolist()}")
        return False

    # Проверяем диапазон 0–10
    if not importance.between(0, 10).all():
        print(f"Ошибка: значения важности (вторая строка) должны быть от 0 до 10, найдены: {importance.tolist()}")
        return False

    # Проверяем номера дисциплин в первом столбце (index 3–15, должны быть 1.0–13.0)
    discipline_numbers = df.iloc[3:, 0].tolist()
    expected_numbers = [float(i) for i in range(1, 14)]
    if discipline_numbers != expected_numbers:
        print(
            f"Ошибка: номера дисциплин в первом столбце не совпадают. Ожидались: {expected_numbers}, найдены: {discipline_numbers}")
        return False

    # Проверяем оценки (index 3–15, столбцы Unnamed: 3–Unnamed: 8)
    scores = df.iloc[3:][deficiency_columns]
    # Преобразуем оценки в числа
    try:
        scores = scores.apply(pd.to_numeric, errors='coerce')
    except Exception as e:
        print(f"Ошибка: оценки (строки 4–16) не могут быть преобразованы в числа: {e}")
        return False

    if not ((scores >= 0) & (scores <= 10)).all().all():
        print(
            f"Ошибка: оценки (строки 4–16, столбцы {deficiency_columns}) должны быть от 0 до 10, найдены некорректные значения")
        return False

    return True


def calculate_sumproduct(df):
    """Вычисляем СУММПРОИЗВ для каждой дисциплины."""
    # Извлекаем важность (вторая строка, столбцы Unnamed: 3–Unnamed: 8)
    deficiency_columns = ['Недостаток','Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']
    importance = pd.to_numeric(df.iloc[1][deficiency_columns], errors='coerce').values
    # Извлекаем оценки (строки 3–15, те же столбцы)
    scores = df.iloc[3:][deficiency_columns].apply(pd.to_numeric, errors='coerce')

    # Вычисляем СУММПРОИЗВ для каждой дисциплины
    discipline_totals = []
    for i in range(len(scores)):
        total = sum(importance * scores.iloc[i])
        discipline_totals.append(total)

    # Создаем новую таблицу для результатов
    result = pd.DataFrame({
        'Дисциплина': DISCIPLINES,
        'СУММПРОИЗВ': discipline_totals
    })

    # Добавляем столбец с рангами (наибольший СУММПРОИЗВ = ранг 1)
    result['Ранг'] = result['СУММПРОИЗВ'].rank(ascending=False, method='min').astype(int)

    return result

def calculate_deficiency_totals(df):
    """Вычисляем взвешенные суммы и ранги для каждого недостатка."""
    # Извлекаем важность (вторая строка, столбцы Unnamed: 3–Unnamed: 8)
    deficiency_columns = ['Недостаток', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']
    importance = pd.to_numeric(df.iloc[1][deficiency_columns], errors='coerce').values
    # Извлекаем оценки (строки 3–15, те же столбцы)
    scores = df.iloc[3:][deficiency_columns].apply(pd.to_numeric, errors='coerce')

    # Проверяем, нет ли пропусков в importance или scores
    if pd.isna(importance).any():
        print(f"Ошибка: в значениях важности есть пропуски: {importance.tolist()}")
        return None
    if scores.isna().any().any():
        print(f"Ошибка: в оценках есть пропуски: {scores.isna().sum().to_dict()}")
        return None

    # Вычисляем сумму оценок по каждому недостатку (по столбцам)
    deficiency_sums = scores.sum(axis=0).values
    # Умножаем суммы на веса недостатков
    weighted_totals = deficiency_sums * importance

    DEF = DEFICIENCIES
    DEF.append('Много теории, но мало практики.')

    # Создаем DataFrame с недостатками, суммами, взвешенными суммами и рангами
    result = pd.DataFrame({
        'Недостаток': DEF,
        'Веса': importance,
        'Сумма оценок': deficiency_sums,
        'Взвешенная сумма': weighted_totals
    })

    # Добавляем столбец с рангами (наибольшая взвешенная сумма = ранг 1)
    result['Ранг'] = result['Взвешенная сумма'].rank(ascending=False, method='min').astype(int)

    return result

def save_results(disciplines_df, deficiencies_df, output_file):
    """Сохраняем результаты по дисциплинам и недостаткам в один Excel-файл на одном листе."""
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Сохраняем таблицу дисциплин в начало листа
            disciplines_df.to_excel(writer, sheet_name='Results', index=False, startrow=0)
            # Сохраняем таблицу недостатков ниже с отступом в 2 строки
            deficiencies_df.to_excel(writer, sheet_name='Results', index=False, startrow=len(disciplines_df) + 3)
        print(f"Результаты сохранены в {output_file}")
    except Exception as e:
        print(f"Ошибка при сохранении файла {output_file}: {e}")


def main():
    # Путь к файлу
    input_file = "1.xlsx"
    output_file = "output_results.xlsx"

    # Читаем файл
    df = read_excel_file(input_file)
    if df is None:
        return

    # Проверяем структуру
    if not check_table_structure(df):
        return

    # Вычисляем СУММПРОИЗВ для дисциплин
    result_disciplines = calculate_sumproduct(df)
    if result_disciplines is None:
        return

    # Вычисляем взвешенные суммы для недостатков
    result_deficiencies = calculate_deficiency_totals(df)
    if result_deficiencies is None:
        return

    # Сохраняем результаты в один файл
    save_results(result_disciplines, result_deficiencies, output_file)


if __name__ == "__main__":
    main()