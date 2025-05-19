import pandas as pd
import os
import glob
import logging
import matplotlib.pyplot as plt
from io import BytesIO
from openpyxl.drawing.image import Image

def setup_logging(log_file):
    """Настраиваем логирование ошибок в файл."""
    logging.basicConfig(
        filename=log_file,
        level=logging.ERROR,
        format='%(asctime)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

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
    """Читаем Excel-файл (.xlsx или .xls)."""
    try:
        if file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path, engine='openpyxl')
        elif file_path.endswith('.xls'):
            df = pd.read_excel(file_path, engine='xlrd')
        else:
            raise ValueError("Неподдерживаемый формат файла. Используйте .xlsx или .xls")
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

    DEF = ['Много теории, но мало практики.'] + DEFICIENCIES

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


def save_results(all_disciplines, all_deficiencies, output_file):
    """Сохраняем агрегированные результаты и графики в Excel."""
    try:
        # Создаём графики
        disciplines_img, deficiencies_img = create_charts(all_disciplines, all_deficiencies)

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Сохраняем таблицы
            all_disciplines.to_excel(writer, sheet_name='Results', index=False, startrow=0)
            all_deficiencies.to_excel(writer, sheet_name='Results', index=False, startrow=len(all_disciplines) + 3)

            # Вставляем графики
            workbook = writer.book
            worksheet = writer.sheets['Results']

            # График дисциплин
            img1 = Image(disciplines_img)
            worksheet.add_image(img1, f'A{len(all_disciplines) + len(all_deficiencies) + 6}')

            # График недостатков
            img2 = Image(deficiencies_img)
            worksheet.add_image(img2, f'A{len(all_disciplines) + len(all_deficiencies) + 26}')

        print(f"Агрегированные результаты и графики сохранены в {output_file}")
    except Exception as e:
        print(f"Ошибка при сохранении файла {output_file}: {e}")

def process_multiple_files(input_dir, output_file, log_file):
    """Обрабатываем все .xlsx и .xls файлы и агрегируем результаты."""
    setup_logging(log_file)

    # Находим файлы
    file_paths = glob.glob(os.path.join(input_dir, "*.xlsx")) + glob.glob(os.path.join(input_dir, "*.xls"))

    if not file_paths:
        error_msg = f"В директории {input_dir} не найдено .xlsx или .xls файлов"
        print(error_msg)
        logging.error(error_msg)
        return

    disciplines_list = []
    deficiencies_list = []

    for file_path in file_paths:
        print(f"\nОбработка файла: {file_path}")
        try:
            df = read_excel_file(file_path)
            if df is None:
                logging.error(f"Не удалось прочитать файл {file_path}")
                continue

            if not check_table_structure(df):
                logging.error(f"Файл {file_path} не прошёл проверку структуры")
                continue

            result_disciplines = calculate_sumproduct(df)
            if result_disciplines is None:
                logging.error(f"Ошибка при вычислении СУММПРОИЗВ для {file_path}")
                continue

            result_deficiencies = calculate_deficiency_totals(df)
            if result_deficiencies is None:
                logging.error(f"Ошибка при вычислении недостатков для {file_path}")
                continue

            disciplines_list.append(result_disciplines)
            deficiencies_list.append(result_deficiencies)

        except Exception as e:
            error_msg = f"Необработанная ошибка в файле {file_path}: {str(e)}"
            print(error_msg)
            logging.error(error_msg)

    if not disciplines_list or not deficiencies_list:
        error_msg = "Не удалось обработать ни один файл"
        print(error_msg)
        logging.error(error_msg)
        return

    # Агрегируем дисциплины
    all_disciplines = pd.concat(disciplines_list, ignore_index=True)
    all_disciplines = all_disciplines.groupby('Дисциплина', sort=False).agg({
        'СУММПРОИЗВ': ['mean', 'std']
    }).reset_index()
    all_disciplines.columns = ['Дисциплина', 'Среднее СУММПРОИЗВ', 'Станд. отклонение СУММПРОИЗВ']
    # Сортируем по порядку DISCIPLINES
    all_disciplines['order'] = all_disciplines['Дисциплина'].map({d: i for i, d in enumerate(DISCIPLINES)})
    all_disciplines = all_disciplines.sort_values('order').drop('order', axis=1).reset_index(drop=True)
    all_disciplines['Ранг'] = all_disciplines['Среднее СУММПРОИЗВ'].rank(ascending=False, method='min').astype(int)

    # Агрегируем недостатки
    all_deficiencies = pd.concat(deficiencies_list, ignore_index=True)
    all_deficiencies = all_deficiencies.groupby('Недостаток', sort=False).agg({
        'Веса': 'mean',
        'Сумма оценок': 'mean',
        'Взвешенная сумма': ['mean', 'std']
    }).reset_index()
    all_deficiencies.columns = ['Недостаток', 'Среднее Веса', 'Средняя Сумма оценок', 'Средняя Взвешенная сумма', 'Станд. отклонение Взвешенной суммы']
    # Сортируем по порядку DEFICIENCIES
    DEF = ['Много теории, но мало практики.'] + DEFICIENCIES
    all_deficiencies['order'] = all_deficiencies['Недостаток'].map({d: i for i, d in enumerate(DEF)})
    all_deficiencies = all_deficiencies.sort_values('order').drop('order', axis=1).reset_index(drop=True)
    all_deficiencies['Ранг'] = all_deficiencies['Средняя Взвешенная сумма'].rank(ascending=False, method='min').astype(int)

    # Сохраняем результаты
    save_results(all_disciplines, all_deficiencies, output_file)
    print(f"\nОбработка завершена. Лог ошибок сохранён в {log_file}")


def create_charts(all_disciplines, all_deficiencies):
    """Создаём столбчатые диаграммы для дисциплин и недостатков."""
    # График для дисциплин
    plt.figure(figsize=(12, 8))  # Увеличиваем размер
    bars = plt.bar(all_disciplines['Дисциплина'], all_disciplines['Средние баллы по недостаткам для дисциплин'], color='#1f77b4')
    plt.title('Средние баллы по дисциплинам')
    plt.xlabel('Дисциплина')
    plt.ylabel('Средние баллы')
    plt.xticks(rotation=45, ha='right', fontsize=8)  # Уменьшаем шрифт
    plt.subplots_adjust(bottom=0.3)  # Увеличиваем нижнюю границу
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, yval, f'{yval:.1f}', va='bottom', fontsize=8)
    disciplines_img = BytesIO()
    plt.savefig(disciplines_img, format='png', bbox_inches='tight')
    plt.close()
    disciplines_img.seek(0)

    # График для недостатков
    plt.figure(figsize=(12, 8))  # Увеличиваем размер
    bars = plt.bar(all_deficiencies['Недостаток'], all_deficiencies['Средние баллы по наличию недостатков'], color='#ff7f0e')
    plt.title('Средняя Взвешенная сумма по недостаткам')
    plt.xlabel('Недостаток')
    plt.ylabel('Средняя Взвешенная сумма')
    plt.xticks(rotation=45, ha='right', fontsize=8)  # Уменьшаем шрифт
    plt.subplots_adjust(bottom=0.3)  # Увеличиваем нижнюю границу
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, yval, f'{yval:.1f}', va='bottom', fontsize=8)
    deficiencies_img = BytesIO()
    plt.savefig(deficiencies_img, format='png', bbox_inches='tight')
    plt.close()
    deficiencies_img.seek(0)

    return disciplines_img, deficiencies_img

def main():
    # Пути
    input_dir = "D:/PythonProject/AppForDevAnalyse"
    output_file = "D:/PythonProject/AppForDevAnalyse/output_results.xlsx"
    log_file = "D:/PythonProject/AppForDevAnalyse/errors.log"

    # Обрабатываем все файлы
    process_multiple_files(input_dir, output_file, log_file)

if __name__ == "__main__":
    main()