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
    """Проверяем, что таблица имеет правильную структуру для обработки (минимум 9 столбцов, 16 строк)."""
    if len(df.columns) < 9:
        print(f"Ошибка: ожидается минимум 9 столбцов, найдено {len(df.columns)}: {list(df.columns)}")
        return False

    if len(df) < 16:
        print(f"Ошибка: ожидается минимум 16 строк, найдено {len(df)}")
        return False

    expected_columns = ['Unnamed: 0', 'Unnamed: 1', 'Недостаток', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5',
                        'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']
    if list(df.columns)[:9] != expected_columns:
        print(f"Ошибка: ожидались столбцы {expected_columns}, найдены {list(df.columns)[:9]}")
        return False

    deficiency_columns = ['Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']
    deficiencies_row = df.iloc[0][deficiency_columns].tolist()
    deficiencies_row = [str(x).strip() if pd.notna(x) else x for x in deficiencies_row]
    if deficiencies_row[:6] != DEFICIENCIES:
        print(f"Ошибка: недостатки в первой строке не совпадают.")
        print(f"Ожидались: {DEFICIENCIES}")
        print(f"Найдены: {deficiencies_row}")
        return False

    nb_text = str(df.iloc[0]['Unnamed: 1']).strip()
    expected_nb = 'NB! Все числа - положительные!'
    if nb_text != expected_nb:
        print(f"Ошибка: в Unnamed: 1 (первая строка) ожидалось '{expected_nb}', найдено: '{nb_text}'")
        return False

    importance = df.iloc[1][deficiency_columns]
    try:
        importance = pd.to_numeric(importance, errors='coerce')
    except Exception as e:
        print(f"Ошибка: значения важности (вторая строка) не могут быть преобразованы в числа: {importance.tolist()}")
        return False

    if importance.isna().any():
        print(f"Ошибка: в значениях важности есть пропуски или некорректные данные: {importance.tolist()}")
        return False

    if not importance.between(0, 10).all():
        print(f"Ошибка: значения важности (вторая строка) должны быть от 0 до 10, найдены: {importance.tolist()}")
        return False

    discipline_numbers = df.iloc[3:16, 0].tolist()
    expected_numbers = [float(i) for i in range(1, 14)]
    if discipline_numbers != expected_numbers:
        print(f"Ошибка: номера дисциплин в первом столбце не совпадают. Ожидались: {expected_numbers}, найдены: {discipline_numbers}")
        return False

    scores = df.iloc[3:16][deficiency_columns]
    try:
        scores = scores.apply(pd.to_numeric, errors='coerce')
    except Exception as e:
        print(f"Ошибка: оценки (строки 4–16) не могут быть преобразованы в числа: {e}")
        return False

    if scores.isna().any().any():
        print(f"Ошибка: в оценках есть пропуски: {scores.isna().sum().to_dict()}")
        return False

    if not ((scores >= 0) & (scores <= 10)).all().all():
        print(f"Ошибка: оценки (строки 4–16, столбцы {deficiency_columns}) должны быть от 0 до 10")
        return False

    return True

def calculate_sumproduct(df):
    """Вычисляем СУММПРОИЗВ для каждой дисциплины."""
    deficiency_columns = ['Недостаток','Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']
    importance = pd.to_numeric(df.iloc[1][deficiency_columns], errors='coerce').values
    scores = df.iloc[3:16][deficiency_columns].apply(pd.to_numeric, errors='coerce')

    discipline_totals = []
    for i in range(len(scores)):
        total = sum(importance * scores.iloc[i])
        discipline_totals.append(total)

    result = pd.DataFrame({
        'Дисциплина': DISCIPLINES,
        'СУММПРОИЗВ': discipline_totals
    })

    result['Ранг'] = result['СУММПРОИЗВ'].rank(ascending=False, method='min').astype(int)
    return result

def calculate_deficiency_totals(df):
    """Вычисляем взвешенные суммы и ранги для каждого недостатка."""
    deficiency_columns = ['Недостаток','Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']
    importance = pd.to_numeric(df.iloc[1][deficiency_columns], errors='coerce').values
    scores = df.iloc[3:16][deficiency_columns].apply(pd.to_numeric, errors='coerce')

    if pd.isna(importance).any():
        print(f"Ошибка: в значениях важности есть пропуски: {importance.tolist()}")
        return None
    if scores.isna().any().any():
        print(f"Ошибка: в оценках есть пропуски: {scores.isna().sum().to_dict()}")
        return None

    deficiency_sums = scores.sum(axis=0).values #суммы по недостаткам (вертикальные)
    weighted_totals = deficiency_sums * importance #верхние суммы умноженные на веса недостатков

    DEF = ['Много теории, но мало практики.'] + DEFICIENCIES

    result = pd.DataFrame({
        'Недостаток': DEF,
        'Веса': importance,
        'Сумма оценок': deficiency_sums,
        'Взвешенная сумма': weighted_totals
    })

    result['Ранг'] = result['Взвешенная сумма'].rank(ascending=False, method='min').astype(int)
    return result

def save_results(all_disciplines, all_deficiencies, first_table, output_file):
    """Сохраняем агрегированные результаты, таблицу со средними и графики в Excel."""
    try:
        disciplines_img, deficiencies_img = create_charts(all_disciplines, all_deficiencies)

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Сохраняем исходную таблицу со средними значениями
            first_table.to_excel(writer, sheet_name='Results', index=False, startrow=0)
            # Сохраняем таблицы дисциплин и недостатков
            all_disciplines.to_excel(writer, sheet_name='Results', index=False, startrow=len(first_table) + 3)
            all_deficiencies.to_excel(writer, sheet_name='Results', index=False, startrow=len(first_table) + len(all_disciplines) + 6)

            workbook = writer.book
            worksheet = writer.sheets['Results']

            img1 = Image(disciplines_img)
            worksheet.add_image(img1, f'A{len(first_table) + len(all_disciplines) + len(all_deficiencies) + 9}')

            img2 = Image(deficiencies_img)
            worksheet.add_image(img2, f'A{len(first_table) + len(all_disciplines) + len(all_deficiencies) + 29}')

        print(f"Агрегированные результаты, таблица со средними и графики сохранены в {output_file}")
    except Exception as e:
        print(f"Ошибка при сохранении файла {output_file}: {e}")

def process_multiple_files(input_dir, output_file, log_file):
    """Обрабатываем все .xlsx и .xls файлы и агрегируем результаты."""
    setup_logging(log_file)

    file_paths = glob.glob(os.path.join(input_dir, "*.xlsx")) + glob.glob(os.path.join(input_dir, "*.xls"))

    if not file_paths:
        error_msg = f"В директории {input_dir} не найдено .xlsx или .xls файлов"
        print(error_msg)
        logging.error(error_msg)
        return

    disciplines_list = [] #из функции calculate_sumproduc
    deficiencies_list = [] #из calculate_deficiency_totals
    numeric_data = []  # Для хранения числовых данных (весов и оценок)
    first_table = None  # Для хранения таблицы из первого файла

    for i, file_path in enumerate(file_paths):
        print(f"\nОбработка файла: {file_path}")
        try:
            df = read_excel_file(file_path)
            if df is None:
                logging.error(f"Не удалось прочитать файл {file_path}")
                continue

            if not check_table_structure(df):
                logging.error(f"Файл {file_path} не прошёл проверку структуры")
                continue

            # Ограничиваем таблицу первыми 9 столбцами и 16 строками
            df = df.iloc[:16, :9].copy()

            # Сохраняем таблицу из первого файла
            if first_table==None:
                first_table = df.copy()

            # Извлекаем числовые данные (весов и оценок)
            deficiency_columns = ['Недостаток','Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']
            importance = pd.to_numeric(df.iloc[1][deficiency_columns], errors='coerce')
            scores = df.iloc[3:16][deficiency_columns].apply(pd.to_numeric, errors='coerce')
            numeric_data.append(pd.concat([importance.to_frame().T, scores], axis=0))

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

    if not disciplines_list or not deficiencies_list or first_table is None:
        error_msg = "Не удалось обработать ни один файл"
        print(error_msg)
        logging.error(error_msg)
        return

    # Рассчитываем средние значения для числовых ячеек
    if numeric_data:
        mean_numeric = pd.concat(numeric_data).groupby(level=0).mean()
        # Обновляем числовые значения в first_table
        deficiency_columns = ['Недостаток','Unnamed: 3', 'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 7', 'Unnamed: 8']
        first_table.loc[1, deficiency_columns] = mean_numeric.iloc[0].values
        first_table.loc[3:15, deficiency_columns] = mean_numeric.iloc[1:].values
        # Добавляем новые столбцы: Негативный рейтинг в абсолютных и относительных единицах
        # Расчёт для строк 3–15 (дисциплины, индексы 2–14)
        mainabs_rating = first_table.loc[1, deficiency_columns].sum()*10 #по недостаткам и 10 - макс абс негатив рейтинг
        for idx in range(3, 16):  # Индексы 2–14 соответствуют строкам 3–15 (дисциплины)
            # Негативный рейтинг в абсолютных единицах = сумма по токам * 8.43
            abs_rating = (first_table.loc[idx, deficiency_columns]*first_table.loc[1, deficiency_columns]).sum()
            # Негативный рейтинг в относительных единицах = (среднее / 10) * 100
            rel_rating = (abs_rating / mainabs_rating) * 100
            first_table.loc[idx, 'Негативный рейтинг в абсолютных единицах'] = round(abs_rating, 2)
            first_table.loc[idx, 'Негативный рейтинг в относительных единицах'] = f"{round(rel_rating, 2)}%"

        # Добавляем заголовки для новых столбцов в строку 2 (индекс 1)
        first_table.loc[1, 'Негативный рейтинг в абсолютных единицах'] = 'Негативный рейтинг в абсолютных единицах'
        first_table.loc[
            1, 'Негативный рейтинг в относительных единицах'] = 'Негативный рейтинг в относительных единицах'

    # Агрегируем дисциплины
    all_disciplines = pd.concat(disciplines_list, ignore_index=True)
    all_disciplines = all_disciplines.groupby('Дисциплина', sort=False).agg({
        'СУММПРОИЗВ': ['mean', 'std']
    }).reset_index()
    all_disciplines.columns = ['Дисциплина', 'Среднее СУММПРОИЗВ', 'Станд. отклонение СУММПРОИЗВ']
    all_disciplines['order'] = all_disciplines['Дисциплина'].map({d: i for i, d in enumerate(DISCIPLINES)})
    all_disciplines = all_disciplines.sort_values('order').drop('order', axis=1).reset_index(drop=True)
    all_disciplines['Ранг'] = all_disciplines['Среднее СУММПРОИЗВ'].rank(ascending=False, method='min').astype(int)

    # Агрегируем недостатки
    DEF = ['Много теории, но мало практики.'] + DEFICIENCIES
    all_deficiencies = pd.concat(deficiencies_list, ignore_index=True)
    all_deficiencies = all_deficiencies.groupby('Недостаток', sort=False).agg({
        'Веса': 'mean',
        'Сумма оценок': 'mean',
        'Взвешенная сумма': ['mean', 'std']
    }).reset_index()
    all_deficiencies.columns = ['Недостаток', 'Среднее Веса', 'Средняя Сумма оценок', 'Средняя Взвешенная сумма', 'Станд. отклонение Взвешенной суммы']
    all_deficiencies['order'] = all_deficiencies['Недостаток'].map({d: i for i, d in enumerate(DEF)})
    all_deficiencies = all_deficiencies.sort_values('order').drop('order', axis=1).reset_index(drop=True)
    all_deficiencies['Ранг'] = all_deficiencies['Средняя Взвешенная сумма'].rank(ascending=False, method='min').astype(int)

    # Сохраняем результаты
    save_results(all_disciplines, all_deficiencies, first_table, output_file)
    print(f"\nОбработка завершена. Лог ошибок сохранён в {log_file}")

def create_charts(all_disciplines, all_deficiencies):
    """Создаём столбчатые диаграммы для дисциплин и недостатков."""
    plt.figure(figsize=(12, 8))
    bars = plt.bar(all_disciplines['Дисциплина'], all_disciplines['Среднее СУММПРОИЗВ'], color='#1f77b4')
    plt.title('Средние баллы по дисциплинам')
    plt.xlabel('Дисциплина')
    plt.ylabel('Средние баллы')
    plt.xticks(rotation=45, ha='right', fontsize=8)
    plt.subplots_adjust(bottom=0.3)
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, yval, f'{yval:.1f}', va='bottom', fontsize=8)
    disciplines_img = BytesIO()
    plt.savefig(disciplines_img, format='png', bbox_inches='tight')
    plt.close()
    disciplines_img.seek(0)

    plt.figure(figsize=(12, 8))
    bars = plt.bar(all_deficiencies['Недостаток'], all_deficiencies['Средняя Взвешенная сумма'], color='#ff7f0e')
    plt.title('Средняя Взвешенная сумма по недостаткам')
    plt.xlabel('Недостаток')
    plt.ylabel('Средняя Взвешенная сумма')
    plt.xticks(rotation=45, ha='right', fontsize=8)
    plt.subplots_adjust(bottom=0.3)
    for bar in bars:
        yval = bar.get_height()
        plt.text(bar.get_x() + bar.get_width()/2, yval, f'{yval:.1f}', va='bottom', fontsize=8)
    deficiencies_img = BytesIO()
    plt.savefig(deficiencies_img, format='png', bbox_inches='tight')
    plt.close()
    deficiencies_img.seek(0)

    return disciplines_img, deficiencies_img

def main():
    input_dir = os.path.join(os.getcwd(), "source")
    output_file = os.path.join(os.getcwd(), "output_results.xlsx")
    log_file = os.path.join(os.getcwd(), "errors.log")

    process_multiple_files(input_dir, output_file, log_file)

if __name__ == "__main__":
    main()