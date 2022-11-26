# Импорт пакетов:
import os
import sys
import traceback

import numpy as np
import openpyxl
import pandas as pd

# Чтобы убрать warning: "Workbook contains no default style, apply openpyxl's default "
import warnings
warnings.simplefilter("ignore")

from utils import LOGGER
from datetime import datetime

directory = os.path.dirname(os.path.realpath("__file__"))
PATH = f'{directory}/data/new/'
PATH_TO_WRITE = f'{directory}/data/result'

# Файл для хранения ИНН и названий компаний.
file_settings = 'Settings.xlsx'

# Кол-во баллов, допустимых для дальнейшего анализа
score_success = 27

# Если файл загружен с сайта налоговой, то True (преобразование текстового формата в цифровой
nalog = False

# Файл для записи компаний с суммарным рейтингом за последний отчетный период > score_success
file_success = f'{PATH_TO_WRITE}/!success.txt'

# Анализируемые годы
period = ['12/2019', '12/2020', '12/2021']

# Минимальные коэффициенты используемые в формулах
norma = [0.7, 1.0, 1.73, 0.5, 1.4, 0.5, 0.5, 0.8]

# **Присуждение баллов (ЗС на 1 рубль):**
# 5 баллов - показатель больше нормы на 5% (от 0 до 1.4)
# 4 балла - +-5%  от нормы (от 1.4 до 1.6)
# 3 балла - до 50% от нормы (от 1.6 до 1.8)Сумма баллов на дату
# 2 балла - 0-50% от нормы (от 1.8 до 2)
# 1 балл - меньше 0 (меньше 0 или больше 2).
points = [5, 4, 3, 2, 1]


def main():
    # Загрузка файлов с отчетами.
    list_files, _path = get_files()
    if not list_files:
        LOGGER.warning(f'В {_path} нет файлов с разширением .xlsx для загрузки')
        sys.exit('Нет файлов для анализа')

    list_company = get_company(file_settings)

    if not list_company:
        LOGGER.warning(f'Список для анализа пустой.')
        sys.exit('Нет компаний для анализа')

    for file in list_files:

        replace_p = replace_path(_path)
        balance, inn = get_info(replace_p+file)

        if balance.empty:
            LOGGER.error(f'Файл {file} с балансом компании {list_company[inn]} не стандартный и не может быть прочитан!')
            continue

        # Преобразование данных в цифровой формат
        if nalog:
            balance = str_to_int(balance, period)

        # Датафрейм для финансовых коэффициентов
        koef_df = pd.DataFrame(columns=['Наименование',
                                        period[2],
                                        period[1],
                                        period[0]])

        # *** Далее используются формулы из статьи https://fapvdo.ru/8-poleznyh-formul-dlja-predskazanija-defolta/ ***

        # Формула 1. Погашение текущих обязательств за счет денежных средств
        # Ф1 = (1240 + 1250)/(1510 + 1520 + 1550) =
        # = (Финансовые вложения + Денежные средства и денежные эквиваленты) /
        # / ({)Заемные средства + Кредиторская задолженность + Прочие обязательства)

        # Кода строчек из бух.баланса для числителя и знаменалеля формулы
        codes_numerator_f1 = [1240, 1250]
        codes_denominator_f1 = [1510, 1520, 1550]

        # Расчет первой формулы и запись результата в датафрейм koef_df
        calc = calculation('Ф1. Погашение текущих обязательств за счет денежных средств', period,
                                                balance, codes_numerator_f1, codes_denominator_f1)
        if not calc:
            LOGGER.error(f'Файл с балансом компании {list_company[inn]} не стандартный и не может быть прочитан!')
            continue  # TODO Обработка нестандартных файлов (Пропуски и дополнительные строки)

        koef_df.loc[len(koef_df)] = calc

        # Формула 2. Погашение текущих обязательств за счет денежных средств и дебиторской задолженности.
        # Ф2 = (1230 + 1240 + 1250)/ (1510 + 1520 + 1550) =
        # = (Дебиторская адолженность + Финансовые вложения + Денежные средства и денежные  эквиваленты)/
        # / (Заемные средства + Кредиторская задолженность + Прочие обязательства)

        codes_numerator_f2 = [1230, 1240, 1250]
        codes_denominator_f2 = [1510, 1520, 1550]

        koef_df.loc[len(koef_df)] = calculation(
            'Ф2. Погашение текущих обязательств за счет денежных средств и дебиторской задолженности.', period,
            balance, codes_numerator_f2, codes_denominator_f2)

        # Формула 3. Достаточность средств для погашения краткосрочных обязательств в течение года.
        # Ф3 = (1200) / (1510 + 1520 + 1550) =
        # = Оборотные активы / (Заемные средства + Кредиторская задолженность + Прочие обязательства)

        codes_numerator_f3 = [1200]

        koef_df.loc[len(koef_df)] = calculation(
            'Ф3. Достаточность средств для погашения краткосрочных обязательств в течение года/', period,
            balance, codes_numerator_f3, codes_denominator_f1)

        # Формула 4. Доля оборотных средств в активах.
        # Ф4 = 1200 / 1600 = Оборотные активы / Баланс

        codes_numerator_f4 = [1200]
        codes_denominator_f4 = [1600]

        koef_df.loc[len(koef_df)] = calculation(
            'Ф4. Доля оборотных средств в активах.', period,
            balance, codes_numerator_f4, codes_denominator_f4)

        # Формула 5. Сколько заемных средств на один собственный рубль.
        # Ф5 = (1400 + 1510 + 1520 + 1550) / (1300 + 1530 + 1540) =
        # = (Долгосрочные обязательства + Заемные средства + Кредиторская задолженность + Прочие обязательства) /
        # / (Капитал и резервы + Доходы будущих периодов + Оценочные обязательства)
        # Вопрос: Почему Оценочные обязательства стоят в знаменатели???

        codes_numerator_f5 = [1400, 1510, 1520, 1550]
        codes_denominator_f5 = [1300, 1530, 1540]

        koef_df.loc[len(koef_df)] = calculation(
            'Ф5. Заемные средства на один собственный рубль.', period,
            balance, codes_numerator_f5, codes_denominator_f5)

        # Формула 6. Какая часть оборотных активов финансируется за счет собственных средств.
        # Ф6 = (1300 + 1530 + 1540 - 1100) / 1200 =
        # = (Капитал и резервы + Доходы будущих периодов + Оценочные обязательства - Внеоборотные активы) /
        # / Оборотные активы
        # Вопрос: Почему отнимаем внеоборотные активы???

        codes_numerator_f6 = [1300, 1530, 1540, -1100]
        codes_denominator_f6 = [1200]

        koef_df.loc[len(koef_df)] = calculation(
            'Ф6. Оборотные активы за счет собств. средств', period,
            balance, codes_numerator_f6, codes_denominator_f6)

        # Формула 7. Удельный вес собственного капитала в общей сумме источников финансирования.
        # Ф7 = (1300 + 1530 + 1540) / 1600 =
        # = (Капитал и резервы + Доходы будущих периодов + Оценочные обязательства) / Баланс

        codes_numerator_f7 = [1300, 1530, 1540]
        codes_denominator_f7 = [1600]

        koef_df.loc[len(koef_df)] = calculation(
            'Ф7. Собств. капитал в общей сумме финансирования', period,
            balance, codes_numerator_f7, codes_denominator_f7)

        # Формула 8. Финансовая устойчивость.
        # Ф8 = (1300 + 1530 + 1540 + 1400) / 1600 =
        # = (Капитал и резервы + Доходы будущих периодов + Оценочные обязательства + Долгосрочные обязательства)/Баланс
        # Вопрос: Почему долгосрочные обязательства в числители и вообще присутствуют в этой формуле?

        codes_numerator_f8 = [1300, 1530, 1540, 1400]
        codes_denominator_f8 = [1600]

        koef_df.loc[len(koef_df)] = calculation(
            'Ф8. Финансовая устойчивость', period,
            balance, codes_numerator_f8, codes_denominator_f8)

        # Добавление в датафрейм столбца с минимальными значениями для формул
        koef_df['Норма'] = norma

        # Преобразование полученных значений в баллы
        calculation_point(period, koef_df)
        LOGGER.debug(f'Таблица коэффциентов для {list_company[inn]}: {koef_df}')

        # Суммарный балл за период
        result = sum_years(koef_df, period)

        # Если сумма баллов за последний период больше score_success, то записываем компанию в файл
        if result[period[2]] > score_success:
            info = f'{list_company[inn]};{file}'
            write_success(info)

        # Запись в файл даты запуска скрипта
        write_file(datetime.now().strftime("%d.%m.%Y"), list_company[inn])

        # Расчет риска
        info = ''
        for k, v in result.items():
            risk = calculation_risk(v)
            info = info + f'Сумма баллов для {list_company[inn]} на дату {k} равна: {v}. Риск {risk}.\n'

        # Запись в файл
        write_file(info + '***\n', list_company[inn])
        LOGGER.info(info)


def write_file(text, name):
    name = replace_quote(name)
    with open(f'{PATH_TO_WRITE}/{name}.txt', 'a', encoding='UTF-8') as writer:
        writer.write(text)
        writer.write('\n')


# Запись списка компаний с суммарным рейтингом за последний год больше 27
def write_success(text):
    with open(file_success, 'a', encoding='UTF-8') as writer:
        writer.write(text)
        writer.write('\n')


def get_info(file):
    inn = 0
    try:
        # Загрузка данных из файла excel:
        df = pd.read_excel(file, sheet_name='Balance')

        # Получение ИНН и КПП
        kpp = df.loc[0][2].split(' ')[-1]
        inn = df.columns[2].split(' ')[-1]
        LOGGER.debug(f'Компания с ИНН: {inn} и КПП: {kpp}')

        # Заголовки
        headers = [df.loc[3][0], df.loc[3][3], df.loc[3][8], period[0], period[1], period[2]]
        # TODO переделать "Заголовки для balance: ['1', '2', nan, '12/2019', '12/2020', '12/2021']"
        LOGGER.debug(f'Заголовки для balance: {headers}')

        # Переменная для датафрейма с балансом.
        balance = pd.DataFrame(columns=headers)

        # Создание датафрейма с необходимыми данными:
        row = []
        for i in range(5, df.shape[0]):
            for j in [0, 3, 8, 10, 13, 16]:
                row.append(df.loc[i][j])
            balance.loc[i - 5] = row
            row.clear()

        return balance, inn

    except Exception:
        LOGGER.error(traceback.format_exc())

        return pd.DataFrame(), inn


# Получение название компаний с их ИНН из файла с настройками
def get_company(settings):
    # Словарь для хранения "инн" - название компании
    inn_company = {}

    try:
        # Переменная для загрузки книги excel c инн и названиями компаний
        wookbook = openpyxl.load_workbook(settings)

        # Переменная для опредения активного листа
        worksheet = wookbook.active

        # Чтение файла по строчкам и запись в словарь
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            inn = str(row[0])
            company = str(row[1])
            inn_company[inn] = company
        LOGGER.debug(f'Список компаний для анализа: {inn_company}')
    except Exception:
        LOGGER.error(traceback.format_exc())

    return inn_company


# Получение файлов с отчетами с сайта ФНС
def get_files():

    list_files = []
    try:
        for root, dirs, files in os.walk(PATH):
            for file in files:
                if file.endswith('.xlsx'):
                    list_files.append(file)
        LOGGER.debug(f'Список загружаемых файлов: {list_files}')
    except Exception:
        LOGGER.error(traceback.format_exc())

    return list_files, PATH


# Преобразование данных из строки в числовой формат:
def str_to_int(df, years):
    for year in years:
        df[year] = df[year].str.replace(' ', '')
        df[year] = pd.to_numeric(df[year], errors='coerce').fillna(0).astype(np.int64)
    return df


# Общая формула для расчетов:
def calculation(name, years, data, codes_numerator, codes_denominator):
    try:

        result = [name]

        for year in years:
            # Числитель:
            numerator = 0

            # Знаменатель:
            denominator = 0
            for code in codes_numerator:
                if code > 0:
                    numerator += data[data['Код строки'] == str(code)][year].values[0]
                else:
                    code = -1 * code
                    numerator -= data[data['Код строки'] == str(code)][year].values[0]

            for code in codes_denominator:
                denominator += data[data['Код строки'] == str(code)][year].values[0]

            result.append(format(numerator / denominator, '.4f'))

        return result

    except Exception:
        LOGGER.error(traceback.format_exc())
        return None

# Функция для расчета баллов
def calculation_point(period, df, choicelist=None):

    if choicelist is None:
        choicelist = points

    for year in period:
        value = df[year].astype('float') / df['Норма'].astype('float')

        condition1 = value > 1.05
        condition2 = (value <= 1.05) & (value >= 0.95)
        condition3 = (value < 0.95) & (value >= 0.50)
        condition4 = (value < 0.50) & (value >= 0)
        condition5 = value < 0

        condlist = [condition1, condition2, condition3, condition4, condition5]

        df[f"Балл{year}"] = 0
        df[f"Балл{year}"] = np.select(condlist, choicelist)

        # Дополнительные условия для "Ф5. Заемные средства на один собственный рубль."
        add_value = float(df.loc[4, f"{year}"])
        add_condition1 = (add_value >= 0) & (add_value <= 1.4)
        add_condition2 = (add_value > 1.4) & (add_value <= 1.6)
        add_condition3 = (add_value > 1.6) & (add_value <= 1.8)
        add_condition4 = (add_value > 1.8) & (add_value <= 2.0)
        add_condition5 = (add_value < 0) | (add_value > 2)
        add_condlist = [add_condition1, add_condition2, add_condition3, add_condition4, add_condition5]
        df.loc[4, f"Балл{year}"] = np.select(add_condlist, choicelist)


# Функция для определения риска по баллам
def calculation_risk(value):

    if (value >= 8) & (value < 16):
        risk = 'Максимум'
    elif (value >= 16) & (value < 22):
        risk = 'Высокий'
    elif (value >= 22) & (value < 28):
        risk = 'Умеренный'
    elif (value >= 28) & (value < 34):
        risk = 'Средний'
    elif (value >= 34) & (value <= 40):
        risk = 'Минимум'
    else:
        risk = 'Невозможно определить.'

    return risk


# Суммирование баллов за период
def sum_years(df, period):
    dict_sum = {}
    for year in period:
        dict_sum[year] = df[f'Балл{year}'].sum()

    return dict_sum


def replace_path(path):
    LOGGER.debug(f'Путь до замены: {path}')
    path = path.replace('\\', '/')
    LOGGER.debug(f'Путь после замены: {path}')
    return path


def replace_quote(name):
    LOGGER.debug(f'Имя файла до замены: {name}')
    name = name.replace('"', '')
    LOGGER.debug(f'Имя файла после замены: {name}')
    return name


if __name__ == '__main__':
    main()