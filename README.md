## Скрипт для анализа финансового состояния компаний на основе данных с сайта Налоговой службы РФ.

Предварительно создать директорию /data.  
В ней создать три поддериктории:
- /checked (можно не создавать, если не планируется сохранения проверенных файлов компаний)
- /err
- /new
- /result

1. Скачать отчетность компании с сайта налоговой: https://bo.nalog.ru/
Для этого на странице компании нажать на "Скачать таблицей или текстом", затем выбрать "Формат отчета": Excel
в формате xlsx. Обязательно указать "Бухгалтерский баланс". Остальное по желанию.  
1.1. Установить nalog=True (если скачан годовой отчет с сайта налоговой. Если данные вбиты в шаблон Template.xlsx, 
то переменная False).
2. Распаковать архив в папку ./data/new
3. Переименовать файл Settings.dev в Settings.xlsx и добавить в него ИНН и Наименование компании.
4. Запустить скрипт.
5. Файлы компаний, имеющих больше 27 баллов (переменная score_success), перемещаются в директорию data/result 
для дальнейшего более детального анализа (при необходимости).
6. Наименования компаний c > score_success, а также названия файлов, записываются в файл !success.txt в той же директории.
7. Если скрипту не удается "прочитать" файл с данными по комании, то этот файл помещается в директорию ./data/err.
Скорее всего файл с нестандартными строками по балансу.
8. *Для дальнейшего анализа можно воспользоваться notebook'ом, который находится, в директории ./extra, переместив 
в эту же директорию файл эксель.
В конце файла убрать "1": Analys.ipynb1