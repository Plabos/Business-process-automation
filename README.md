## Описание проекта:

Ежедневно от банка приходят выписки о том, сколько денежных средств поступило на счет. Эти деньги необходимо перевести клиенту с удержанием комиссии. Для этого необходимо:
•	загрузить выписки в программу
•	посчитать комиссию и сумму к перечислению
•	выдать итоговый файл (отчет) с конкретным набором столбцов (в банковских выписках нет всей информации о платеже, поэтому ее необходимо добавить из другого файла).

## Задачи проекта: 

Написать скрипт на Python для автоматизации бизнес-процесса

## Инструменты:

- Python
- Pandas
- Os
- Datetime
- Openpyxl

## Полученные результаты:
- Написаны функции, которые подбирают по названию необходимые файлы в текущей директории и формируют из них общий датафрейм.
- Написана функция, которая производит необходимые расчеты.
- Написана функция, которая создает из датафрейма отчет в формате .xlcx, выравнивает столбцы по ширине, рисует границы ячеек.
