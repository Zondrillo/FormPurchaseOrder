### Краткое описание

Проект позволяет автоматизировать формирование заявки на закупку продукции.

Алгоритм формирует:

* Общие ТЗ для всех станций/тепловых сетей с разбивкой по статьям бюджета (Ремонт, Эксплуатация, Инвестиции)

* Раздельные ТЗ для каждой станции/тепловых сетей и для каждой статьи бюджета (Ремонт, Эксплуатация, Инвестиции)

* Сведения о начальной минимальной цене продукции
---

### Структура проекта
 - classes
   - form_technical_task_common.py - формирует общие ТЗ
   - form_technical_task_sep.py - формирует раздельные ТЗ
   - form_nmp_info.py - формирует сведения о НМЦ (начальные максимальные цены)
   - base_class.py - базовый класс, от которого наследуются классы form_technical_task_common & form_technical_task_sep
 - configs
   - config.py - конфигурационный файл
   - texts.py - файл с текстовками
 - utilities
   - helpers.py - вспомогательные методы
 - main.py - главный скрипт, в котором происходит вся магия
---

### Зависимости
* [Python](https://www.python.org/downloads/) >= 3.8.5
* [pandas](https://pandas.pydata.org/)
* [numpy](https://numpy.org/)
* [xlsxwriter](https://xlsxwriter.readthedocs.io/)
* [openpyxl](https://openpyxl.readthedocs.io/en/stable/)