# generateDocx
Главная программа "genDocx.py".

Скрипт позволяет создать шаблоны .DOCX договоров (в директории /docs) на основе приложенного файла "Конструктор.xlsx". В "конструкторе" на странице "data" указываются пунты договора и флаги, которые показывают, что этот пункт должен быть включен в тот или иной документ. На странице "Settings" указываются общие для всех документов переменные. Также скрипт автоматически разбирает текст в ячейке эксель и разбивает его на блоки, к которым будет применен тот или иной стиль. Для разбиения текста применяется библиотека mistletoe.

Чтобы была возможность передать программу рядовым пользователям, была создана возможность создать исполняемый файл .EXE, который создается в директории /output.
