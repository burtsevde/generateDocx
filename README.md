# generateDocx
Главная программа "genDocx.py".

Скрипт позволяет создать шаблоны .DOCX договоров (в директории /docs) на основе приложенного файла "Конструктор.xlsx". В "конструкторе" на странице "data" указываются пунты договора и флаги, которые показывают, что этот пункт должен быть включен в тот или иной документ. На странице "Settings" указываются общие для всех документов переменные. Также скрипт автоматически разбирает текст в ячейке эксель и разбивает его на блоки, к которым будет применен тот или иной стиль. Для разбиения текста применяется библиотека mistletoe.

Чтобы была возможность передать программу рядовым пользователям, была создана возможность создать исполняемый файл .EXE, который создается в директории /output.
 

Copyright (c) 2022 Burtsev Denis

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
