    Скрипт позволяет создать шаблоны .DOCX договоров (в директории /docs) на основе приложенного файла
"Конструктор.xlsx". В "конструкторе" на странице "data" указываются пунты договора и флаги, которые показывают,
что этот пункт должен быть включен в тот или иной документ. На странице "Settings" указываются общие для всех
документов переменные. Также скрипт автоматически разбирает текст в ячейке эксель и разбивает его на блоки, к которым
будет применен тот или иной стиль. Для разбиения текста применяется библиотека mistletoe.

Инструкция:
1. Заменить файл "sign_1.jpg" на фотографию подписи, которая должна быть вставлена в документе.
Имя нового файла должно быть "sign_1.jpg"

2. Изменить файл "Контруктор.xlsx", как необходимо.

3. Запустить файл "genDocx.exe"



Чтобы в Word документ были вставлены переменные и соответствующие стили необходимо:

«$name» - где name, это переменная, которая должна быть вставлена в шаблон. Название переменной будет начинатся с $.

**text** - text будет выделен жирным

*text* - text будет выделен курсивом
