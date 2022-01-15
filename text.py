import mistletoe
from mistletoe.ast_renderer import ASTRenderer
import json
import regex as re
from docx_extension import add_mergefield

def set_paramed_text(paragraph, text, styles=None):
    param = str(text)

    if param == '' or param == 'nan':
        return paragraph

    m = re.split(r'(«\$.*?»)', param)

    for i in m:
        if i[:2] == '«$':
            paragraph.add_run()._r.append(add_mergefield(i))
        else:
            paragraph.add_run(i)
        # add styles on text
        #     get last Run
        if 'bold' in styles:
            paragraph.runs[-1].bold = True
        if 'italic' in styles:
            paragraph.runs[-1].italic = True

    return paragraph

def compose_text(paragraph, dict, styles=None):
    for i in dict:
        # styles = []

        if i['type'] == 'Strong':
            styles.append('bold')
        if i['type'] == 'Emphasis':
            styles.append('italic')
        if i['type'] == 'LineBreak':
            paragraph.add_run('\n')

        if 'content' in i:
            set_paramed_text(paragraph, i['content'], styles)
        else:
            compose_text(paragraph, i['children'], styles)
            styles = styles[:-1] #убираем последний используемый стиль

def text_stylization(paragraph, text):
    if text == '':
        return paragraph
    parsed_text = json.loads(
        mistletoe.markdown(
            text,  # текст для обработки
            ASTRenderer
        )
    )['children'][0]['children']

    compose_text(paragraph, parsed_text, [])