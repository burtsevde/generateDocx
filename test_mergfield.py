from __future__ import print_function
from mailmerge import MailMerge
from datetime import date
import os

path = os.getcwd()
new_doc_path = os.path.join(path, 'docs')

template = new_doc_path + "\offer_2020_pdl_classic.docx"

document = MailMerge(template)
fields = document.get_merge_fields()
print(fields)
m_fields = {}

print(m_fields)
document.merge(
    offerNumber='asd-1111111',
)

document.write(new_doc_path + '\\test-output.docx')