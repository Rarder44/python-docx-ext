from docx import Document

import os
from EzDocxTemplate import EzDocxTemplate


# set the current working directory
os.chdir(os.path.dirname(os.path.realpath(__file__)))


src = "docx/1_mainSource.docx"
src2 = "docx/2_otherSource.docx"
dst = "docx/out/out.docx"
document = Document(src)


data = {
    "NAME": "Mario Rossi",
    "SECTOR": "IT",
    "INFO1": "AAA",
    "INFO2": "BBB",
    "FLAG1": "🞋",
    "FLAG2": "X",
    "FLAG3": "🞋",
    "FLAG4": "X",
    "RATE": "6",
    "HOUR": 40,
    "YEAR": "2023",
}

tables = {
    "TAB_HOUR": [
        ["Mario", "Rossi", 5],
        ["Pippo", "Verdi", 3],
        ["Marco", "Gialli", 7],
        ["Gianni", "Neri", 8],
    ]
}


# sostituisci valori

EzDocxTemplate.ReplaceAll(data, document)


# popola tabella
for tableName in tables:
    table = EzDocxTemplate.FindTableToPopulate(document, tableName)
    if table is not None:
        EzDocxTemplate.PopulateTable(tables[tableName], table)


# aggiunta "massiva"
document2 = Document(src2)
# trovo il punto di inserimento
# scorro tutti gli elementi del document2 ( ricorsivo )
# per ciascunno lo copio e lo inserisco nel document originale nella posizione
# rimuovo il paragrafo di placeholder


p = EzDocxTemplate.findPlaceholder(document, "${{OTHER_PAGE}}")
elements = [e for e in document2.iter_inner_content()]
EzDocxTemplate.copy_element_after_paraph(p, elements=elements)


document.save(dst)
