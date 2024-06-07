from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table,_Cell
from docx.document import Document as DocumentObject
from docx.blkcntnr import BlockItemContainer
from copy import deepcopy


#TODO: faccio un hashmap con tutti i placeholder??


src="demo.docx"
src2="_otherSRC.docx"
dst="out.docx"
document = Document(src)


data={
    "NOME":"Pozzi Luca",
    "DISCIPLINA":"Informatica",
    "CLASSE":"4BI",
    "INDIRIZZO":"Informatico",
    "ITC_ITIS":"X",
    "BIEN_TRIEN":"X",
    "NUM_ALUNNI":"23"
}

tables={
    "TAB_VOTI":
    [
        ["Mario","Rossi",5],
        ["Pippo","Verdi",3],
        ["Marco","Gialli",7],
        ["Gianni","Neri",8],
    ]
}

def ReplaceInParagraph(data,par:Paragraph):

    for keyword in data:
        keywordBracket="${{"+keyword+"}}"
        if keywordBracket in par.text:
            
            inline = par.runs

            repl=False
            for i in range(len(inline)):
                if keywordBracket in inline[i].text:
                    text = inline[i].text.replace(str(keywordBracket), str(data[keyword]))
                    inline[i].text = text
                    repl=True
            if not repl:
                print("ERR: malformed keywordBraket nel paragrafo: ",par.text)
            


def ReplaceAll(data, element):
    #blocco contenente altri blocchi
    if isinstance(element,DocumentObject) or isinstance(element,BlockItemContainer):
        for loop_element in element.iter_inner_content():
            ReplaceAll(data,loop_element)

    #tabella
    elif isinstance(element,Table):
        element:Table
        for row in element.rows:
            for cell in row.cells:
                ReplaceAll(data,cell)
        
    #paragro ( effetto la sostituzione)
    elif isinstance(element,Paragraph):
        #print(element.text)
        ReplaceInParagraph(data,element)


    else:
        print("BHO",element)


def findPlaceholder(element,placeholderName):
    #blocco contenente altri blocchi
    if isinstance(element,DocumentObject) or isinstance(element,BlockItemContainer):
        for loop_element in element.iter_inner_content():
            p = findPlaceholder(loop_element,placeholderName)
            if p:
                return p

    #tabella
    elif isinstance(element,Table):
        element:Table
        for row in element.rows:
            for cell in row.cells:
                p = findPlaceholder(cell,placeholderName)
                if p: 
                    return p
        
    #paragrafo
    elif isinstance(element,Paragraph):
        if placeholderName in element.text:
            return element


    else:
        raise "ERR! non conosco questo tipo!"+str(element)

    return None


def _remove_row(table, row):
    tbl = table._tbl
    tr = row._tr
    tbl.remove(tr)

    

def _findCellIndexes(table:Table,cell:_Cell):
    for r in range(len(table.rows)):
        for c in range(len(table.rows[r].rows)):
            if table.rows[r].cells[c]==cell:
                return r,c
            
    return None
        

def PopulateTable(data,table:Table):
    """Data una tabella ed un array di "row" ( array ), riempe la tabella con quei dati"""
    for row in data:
        row_cells = table.add_row().cells
        for i in range(len(row)):
            row_cells[i].text = str(row[i])


def FindTableToPopulate(element,tableName):
    """la funzione rimuove già la riga con il nome della tabella, da li in poi verranno aggiunte le righe"""

    #blocco contenente altri blocchi
    if isinstance(element,DocumentObject) or isinstance(element,BlockItemContainer):
        for loop_element in element.iter_inner_content():
            el = FindTableToPopulate(loop_element,tableName)
            if el!=None:
                return el

    #tabella
    elif isinstance(element,Table):
        element:Table
        for row in element.rows:
            for cell in row.cells:
                if tableName in cell.text:
                    _remove_row(element,row)
                    return element
    return None



#sostituisci valori

ReplaceAll(data,document)
    


#popola tabella
for tableName in tables:
    tableNameKeyword = "${{"+tableName+"}}"
    table = FindTableToPopulate(document,tableNameKeyword)
    if table!=None:
        PopulateTable(tables[tableName],table)



#aggiunta "massiva"
document2 = Document(src2)
#trovo il punto di inserimento
#scorro tutti gli elementi del document2 ( ricorsivo ) 
#per ciascunno lo copio e lo inserisco nel document originale nella posizione 
#rimuovo il paragrafo di placeholder
#paragraph._p.getparent().remove(paragraph._p)

#
#
#def copy_table_after(table, paragraph):
#    tbl, p = table._tbl, paragraph._p
#    new_tbl = deepcopy(tbl)
#    p.addnext(new_tbl)
#

def copyElement(element):
    if isinstance(element,Table):
        element:Table
        return deepcopy(element._tbl)
    elif isinstance(element,Paragraph):
        return deepcopy(element._p)
    else:
        raise "ERR! Tipo da copiare non riconosciuto!!"
    

def copy_element_after_paraph(placeholder:Paragraph,element=None,elements=[],deletePlaceholder=True):
    if element:
        elements.insert(0,element)

    if elements:
        for e in elements:
            placeholder._p.addprevious(copyElement(e))
            #placeholder._p.addprevious
        
    if deletePlaceholder:
        placeholder._p.getparent().remove(placeholder._p)
        

p = findPlaceholder(document,"${{OTHER}}")

elements = [e for e in document2.iter_inner_content()]
copy_element_after_paraph(p,elements=elements)




document.save(dst)

