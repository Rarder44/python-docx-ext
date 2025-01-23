from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table,_Cell
from docx.document import Document as DocumentObject
from docx.blkcntnr import BlockItemContainer
from copy import deepcopy



#TODO: 
# - gesisco un hashmap dei placeholder??


class EzDocxTemplate:
    
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
                    print("ERR: malformed keywordBracket in paragraph: ", par.text)
                


    def ReplaceAll(data, element):
        #blocco contenente altri blocchi
        if isinstance(element,DocumentObject):
            if element.sections[0].footer is not None:
                EzDocxTemplate.ReplaceAll(data,element.sections[0].footer)
            if element.sections[0].header is not None:
                EzDocxTemplate.ReplaceAll(data,element.sections[0].header)

           

        if isinstance(element,DocumentObject) or isinstance(element,BlockItemContainer):          
            for loop_element in element.iter_inner_content():
                EzDocxTemplate .ReplaceAll(data,loop_element)

        #tabella
        elif isinstance(element,Table):
            element:Table
            for row in element.rows:
                for cell in row.cells:
                    EzDocxTemplate .ReplaceAll(data,cell)
            
        #paragro ( effetto la sostituzione)
        elif isinstance(element,Paragraph):
            #print(element.text)
            EzDocxTemplate .ReplaceInParagraph(data,element)
        else:
            raise "ERR! type not recognized!"+str(element)


    def findPlaceholder(element,placeholderName):
        #blocco contenente altri blocchi
        if isinstance(element,DocumentObject) or isinstance(element,BlockItemContainer):
            for loop_element in element.iter_inner_content():
                p = EzDocxTemplate .findPlaceholder(loop_element,placeholderName)
                if p:
                    return p

        #tabella
        elif isinstance(element,Table):
            element:Table
            for row in element.rows:
                for cell in row.cells:
                    p = EzDocxTemplate .findPlaceholder(cell,placeholderName)
                    if p: 
                        return p
            
        #paragrafo
        elif isinstance(element,Paragraph):
            if placeholderName in element.text:
                return element


        else:
            raise "ERR! type not recognized!"+str(element)


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
        """Given a table and an array of "row" (array), fills the table with that data"""

        for row in data:
            row_cells = table.add_row().cells
            for i in range(len(row)):
                row_cells[i].text = str(row[i])


    def FindTableToPopulate(element,tableName):
        if (tableName.startswith("${{") and tableName.endswith("}}")):
            pass
        else:
            tableName="${{"+tableName+"}}"

        return EzDocxTemplate.__FindTableToPopulate(element,tableName)
    

    def __FindTableToPopulate(element,tableName):
        """The function already removes the row with the table name, from there new rows will be added"""

        #blocco contenente altri blocchi
        if isinstance(element,DocumentObject) or isinstance(element,BlockItemContainer):
            for loop_element in element.iter_inner_content():
                el = EzDocxTemplate .FindTableToPopulate(loop_element,tableName)
                if el!=None:
                    return el

        #tabella
        elif isinstance(element,Table):
            element:Table
            for row in element.rows:
                for cell in row.cells:
                    if tableName in cell.text:
                        EzDocxTemplate ._remove_row(element,row)
                        return element
        return None


    def copyElement(element):
        if isinstance(element,Table):
            element:Table
            return deepcopy(element._tbl)
        elif isinstance(element,Paragraph):
            return deepcopy(element._p)
        else:
            raise ValueError("Error! Unrecognized type to copy!")

    def copy_element_after_paraph(placeholder:Paragraph,element=None,elements=[],deletePlaceholder=True):
        if element:
            elements.insert(0,element)

        if elements:
            for e in elements:
                placeholder._p.addprevious(EzDocxTemplate.copyElement(e))
                #placeholder._p.addprevious
            
        if deletePlaceholder:
            placeholder._p.getparent().remove(placeholder._p)



