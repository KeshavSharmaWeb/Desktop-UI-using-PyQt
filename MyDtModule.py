##############################
##### Convert Word or PDF ####
#### To a XML or JSON file ###
##############################
import json
import csv
import re
import os
import sys
import xlrd
import pandas as pd
import xml.etree.ElementTree as ET
from docx import Document
from docx.oxml.ns import qn
import textract
import fitz
from pdfminer.high_level import extract_text
import lxml.etree as etree
from dict2xml import dict2xml
import dict2xml
# nsmap = {None: "ROOT"}
# root = etree.Element('Roughness-Profiles', nsmap=nsmap)

alphabets = "([A-Za-z])"
prefixes = "(Mr|St|Mrs|Ms|Dr)[.]"
suffixes = "(Inc|Ltd|Jr|Sr|Co)"
starters = "(Mr|Mrs|Ms|Dr|He\s|She\s|It\s|They\s|Their\s|Our\s|We\s|But\s|However\s|That\s|This\s|Wherever)"
acronyms = "([A-Z][.][A-Z][.](?:[A-Z][.])?)"
websites = "[.](com|net|org|io|gov)"
xlsx_filter_condition = [{
                'Risk Description': ['Risk Description', 'Risk Objective'],
                'Control Owner': ['Owner', 'Control Owner', 'Control Stakeholder', 'Stakeholder', 'Business Owner'],
                'Control Description': ['Control Description'],
                'Control Objective': ['Control Objective'],
                'Business Unit': ['BU', 'Business Unit', 'Bussiness Function'],
                'Process Domain' : ['Process Domain', 'process Area', 'Process Type', 'Process','Process Name'],
               'Process ID' :['Process ID', 'Process Number', 'Process Code', 'Process#', 'Process No.', 'Process Reference', 'Process Ref.'],
               'Process Description': ['Process Description', 'Process Objective'],
               'Risk ID': ['Risk ID', 'Risk Number', 'Risk Code', 'Risk# ', 'Risk No.', 'Risk Reference', 'Risk Ref'],
               'Control ID': ['Control ID', 'Control Number', 'Control Code', 'Control #', 'Control No.', 'Control Reference', 'Control Ref'],
               'Control Type (Preventive , Detective)': ['Control Category', 'Preventative / Detective', 'Preventive , Detective', 'Control Nature'],
               'Control Type (Manual, Semi-automated, Automated)': ['Auto/Manual','Control Type','Automated/Manual', 'Auto/Semi-Auto/Manual', 'Manual, Semi-automated, Automated', '(Manual / Semi-automated / Automated)', 'Control Nature'],
               'Control Frequency' : ['Control Frequency', 'Control Freq.', 'Frequency', 'Control Occurrence'],
               'Key control' : ['key/non-Key', 'Control Status', 'Key Controls','Key Control']
               }]

#########################
#### docx processing ####
#########################


buffer = []

infolist = []

output_file_path = ""
input_folder_path = ""
cwd = ""
output_folder_path = ""
output_type = ""
infobuffer = []
nsmap = {None: "ROOT"}
root = etree.Element('Roughness-Profiles', nsmap=nsmap)
def fixtocorrecttype(outputfile_content):
    arrayrows = []
    for key in outputfile_content[0]:
        nrows = len(outputfile_content[0][key])
        arrayrows.append(nrows)
    maxrow = max(arrayrows)
    for key in outputfile_content[0]:
        for ii in range(maxrow-len(outputfile_content[0][key])):
            outputfile_content[0][key].append('')

def make_json_csv(csvFilePath, jsonFilePath,output_folder_path):
    try:
        with open(csvFilePath, encoding='utf-8') as csvFile:
            csvReader = csv.DictReader(csvFile)
            os.chdir(output_folder_path)
            with open(jsonFilePath, 'w', encoding='utf-8') as jsonFile:
                jsonFile.write(json.dumps(list(csvReader), indent=4))
    except Exception as e:
        with open(csvFilePath) as csvFile:
            csvReader = csv.DictReader(csvFile)
            os.chdir(output_folder_path)
            with open(jsonFilePath, 'w', encoding='utf-8') as jsonFile:
                jsonFile.write(json.dumps(list(csvReader), indent=4))

def make_json_xml(csvFilePath, xmlFilePath,output_folder_path):
    datalist = []
    printdata = []
    try:
        with open(csvFilePath, encoding='utf-8') as csvFile:
            data = csv.DictReader(csvFile)
            datalist = list(data)
        for i in datalist:
            printdata.append("<item>\n")
            for k in i:
                printdata.append("<" + k.replace(" ", "_") + "> " + i[k] + " </" + k.replace(" ", "_") + ">")
            printdata.append("\n</item>\n")
        os.chdir(output_folder_path)
        with open(xmlFilePath, 'w', encoding='utf-8') as xmlFile:
            for x in printdata:
                xmlFile.write(x)
    except Exception as e:
        with open(csvFilePath) as csvFile:
            data = csv.DictReader(csvFile)
            datalist = list(data)
        for i in datalist:
            printdata.append("<item>\n")
            for k in i:
                printdata.append("<" + k.replace(" ", "_") + "> " + i[k] + " </" + k.replace(" ", "_") + ">")
            printdata.append("\n</item>\n")
        os.chdir(output_folder_path)
        with open(xmlFilePath, 'w') as xmlFile:
            for x in printdata:
                xmlFile.write(x)

def xlsx_to_json(fname, fname2,output_folder_path,resfiles):
    workbook = xlrd.open_workbook(fname)
    os.chdir(output_folder_path)
    if (len(workbook.sheet_names()) > 1):
        for sheetno in range(0, len(workbook.sheet_names())):
            sheet = workbook.sheet_by_index(sheetno)
            header = []
            mainlist = []
            ci = 0
            rj = 0
            if 'input 1' in fname.lower():
                ci, rj = getused_ranged(ci, rj, sheet)
            else:
                ci, rj = get_used_range(ci, rj, sheet)
            i = rj
            for rowx in range(rj, sheet.nrows):
                values = sheet.row_values(rowx)
                basejson = {}
                if (i == rj):
                    header = values
                else:
                    for x in range(ci, len(values) - 1):
                        basejson[header[x]] = values[x]
                    mainlist.append(basejson)
                i += 1
            with open(fname2 + "_" + str(sheetno) + ".json", 'w', encoding='utf-8') as jsonFile:
                jsonFile.write(json.dumps(list(mainlist), indent=4))
            resfiles.append(fname2+"_"+str(sheetno)+".json")
    else:
        sheet = workbook.sheet_by_index(0)
        ci = 0
        rj = 0
        if 'input 1' in fname.lower():
            ci, rj = getused_ranged(ci, rj, sheet)
        else:
            ci, rj = get_used_range(ci, rj, sheet)
        header = []
        mainlist = []
        i = rj
        for rowx in range(rj, sheet.nrows):
            values = sheet.row_values(rowx)
            basejson = {}
            if (i == rj):
                header = values
            else:
                for x in range(ci, len(values) - 1):
                    basejson[header[x]] = values[x]
                mainlist.append(basejson)
            i += 1

        with open(fname2 + ".json", 'w', encoding='utf-8') as jsonFile:
            jsonFile.write(json.dumps(list(mainlist), indent=4))
        resfiles.append(fname2+".json")

def xlsx_to_xml(fname, fname2,output_folder_path,resfiles):
    workbook = xlrd.open_workbook(fname)
    os.chdir(output_folder_path)
    if (len(workbook.sheet_names()) > 1):
        for sheetno in range(0, len(workbook.sheet_names())):
            sheet = workbook.sheet_by_index(sheetno)
            header = []
            mainlist = []
            ci = 0
            rj = 0

            if 'input 1' in fname.lower():
                ci, rj = getused_ranged(ci, rj, sheet)
            else:
                ci, rj = get_used_range(ci, rj, sheet)
            i = rj
            for rowx in range(rj, sheet.nrows):
                values = sheet.row_values(rowx)
                mainlist.append('<item>\n')
                if (i == rj):
                    header = values
                else:
                    for x in range(ci, len(values) - 1):
                        mainlist.append("<" + str(header[x]).replace(" ", "_") + "> " + str(values[x]).replace(" ",                                                                                                               "_") + " </" + str(
                            header[x]).replace(" ", "_") + ">\n")
                mainlist.append('<\item>\n')
                i += 1
            with open(fname2 + "_" + str(sheetno) + ".xml", 'w', encoding='utf-8') as xmlfile:
                for x in mainlist:
                    xmlfile.write(x)
            resfiles.append(fname2+"_"+str(sheetno)+".xml")
    else:
        sheet = workbook.sheet_by_index(0)
        header = []
        mainlist = []
        ci = 0
        rj = 0

        if 'input 1' in fname.lower():
            ci, rj = getused_ranged(ci, rj, sheet)
        else:
            ci, rj = get_used_range(ci, rj, sheet)
        i = rj
        for rowx in range(rj, sheet.nrows):
            values = sheet.row_values(rowx)
            basejson = {}
            mainlist.append('<item>\n')
            if (i == rj):
                header = values
            else:
                for x in range(ci, len(values) - 1):
                    mainlist.append(
                        "<" + str(header[x]).replace(" ", "_") + "> " + str(values[x]).replace(" ", "_") + " </" + str(
                            header[x]).replace(" ", "_") + ">\n")
            mainlist.append('<\item>\n')
            i += 1

        with open(fname2 + ".xml", 'w', encoding='utf-8') as xmlfile:
            for x in mainlist:
                xmlfile.write(x)
        resfiles.append(fname2+".xml")

def nrtpdfextract(inp):
    pdffile = inp
    doc = fitz.open(pdffile)
    page = doc.loadPage(0)
    pix = page.getPixmap()
    output = "outfile.png"
    pix.writePNG(output)
    text = textract.process(
        output,
        method='tesseract',
        language='eng',
    )

    # with open("../testfolder/outputfolder/"+out,'wb') as outFile:
    #     outFile.write(text)
    os.remove(output)
    return (text)


def func(row):
    xml = ['<item>']
    for field in row.index:
        xml.append('  <{0}>{1}</{0}}>'.format(field, row[field]))
    xml.append('</item>')
    return '\n'.join(xml)


def split_into_sentences(text):
    text = " " + text + "  "
    text = text.replace("\n", " ")
    text = re.sub(prefixes, "\\1<prd>", text)
    text = re.sub(websites, "<prd>\\1", text)
    if "Ph.D" in text: text = text.replace("Ph.D.", "Ph<prd>D<prd>")
    text = re.sub("\s" + alphabets + "[.] ", " \\1<prd> ", text)
    text = re.sub(acronyms + " " + starters, "\\1<stop> \\2", text)
    text = re.sub(alphabets + "[.]" + alphabets + "[.]" + alphabets + "[.]", "\\1<prd>\\2<prd>\\3<prd>", text)
    text = re.sub(alphabets + "[.]" + alphabets + "[.]", "\\1<prd>\\2<prd>", text)
    text = re.sub(" " + suffixes + "[.] " + starters, " \\1<stop> \\2", text)
    text = re.sub(" " + suffixes + "[.]", " \\1<prd>", text)
    text = re.sub(" " + alphabets + "[.]", " \\1<prd>", text)
    if "”" in text: text = text.replace(".”", "”.")
    if "\"" in text: text = text.replace(".\"", "\".")
    if "!" in text: text = text.replace("!\"", "\"!")
    if "?" in text: text = text.replace("?\"", "\"?")
    text = text.replace(".", ".<stop>")
    text = text.replace("?", "?<stop>")
    text = text.replace("!", "!<stop>")
    text = text.replace("<prd>", ".")
    sentences = text.split("<stop>")
    sentences = sentences[:-1]
    sentences = [s.strip() for s in sentences]
    return sentences


def hasImage(par):
    """get all of the images in a paragraph
    :param par: a paragraph object from docx
    :return: a list of r:embed
    """
    ids = []
    id = ""
    root = ET.fromstring(par._p.xml)
    namespace = {
        'a': "http://schemas.openxmlformats.org/drawingml/2006/main", \
        'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships", \
        'wp': "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}

    inlines = root.findall('.//wp:inline', namespace)
    for inline in inlines:
        imgs = inline.findall('.//a:blip', namespace)
        for img in imgs:
            id = img.attrib['{{{0}}}embed'.format(namespace['r'])]
        ids.append(id)

    return ids


def cleannrtdata(data):
    outputlist = []
    i = 0
    for x in data:
        if ("\\r\\n" in x):
            outputlist.append({"text" + str(i + 1): x.replace("\\r\\n", " ")})
            i += 1
        elif ("\\r\\n\\r" in x):
            outputlist.append({"text" + str(i + 1): x.split("\\r\\n\\r\\n", 1)[0]})
            i += 1
            outputlist.append({"text" + str(i + 1): x.split("\\r\\n\\r\\n", 1)[1]})
            i += 1
        else:
            outputlist.append({"text" + str(i + 1): x + "<text" + str(i)})
            i += 1
    return outputlist


def read_docs(path):
    global dc
    dc = Document(path)

    data = []
    para_count = 0

    for line in dc.paragraphs:
        has_im = hasImage(line)
        if len(has_im) > 0:
            dic = {
                "para" + str(para_count): has_im[0],
                "type": "img"
            }
            data.append(dic)
        else:
            hyperlink = line._p.xpath("./w:hyperlink")
            if len(hyperlink) > 0:
                link = " "
                for hyperl in hyperlink:
                    hyperlink_rel_id = hyperl.get(qn("r:id"))
                    if hyperlink_rel_id is None:
                        continue
                    link += dc.part.rels[hyperlink_rel_id]._target + " "
                dic = {
                    "para" + str(para_count): line.text,
                    "type": "paralink",
                    "link": link
                }
                data.append(dic)
            else:
                dic = {
                    "para" + str(para_count): line.text,
                    "type": "para",
                    "style": line.style.name
                }
                data.append(dic)
        para_count += 1

    return data


#########################
##### pdf processing ####
#########################

def pdf_extracter(pdf_path):
    text = extract_text(pdf_path)
    data = []
    count = 0
    hold = ""
    for i in text:
        if i != "\n":
            hold += i
        else:
            count += 1
            data.append({
                "text" + str(count): hold
            })
            hold = ""
    return data


##############################
##### Save Processed File ####
##############################

# saving to json
def save_json(data, file_name):
    with open(file_name, "w") as file:
        json.dump(data, file)
    # print(data)
    return True


# saving to xml
def save_xml(data, file_name):
    # Converting Python Dictionary to XML
    # xml = dict2xml(data)
    xml = dict2xml.dict2xml(data)

    with open(file_name, "wb") as file:
        file.write(xml.encode())
    return True


def is_number(n):
    try:
        float(n)
    except ValueError:
        return False
    return True


def chkcond(text):
    flag = False
    #####################
    mainlist = []
    for i in range(len(mainlist[10]) - 1):
        if (mainlist[10][i].isdigit()):
            if (mainlist[10][i + 1] == "."):
                flag = True
    print(flag)


def findpat(val):
    if (val[0].isdigit() and val[1] == '.'):
        return True


def applyruledoc(data):
    mainlist = []
    contcntr = False
    i = 0
    j = 0
    for k in range(0, len(data) - 1):
        if (str(list(data[k])[0]).find('para') != -1):
            length = len(data[k][list(data[k])[0]].split())
            # print(data[k][list(data[k])[0]])
            if (length > 1):
                start = data[k][list(data[k])[0]].split()[0]
                end = data[k][list(data[k])[0]].split()[-1]
                value = data[k][list(data[k])[0]]
                info = {"type": data[k][list(data[k])[1]], "style": data[k][list(data[k])[2]]}
                firstletter = ''
                lastletter = ''
                for word in start.split():
                    firstletter = word[0]
                for word1 in end.split():
                    lastletter = word1[-1]
                sentlist = split_into_sentences(value)
                for x in sentlist:
                    if (len(x) > 1):
                        if (contcntr == False and findpat(x)):
                            contcntr = True
                            buffer.append(value)
                            infobuffer.append(info)
                        elif (contcntr == True):
                            val = ""
                            for x in buffer:
                                val = val + x
                            val = val + value
                            contcntr = False
                            igi = {"para" + str(j): x, "type": infobuffer[0]["type"], "style": infobuffer[0]["style"]}
                            if (igi not in mainlist):
                                mainlist.append(igi)
                            buffer.clear()
                            infobuffer.clear()
                        else:
                            igi = {"para" + str(j): x, "type": info["type"], "style": info["style"]}
                            if (igi not in mainlist):
                                mainlist.append(igi)
                        j += 1

    return mainlist


def applyrulepdf(data):
    contcntr = False
    mainlist = []
    remlist = []
    output = []
    for k in range(0, len(data) - 1):
        if (str(data[k].keys()).find('text') != -1):
            for i in data[k].keys():
                length = len(data[k].get(i).split())
                if (length > 1):
                    start = data[k].get(i).split()[0]
                    end = data[k].get(i).split()[-1]
                    value = data[k].get(i)
                    firstletter = ''
                    lastletter = ''
                    for word in start.split():
                        firstletter = word[0]
                    for word1 in end.split():
                        lastletter = word1[-1]
                    if (contcntr == False and length > 4 and (
                            firstletter.isupper() or is_number(firstletter)) and lastletter == '.'):
                        mainlist.append(value)
                    elif (contcntr == False and length > 4 and (
                            firstletter.isupper() or is_number(firstletter)) and lastletter != '.'):
                        contcntr = True
                        buffer.append(value)
                    elif (contcntr == True and length > 4 and is_number(firstletter)):
                        contcntr = True
                        val = ""
                        for x in buffer:
                            val = val + x
                        mainlist.append(val)
                        buffer.clear()
                        buffer.append(value)
                    elif (contcntr == True and length >= 4 and lastletter == '.'):
                        contcntr = False
                        buffer.append(value)
                        val = ""
                        for x in buffer:
                            val = val + x
                        mainlist.append(val)
                        buffer.clear()
                    elif (contcntr == True and length > 4 and value.find('.') != -1 and lastletter != '.'):
                        result = re.split("\.\s+", value)
                        for i in range(0, len(result)):
                            if (i == 0):
                                buffer.append(result[i] + '.')
                                val = ""
                                for x in buffer:
                                    val = val + x
                                mainlist.append(val)
                                buffer.clear()
                            else:
                                buffer.append(result[i])
                    elif (contcntr == True and length > 4 and value.find('.', 0,
                                                                         len(value) - 1) != -1 and lastletter == '.'):
                        result = re.split("\.\s+", value)
                        for i in range(0, len(result)):
                            if (i == 0):
                                buffer.append(result[i] + '.')
                                val = ""
                                for x in buffer:
                                    val = val + x
                                mainlist.append(val)
                                buffer.clear()
                            else:
                                mainlist.append(result[i] + '.')
                    elif (contcntr == True and length > 4):
                        contcntr = True
                        buffer.append(value)
                    else:
                        if (data[k] not in remlist):
                            remlist.append(data[k])
                            if (contcntr != True):
                                contcntr = False
                else:
                    if (data[k] not in remlist):
                        remlist.append(data[k])
                        if (contcntr != True):
                            contcntr = False

    for y in remlist:
        data.remove(y)
    i = 1
    for x in mainlist:
        output.append({'text' + str(i): x})
        i += 1
    return output

def xlsxsToxlsxInFilter(inputfilename, outputfile_content, orfilecon):
    workbook = xlrd.open_workbook(inputfilename)
    if(len(workbook.sheet_names())>1):
        for sheetno in range(0, len(workbook.sheet_names())):
            sheet = workbook.sheet_by_index(sheetno)
            ################################
            ci = 0
            rj = 0
            if 'input1' in inputfilename.replace(' ', '').lower():
                ci, rj = getused_ranged(ci, rj, sheet)
            else:
                ci, rj = get_used_range(ci, rj, sheet)
            ###################################

            for key in xlsx_filter_condition[0]:
                isempty = True
                for con in xlsx_filter_condition[0][key]:
                    for cols in range(ci,sheet.ncols):
                        values = sheet.col_values(cols)
                        incon = values[rj].replace(' ','').lower()
                        sucon = con.replace(' ','').lower()
                        if sucon == incon:
                            isempty = False
                            for vals in range(rj+1, len(values)):
                                outputfile_content[0][key].append(str(values[vals]))
                        elif '('+sucon+')' in incon :
                            isempty = False
                            for vals in range(rj+1, len(values)):
                                outputfile_content[0][key].append(str(values[vals]))
                if isempty:
                    for vals in range(rj+1, len(values)):
                           outputfile_content[0][key].append('')
                ##############################
            for rows in range(1, sheet.nrows - rj):
                    orfilecon.append(inputfilename)
    else:
        sheet = workbook.sheet_by_index(0)
        ##############################
        rj = 0
        ci = 0
        if 'input1' in inputfilename.replace(' ','').lower():
            ci, rj = getused_ranged(ci, rj, sheet)
        else:
            ci, rj = get_used_range(ci, rj, sheet)
        ##############################
        for key in xlsx_filter_condition[0]:
            for con in xlsx_filter_condition[0][key]:
                for cols in range(ci, sheet.ncols):
                    values = sheet.col_values(cols)
                    incon = values[rj].replace(' ', '').lower()
                    sucon = con.replace(' ', '').lower()
                    if sucon == incon:
                        isempty = False
                        for vals in range(rj + 1, len(values)):
                            outputfile_content[0][key].append(str(values[vals]))
                    elif '(' + sucon + ')' in incon:
                        isempty = False
                        for vals in range(rj + 1, len(values)):
                            outputfile_content[0][key].append(str(values[vals]))
            if isempty:
                for vals in range(rj + 1, len(values)):
                    outputfile_content[0][key].append('')
                isempty = True
            ##############################
        for rows in range(1, sheet.nrows - rj):
            orfilecon.append(inputfilename)
            ##############################


def get_used_range(ci, rj, sheet):
    for ii in range(sheet.ncols):
        rv = sheet.row_values(ii)
        if '' not in rv:
            break
        else:
            rj += 1
    return ci, rj


def getused_ranged(ci, rj, sheet):
    nr = sheet.nrows
    nc = sheet.ncols
    isbreak = False
    for ii in range(nc):
        if isbreak:
            break
        cv = sheet.col_values(ii)
        for jj in range(nr):
            if '' != cv[jj]:
                ci = ii
                rj = jj
                isbreak = True
                break
    return ci, rj


def write_output_content_to_xlsx(file_content,original,outputfilename):

    file_content[0]['Original File'] = original
    print(file_content[0])
    df = pd.DataFrame(file_content[0])
    writer = pd.ExcelWriter(outputfilename)
    # write the data frame into a excel file
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    # Add a header format.
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1})

    # Write the column headers with the defined format.
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
    writer.save()
    writer.close()
def isAllNanInList(list):
    count = 0
    for v in list:
        if 'nan' == str(v):
           count += 1
    if count == len(list):
        return True
    else:
        return False