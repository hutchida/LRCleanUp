#Prefabrication clean up script for Lexis Recommends

import pandas as pd
import os
import glob
import re
import xml.etree.ElementTree as ET
from lxml import etree
import time
import codecs
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning) #ingore a particular warning from pandas that is actually a bug

root = '\\\\atlas\\knowhow\\PSL_Content_Management\\Digital Editors\\Lexis_Recommends\\'
xlsdir = root + 'xls\\'
xmldircopy = root + 'xml\\'
xmldir = '\\\\lngoxfclup24va\\glpfab4\\Build\\00WL\\Update\\'
date = time.strftime("%d-%m-%Y")
hms = time.strftime("%H%M%S")
logdir = root + 'log\\'
csvdir = root + 'csv\\'
lookupdpsi = root + 'lookupdpsi\\'
logfilepath = logdir + 'lr-' + date + '-' + hms +'.txt'
file = 'LexisPSL_Recommends_Word_List.xlsx'
filepath = xlsdir + file

#aicerdir = "\\\\atlas\\knowhow\\PSL_Content_Management\\AICER_Reports\\AICER\\"

aicerdir = "C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\"

def MostRecentAicer():
    print("Finding most recent AICER report...this can take up to 20 minutes so be patient...")
    dirjoined = os.path.join(aicerdir, '*AllContentItemsExport_[0-9][0-9][0-9][0-9].csv') # this joins the directory variable with the filenames within it, but limits it to filenames ending in 'AllContentItemsExport_xxxx.csv', note this is hardcoded as 4 digits only. This is not regex but unix shell wildcards, as far as I know there's no way to specifiy multiple unknown amounts of numbers, hence the hardcoding of 4 digits. When the aicer report goes into 5 digits this will need to be modified, should be a few years until then though
    files = sorted(glob.iglob(dirjoined), key=os.path.getmtime, reverse=True) #search directory and add all files to dict
    aicer = files[0]
    return aicer

aicer = MostRecentAicer()
#aicer = 'C:\\Users\\Hutchida\\Documents\\PSL\\AICER\\AllContentItemsExport_1377.csv'
print('Loaded most recent AICER report...' + aicer)

dfaicer = pd.read_csv(aicer, encoding='UTF-8', low_memory=False)   
dfdpsi = pd.read_csv(lookupdpsi + 'lookup-dpsis.csv')
df = pd.ExcelFile(filepath)
df = df.parse("Sheet1") #Select which sheet to parse
dfcsv = pd.DataFrame()
deletetext = '\r\n\r\n...DELETIONS...\r\n\r\nDeleted the following entries because DocId does not exist:\r\n\r\n'
replacedoctitletext = '\r\n\r\n...TITLE REPLACEMENTS...\r\n\r\nReplaced the following DocTitles to match the latest AICER report:\r\n\r\n'
replacedpsitext = '\r\n\r\n...DPSI REPLACEMENTS...\n\nReplaced the following DPSIs with the correct ones based on PA and Content Type and the DPSI lookup table:\r\n\r\n'

#set xml parents
Workbook = ET.Element('Workbook')
Workbook.set('xmlns', 'urn:schemas-microsoft-com:office:spreadsheet')
Worksheet = ET.SubElement(Workbook, 'Worksheet')
Table = ET.SubElement(Worksheet, 'Table')

#set column header
Row = ET.SubElement(Table, 'Row') #add new row in xml
Cell = ET.SubElement(Row, 'Cell')
Data = ET.SubElement(Cell, 'Data')
Data.text = 'Practice Area'
Cell = ET.SubElement(Row, 'Cell')
Data = ET.SubElement(Cell, 'Data')
Data.text = 'Search Term'
Cell = ET.SubElement(Row, 'Cell')
Data = ET.SubElement(Cell, 'Data')
Data.text = '1st Recommended Doc'
Cell = ET.SubElement(Row, 'Cell')
Data = ET.SubElement(Cell, 'Data')
Data.text = 'DPSI'
Cell = ET.SubElement(Row, 'Cell')
Data = ET.SubElement(Cell, 'Data')
Data.text = 'PAM ID'
Cell = ET.SubElement(Row, 'Cell')
Data = ET.SubElement(Cell, 'Data')
Data.text = '2nd Recommended Doc'
Cell = ET.SubElement(Row, 'Cell')
Data = ET.SubElement(Cell, 'Data')
Data.text = 'DPSI'
Cell = ET.SubElement(Row, 'Cell')
Data = ET.SubElement(Cell, 'Data')
Data.text = 'PAM ID'
Cell = ET.SubElement(Row, 'Cell')
Data = ET.SubElement(Cell, 'Data')
Data.text = '3rd Recommended Doc'
Cell = ET.SubElement(Row, 'Cell')
Data = ET.SubElement(Cell, 'Data')
Data.text = 'DPSI'
Cell = ET.SubElement(Row, 'Cell')
Data = ET.SubElement(Cell, 'Data')
Data.text = 'PAM ID'

i=0
rt=1 #counter for replacement titles
rd=1 #counter for replacement dpsis
d=1 #counter for deletions

print('Looping through the LR spreadsheet and comparing with the AICER report...')
for index, row in df.iterrows():
    Row = ET.SubElement(Table, 'Row') #add new row in xml

    PA = df.iloc[i,0]
    
    Cell = ET.SubElement(Row, 'Cell')
    Data = ET.SubElement(Cell, 'Data')
    Data.text = str(PA)

    SearchTerm = df.iloc[i,1]

    Cell = ET.SubElement(Row, 'Cell')
    Data = ET.SubElement(Cell, 'Data')
    Data.text = str(SearchTerm)
    
    Title1 = df.iloc[i,2]
    DPSI1 = df.iloc[i,3]
    DocId1 = df.iloc[i,4]
    Title2 = df.iloc[i,5]
    DPSI2 = df.iloc[i,6]
    DocId2 = df.iloc[i,7]
    Title3 = df.iloc[i,8]
    DPSI3 = df.iloc[i,9]
    DocId3 = df.iloc[i,10]
    
    #1st recommended doc test
    try: #if docID appears in the Aicer
        if any(dfaicer['id'] == int(DocId1)) == False:      
            if DPSI1 != '02O0': #if dpsi is a leg dpsi, doc ID is valid and shouldn't be rejected
                deletetext += '\r\n\r\n' + str(DocId1) + ' ...from 1st recommended doc... ' + 'PA: ' + PA + ', Search term: ' + str(SearchTerm)
                d=d+1
                Title1 = 'nan'
                DPSI1 = 'nan'
                DocId1 = 'nan'
        else: #DocID found on Aicer
            #Compare doc titles with the aicer
            dfaicertitle = dfaicer.loc[dfaicer['id'] == int(DocId1), 'Label'].iloc[0]
            Title1.replace('--', '—')
            if Title1 != dfaicertitle:
                replacedoctitletext += '\r\n\r\n' + str(rt) + ' ...from 1st recommended doc... ' + 'PA: ' + PA + ', Search term: ' + str(SearchTerm) + ', DocID: ' + str(DocId1) + '\r\nOriginal: ' + str(Title1) + '\r\r\nReplacement: ' + str(dfaicertitle)
                Title1 = dfaicertitle #make the change
                rt=rt+1
            #Compare dpsis with the aicer and the lookupdpsi spreadsheet
            dfaicercontenttype = dfaicer.loc[dfaicer['id'] == int(DocId1), 'ContentItemType'].iloc[0]
            if dfaicercontenttype == 'Checklist': dfaicercontenttype = 'PracticeNote' #checklists live in practice note dpsis so here we fool it to think it's actually a PN
            dfaicerpa = dfaicer.loc[dfaicer['id'] == int(DocId1), 'TopicTreeLevel1'].iloc[0]
            dpsi = dfdpsi.loc[(dfdpsi['ContentType'] == dfaicercontenttype) & (dfdpsi['PA'] == dfaicerpa), 'path'].item() #filters dataframe by contenttype and PA then tries to extract the only value under the column of path
            dpsi = re.search('.*\\\\Build\\\\(.*)\\\\Data_RX',dpsi).group(1)
            
            if "synopsis" not in str(DPSI1.lower()):
                if DPSI1 != dpsi:
                    replacedpsitext += '\r\n\r\n' + str(rd) + ' ...from 1st recommended doc... ' + 'PA: ' + PA + ', Search term: ' + str(SearchTerm) + ', DocID: ' + str(DocId1) + '\r\nOriginal: ' + str(DPSI1) + '\r\r\nReplacement: ' + str(dpsi)
                    DPSI1 = dpsi #make the change
                    rd=rd+1
        
                
    except ValueError:
        pass

    #2nd recommended doc test
    try:
        if any(dfaicer['id'] == int(DocId2)) == False:
            if DPSI1 != '02O0': #if dpsi is a leg dpsi, doc ID is valid and shouldn't be rejected
                deletetext += '\r\n\r\n' + str(DocId2) + '...from 2nd recommended doc... ' + 'PA: ' + PA + ', Search term: ' + str(SearchTerm)
                d=d+1
                Title2 = 'nan'
                DPSI2 = 'nan'
                DocId2 = 'nan'
        else:
            dfaicertitle = dfaicer.loc[dfaicer['id'] == int(DocId2), 'Label'].iloc[0]
            Title2.replace('--', '—')
            if Title2 != dfaicertitle:
                replacedoctitletext += '\r\n\r\n' + str(rt) + ' ...from 2nd recommended doc... ' + 'PA: ' + PA + ', Search term: ' + str(SearchTerm) + ', DocID: ' + str(DocId2) + '\r\nOriginal: ' + str(Title2) + '\r\r\nReplacement: ' + str(dfaicertitle)
                Title2 = dfaicertitle #make the change
                rt=rt+1
            
            #Compare dpsis with the aicer and the lookupdpsi spreadsheet            
            dfaicercontenttype = dfaicer.loc[dfaicer['id'] == int(DocId2), 'ContentItemType'].iloc[0]
            if dfaicercontenttype == 'Checklist': dfaicercontenttype = 'PracticeNote' #checklists live in practice note dpsis so here we fool it to think it's actually a PN            
            dfaicerpa = dfaicer.loc[dfaicer['id'] == int(DocId2), 'TopicTreeLevel1'].iloc[0]
            dpsi = dfdpsi.loc[(dfdpsi['ContentType'] == dfaicercontenttype) & (dfdpsi['PA'] == dfaicerpa), 'path'].item() #filters dataframe by contenttype and PA then tries to extract the only value under the column of path
            dpsi = re.search('.*\\\\Build\\\\(.*)\\\\Data_RX',dpsi).group(1)
            
            if "synopsis" not in str(DPSI2.lower()):
                if DPSI2 != dpsi:
                    replacedpsitext += '\r\n\r\n' + str(rd) + ' ...from 2nd recommended doc... ' + 'PA: ' + PA + ', Search term: ' + str(SearchTerm) + ', DocID: ' + str(DocId2) + '\r\nOriginal: ' + str(DPSI2) + '\r\r\nReplacement: ' + str(dpsi)
                    DPSI2 = dpsi #make the change
                    rd=rd+1
        
    except:
        pass
    
    #3rd recommended doc test
    try:
        if any(dfaicer['id'] == int(DocId3)) == False:
            if DPSI1 != '02O0': #if dpsi is a leg dpsi, doc ID is valid and shouldn't be rejected
                deletetext += '\r\n\r\n' + str(DocId3) + '...from 3rd recommended doc... ' + 'PA: ' + PA + ', Search term: ' + str(SearchTerm)
                d=d+1
                Title3 = 'nan'
                DPSI3 = 'nan'
                DocId3 = 'nan'
        else:    
            dfaicertitle = dfaicer.loc[dfaicer['id'] == int(DocId3), 'Label'].iloc[0]            
            Title3.replace('--', '—')
            if Title3 != dfaicertitle:
                replacedoctitletext += '\r\n\r\n' + str(rt) + ' ...from 3rd recommended doc... ' + 'PA: ' + PA + ', Search term: ' + str(SearchTerm) + ', DocID: ' + str(DocId3) + '\r\nOriginal: ' + str(Title3) + '\r\r\nReplacement: ' + str(dfaicertitle)
                Title3 = dfaicertitle #make the change
                rt=rt+1

            #Compare dpsis with the aicer and the lookupdpsi spreadsheet
            
            dfaicercontenttype = dfaicer.loc[dfaicer['id'] == int(DocId3), 'ContentItemType'].iloc[0]
            if dfaicercontenttype == 'Checklist': dfaicercontenttype = 'PracticeNote' #checklists live in practice note dpsis so here we fool it to think it's actually a PN
            dfaicerpa = dfaicer.loc[dfaicer['id'] == int(DocId3), 'TopicTreeLevel1'].iloc[0]
            dpsi = dfdpsi.loc[(dfdpsi['ContentType'] == dfaicercontenttype) & (dfdpsi['PA'] == dfaicerpa), 'path'].item() #filters dataframe by contenttype and PA then tries to extract the only value under the column of path
            dpsi = re.search('.*\\\\Build\\\\(.*)\\\\Data_RX',dpsi).group(1)
            
            if "synopsis" not in str(DPSI3.lower()):
                if DPSI3 != dpsi:
                    replacedpsitext += '\r\n\r\n' + str(rd) + ' ...from 3rd recommended doc... ' + 'PA: ' + PA + ', Search term: ' + str(SearchTerm) + ', DocID: ' + str(DocId3) + '\r\nOriginal: ' + str(DPSI3) + '\r\r\nReplacement: ' + str(dpsi)
                    DPSI3 = dpsi #make the change
                    rd=rd+1
        
    except:
        pass

    #if all values are empty, don't bother writing a row to xml or csv
    if str(Title1) or str(DPSI1) or str(DocId1) or str(Title2) or str(DPSI2) or str(DocId2) or str(Title3) or str(DPSI3) or str(DocId3) != 'nan':   
    #Setting xml after correct values have been found/not found above
        #1st recommended doc elements
        if str(Title1) != 'nan': #if title not empty, create tags with title as the text
            Cell = ET.SubElement(Row, 'Cell')
            Data = ET.SubElement(Cell, 'Data')
            Data.text = str(Title1)
        else: #else title empty so create empty tags
            Cell = ET.SubElement(Row, 'Cell')
            NamedCell = ET.SubElement(Cell, 'NamedCell')
        if str(DPSI1) != 'nan':
            Cell = ET.SubElement(Row, 'Cell')
            Data = ET.SubElement(Cell, 'Data')
            Data.text = str(DPSI1)
        else:
            Cell = ET.SubElement(Row, 'Cell')
            NamedCell = ET.SubElement(Cell, 'NamedCell')
        if str(DocId1) != 'nan':
            Cell = ET.SubElement(Row, 'Cell')
            Data = ET.SubElement(Cell, 'Data')
            Data.text = str(DocId1)
        else:
            Cell = ET.SubElement(Row, 'Cell')
            NamedCell = ET.SubElement(Cell, 'NamedCell')
        
        #2nd recommended doc elements
        if str(Title2) != 'nan':
            Cell = ET.SubElement(Row, 'Cell')
            Data = ET.SubElement(Cell, 'Data')
            Data.text = str(Title2)
        else:
            Cell = ET.SubElement(Row, 'Cell')
            NamedCell = ET.SubElement(Cell, 'NamedCell')
        if str(DPSI2) != 'nan':
            Cell = ET.SubElement(Row, 'Cell')
            Data = ET.SubElement(Cell, 'Data')
            Data.text = str(DPSI2)
        else:
            Cell = ET.SubElement(Row, 'Cell')
            NamedCell = ET.SubElement(Cell, 'NamedCell')
        if str(DocId2) != 'nan':
            Cell = ET.SubElement(Row, 'Cell')
            Data = ET.SubElement(Cell, 'Data')
            Data.text = str(DocId2)
        else:
            Cell = ET.SubElement(Row, 'Cell')
            NamedCell = ET.SubElement(Cell, 'NamedCell')
        
        #3rd recommended doc elements        
        if str(Title3) != 'nan':
            Cell = ET.SubElement(Row, 'Cell')
            Data = ET.SubElement(Cell, 'Data')
            Data.text = str(Title3)
        else:
            Cell = ET.SubElement(Row, 'Cell')
            NamedCell = ET.SubElement(Cell, 'NamedCell')
        if str(DPSI3) != 'nan':
            Cell = ET.SubElement(Row, 'Cell')
            Data = ET.SubElement(Cell, 'Data')
            Data.text = str(DPSI3)
        else:
            Cell = ET.SubElement(Row, 'Cell')
            NamedCell = ET.SubElement(Cell, 'NamedCell')
        if str(DocId3) != 'nan':
            Cell = ET.SubElement(Row, 'Cell')
            Data = ET.SubElement(Cell, 'Data')
            Data.text = str(DocId3)
        else:
            Cell = ET.SubElement(Row, 'Cell')
            NamedCell = ET.SubElement(Cell, 'NamedCell')

        #Export to list
        list1 = [[PA, SearchTerm, Title1, DPSI1, DocId1, Title2, DPSI2, DocId2, Title3, DPSI3, DocId3]]
        #Append to dataframe
        dfcsv = dfcsv.append(list1)  

        i=i+1
    else: 
        print('Row empty, SKIPPING...')

with codecs.open(logfilepath, "w", encoding='utf-8', errors='ignore') as f:
    f.write('Lexis Recommends clean up script log, dated: ' + date + ' ' + hms)
    f.write('\r\n\r\nNumber of entries deleted due to incorrect DocIds: ' + str(d))
    f.write('\r\nNumber of doc titles replaced to reflect AICER report entries: ' + str(rt))
    f.write('\r\nNumber of dpsis replaced: ' + str(rd))
    f.write(str(deletetext))
    f.write(str(replacedoctitletext))
    f.write(str(replacedpsitext))
    f.close()
print('Exported log here...' + logfilepath)
dfcsv.to_csv(csvdir + 'LR.csv', sep=',',index=False, header=["PA", "Search Term", "Title1", "DPSI1", "DocID1", "Title2", "DPSI2", "DocID2", "Title3", "DPSI3", "DocID3"])
print('Exported to csv here...' + csvdir + 'LR.csv')
tree = ET.ElementTree(Workbook)
tree.write(xmldir + 'Recommends Word List.xml',encoding='utf-8')
print('Exported to...' + xmldir + 'Recommends Word List.xml')
tree.write(xmldircopy + 'Recommends Word List.xml',encoding='utf-8')
print('Exported to...' + xmldircopy + 'Recommends Word List.xml')

