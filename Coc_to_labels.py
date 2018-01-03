# -*- coding: utf-8 -*-
"""
Created on Fri Sep 09 14:35:31 2016

@author: alex.messina
"""

## Take a generic Excel File of Chain of Custody (CoC) and generate the appropriate labels

## Basic modules
import pandas as pd
from mailmerge import MailMerge


###############################################
##### CHANGE THESE FOR EACH COC ############

## Set directory where the Chain of Custody is:
maindir = "C:/Users/alex.messina/Documents/GitHub/CoC-labels/Bridge/"
## Input the name of the excel file that is the Chain of Custody form
Coc_Excel_File = 'CoC Pier E4 E5 Pre Sediment.xlsx'
## How many extra, blank labels (Project Name only) do you want?
extra_labels=8

###############################################
CoC = pd.ExcelFile(maindir+Coc_Excel_File)

#%%

## Info for All Chain sheets in workbook (don't iteratively write over!)
label_export_sheet = pd.DataFrame(columns=['ProjectName','ProjectNumber','SampleID','Container','Preservative','BottleNumber','NumberOfBottles','AnalysisSuite','SampleType','Matrix'])

### Iterate over sheets in workbook (one workbook per project)
for sheet in CoC.sheet_names:
    ## Extract project information and other data
    project_info = CoC.parse(sheetname=sheet,header=8,parse_cols='A:F')
    ## Project Name, Number, Sample Matrix
    project_name = project_info.ix[0]['Project Name:']
    project_number = project_info.ix[0]['Project Number:']
    sample_matrix = project_info.ix[0]['Sample Matrix:']
    
    ## Extract Label Data like Sample ID, etc.
    form = CoC.parse(sheetname=sheet,header=10,index_col=0,parse_cols='A:H',skip_footer=12)
    ## Drop blank lines from the form
    form = form[pd.notnull(form.index)]
    
    print 
    print
    print 'Generating labels for '+sheet+' Chain of Custody for '+project_name
    
    #%%
    
    ### CREATE LABEL EXPORT
    ## Iterate over the rows in the CoC to generate labels
    for row in form.iterrows():
        ## Indexed by the "SampleID"
        sampleID = row[0]
        print sampleID
        ## Label data is in the rest of the row
        info = row[1]
        ## Determine how many labels are needed, based on how many bottles are needed
        no_of_bottles = int(info['No. of Bottles'])
        print "%.0f"%no_of_bottles + ' bottles'
        print
        
        ## For composite samples no matter what the no. of bottles, print 4 blanks
        if 'comp' in info['Sample Type'] or 'Comp' in info['Sample Type']:
            #print 'composite sample, make 4 labels just in case'
            no_of_bottles = 4 ## can be set to anything
            ## Generate required number of labels
            for bottle_num in range(1,no_of_bottles+1):
                ## Format the info for each label
                new_label_info = pd.DataFrame({'ProjectName':project_name,'ProjectNumber':project_number,'SampleID':sampleID,'Container':info['Container'],'Preservative':info['Pres'],'BottleNumber':'__','NumberOfBottles':'__','AnalysisSuite':info['Analysis'],'SampleType':info['Sample Type'],'Matrix':sample_matrix},index=[sampleID+'_'+str(bottle_num)])
                ## Append label info
                label_export_sheet = label_export_sheet.append(new_label_info)
    
                
        ## For everything else go by the no. of bottles in the chain
        elif 'comp' not in info['Sample Type']:
            ## Generate required number of labels
            for bottle_num in range(1,no_of_bottles+1):
                ## Format the info for each label
                new_label_info = pd.DataFrame({'ProjectName':project_name,'ProjectNumber':project_number,'SampleID':sampleID,'Container':info['Container'],'Preservative':info['Pres'],'BottleNumber':str(bottle_num),'NumberOfBottles':str(no_of_bottles),'AnalysisSuite':info['Analysis'],'SampleType':info['Sample Type'],'Matrix':sample_matrix},index=[sampleID+'_'+str(bottle_num)])
                ## Append label info
                label_export_sheet = label_export_sheet.append(new_label_info)
    

### Reorder label export sheet columns (used to line up withe Excel Macro, want to keep it just in case)
label_export_sheet =label_export_sheet[['ProjectName','ProjectNumber','SampleID','Container','Preservative','BottleNumber','NumberOfBottles','AnalysisSuite','SampleType','Matrix']]


#%%

###### Open template document

## In folder
#document = MailMerge(maindir+'Messina template - Avery 5523 2x4.docx')

## On GitHub
## Get template docx file from GitHub and write to folder
import requests
git_url = "https://raw.githubusercontent.com/CaptainAL/CoC-labels/master/Avery_5523_template.docx" ## URL of template file
r = requests.get(git_url) ## Access the file at the url
f = open(maindir+'git_template.docx','wb') ## Open a new file to download data into
f.write(r.content) ## Put downloaded data into the file
f.close() ## Close the file
document = MailMerge(maindir+'git_template.docx') ## Opens template file downloaded from GitHub above


## Check which merge fields are present, this is where the label data goes into
print document.get_merge_fields()
get_fields = document.get_merge_fields()

## Fields that are required for the label:
merge_fields = label_export_sheet[['NumberOfBottles','Container','ProjectName','AnalysisSuite','SampleID','BottleNumber','Preservative']]


## build list of data for each label
row_list = []
for row in merge_fields.iterrows():
    row_list.append({'NumberOfBottles':row[1]['NumberOfBottles'],'Container':row[1]['Container'],'ProjectName':row[1]['ProjectName'],'AnalysisSuite':row[1]['AnalysisSuite'],'SampleID':row[1]['SampleID'],'BottleNumber':row[1]['BottleNumber'],'Preservative':row[1]['Preservative']})
    
    
    
## Print some extra labels with just the project name
## Number is set up at the top of this code
for row in range(extra_labels):    
    row_list.append({'NumberOfBottles':'__','Container':'','ProjectName':project_name,'AnalysisSuite':'','SampleID':'','BottleNumber':'__','Preservative':''})
    
## Do the Mail Merge
print 'Merging fields'
document.merge_rows('NumberOfBottles',row_list)

## Write the file
print 'Writing labels'
document.write(maindir+Coc_Excel_File+'-labels_output.docx')

## Number of Labels
number_of_labels = len(merge_fields) + extra_labels
print 
print
print 'Wrote '+str(number_of_labels)+' labels: '+str(len(merge_fields))+' project labels, and '+str(extra_labels)+' extra labels'




#%%


