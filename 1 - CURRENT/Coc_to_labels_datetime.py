"""
Created on Fri Sep 09 14:35:31 2016

@author: alex.messina
"""


if __name__=='__main__':
    try:

        ## Take a generic Excel File of Chain of Custody (CoC) and generate the appropriate labels
        ## Basic modules
        import pandas as pd
        from mailmerge import MailMerge #https://github.com/Bouke/docx-mailmerge
        from requests import get
        import time
    
        
        
        ###############################################
        ################################## ############
        print('######################################')
        print('#### CHAIN OF CUSTODY LABEL-MAKER ####')
        print('####    written by Alex Messina   ####')
        print('######################################')
        print()
        print('This program will generate labels from a Chain of Custody Form (filled out in the proper Excel template)')
        print('If you are connected to the internet you can use a label template stored online...')
        print('...or you can use your own stored locally')
        print()
        user_input = input('Type the name of the Excel file, or drag and drop the Chain of Custody here: \n Example: C:/Users/alex.messina/Documents/Executable/CoC MS4 C034.xlsx \n \n' ).replace("\\", "/").strip('"').strip("r'") # uncomment -.strip("r'")- for use in spyder testing
        print
        print('Chain of Custody file selected: ')
        print(user_input)
        print()
        
        ## Set directory where the Chain of Custody is:
        #maindir = "C:/Users/alex.messina/Documents/GitHub/CoC-labels/Executable/"
        maindir = '/'.join(user_input.split("/")[:-1]) + '/'
        ## Input the name of the excel file that is the Chain of Custody form
        Coc_Excel_File = user_input.split("/")[-1]
        file_input = maindir+Coc_Excel_File
        CoC = pd.ExcelFile(maindir+Coc_Excel_File)
        
        ## How many extra, blank labels (Project Name only) do you want?
        #extra_labels=4
        print_extra_labels= str(input('Want to fill your last sheet with extra labels? \n These extras will only have the project name printed \n Y or n? \n \n' ))
        print ()
        
        if print_extra_labels == 'N' or print_extra_labels == 'n' or print_extra_labels == 'No' or print_extra_labels == 'no':
            extra_labels = 0
        
        
        ###############################################
        
        
        ###### Open template document
        git_template = str(input('Would you like to use the Default Avery 5523 template from the web (github)? Type Y or n \n'))  
        if git_template == 'Y' or git_template == 'y':
            try:
                ## On GitHub
                ## Get template docx file from GitHub and write to folder
                #git_url = "https://raw.github.com/5025-Main/Labelmaker/7610ab5a9dd00ee1465c491abd785aabc02c62d2/Avery%20Label%20Templates/Avery_5523_template_datetime_bigfont.docx" ## URL of template file
                git_url = "https://github.com/5025-Main/Labelmaker/raw/master/Avery%20Label%20Templates/Avery_5523_template_datetime_bigfont.docx" ## URL of template file
                
                r = get(git_url) ## Access the file at the url
                f = open(maindir+'default_git_template.docx','wb') ## Open a new file to download data into
                f.write(r.content) ## Put downloaded data into the file
                f.close() ## Close the file
                label_template_file = maindir+'default_git_template.docx'
            except:
                print()
                print("Couldn't get GitHub template. Try another template?")
                input()
        elif git_template == 'n' or git_template == 'N':
            try:
                label_template_file = input('Type the name of the Label template file, or drag and drop the Template file here: \n Example: C:/Users/alex.messina/Documents/Executable/avery_5523_template.docx \n \n' ).replace("\\", "/").strip('"').strip("r'") # uncomment -.strip("r'")- for use in spyder testing
            except:
                print()
                print("Couldn't use that template. Try again....")
                input()
        else:
            print ()
            print("Didn't get that. Try again...")
            input()

        
        template_dir = '/'.join(label_template_file.split("/")[:-1]) + '/'
        template_file = label_template_file.split("/")[-1]
    
        
        ## Info for All Chain sheets in workbook (don't iteratively write over!)
        label_export_sheet = pd.DataFrame(columns=['ProjectName','ProjectNumber','SampleID','Container','Preservative','BottleNumber','NumberOfBottles','AnalysisSuite','SampleType','Matrix','Sample Date','Sample Time'],dtype=object)
        ### Iterate over sheets in workbook (one workbook per project)
        for sheet in [s for s in CoC.sheet_names if s!= 'ESRI_MAPINFO_SHEET']:
            ## Extract project information and other data
            project_info = CoC.parse(sheet_name=sheet,header=8,parse_cols='A:F')
            ## Project Name, Number, Sample Matrix
            project_name = project_info.loc[0,'Project Name:']
            project_number = project_info.loc[0,'Project Number:']
            sample_matrix = project_info.loc[0,'Sample Matrix:']
            
            ## Extract Label Data like Sample ID, etc.
            form = CoC.parse(sheet_name=sheet,header=10,index_col=0,parse_cols='A:H')#,skip_footer=11)
            ## Drop blank lines from the form
            form = form.drop(form.loc['Special Instructions/Comments:':].index)
            form = form[pd.notnull(form.index)]
            
            print ()
            print ()
            print ('Generating labels for '+sheet+' Chain of Custody for '+project_name)
            
            
            ### CREATE LABEL EXPORT
            ## Iterate over the rows in the CoC to generate labels
            
            for row in form.iterrows():
                ## Indexed by the "SampleID"
                sampleID = row[0]
                print(sampleID)
                ## Label data is in the rest of the row
                info = row[1]
                if len(info['Analysis']) >= 140:
                    print(len(info['Analysis']) )
                    info['Analysis'] = 'See attached'
                else:
                    pass
                
             
                try:
                    sample_date = row[1]['Sample Date'].date().strftime('%m/%d/%Y')
                except:
                    sample_date=''
                #print('Sample Time '+str(row[1]['Sample Time'])
                #print('Sample Time '+str(row[1]['Sample Time']).replace('.0','')
                sample_time = str(row[1]['Sample Time']).replace('.0','')
                if sample_time == 'nan':
                    sample_time = ''
                if len(sample_time) == 1:
                    sample_time = '000'+sample_time
                if len(sample_time) == 2:
                    sample_time = '00'+sample_time
                if len(sample_time) == 3:
                    sample_time = '0'+sample_time
                try:
                    sample_time = pd.to_datetime(sample_time).strftime('%H:%M')
                except ValueError:
                    pass
                ## Determine how many labels are needed, based on how many bottles are needed
                try:
                    no_of_bottles = int(info['No. of Bottles'])
                except:
                    no_of_bottles = 0
                print("%.0f"%no_of_bottles + ' bottles')
                print ()
                
                ## For composite samples, if no. of bottles is blank, print 4 blanks
                ### BLANK COMPOSITE
                if 'comp' in info['Sample Type'] or 'Comp' in info['Sample Type'] or 'Composite' in info['Sample Type'] or 'composite' in info['Sample Type']:
                    if no_of_bottles == 0:
                        print('composite sample, make 4 labels just in case')
                        no_of_bottles = 4 ## can be set to anything
                        ## Generate required number of labels
                        for bottle_num in range(1,no_of_bottles+1):
                            ## Format the info for each label
                            new_label_info = pd.DataFrame({'ProjectName':project_name,'ProjectNumber':project_number,'SampleID':sampleID,'Container':info['Container'],'Preservative':info['Pres'],'BottleNumber':'__','NumberOfBottles':'__','AnalysisSuite':info['Analysis'],'SampleType':info['Sample Type'],'Matrix':sample_matrix,'Sample Date':sample_date,'Sample Time':sample_time},index=[sampleID+'_'+str(bottle_num)],dtype=object)
                            ## Append label info
                            label_export_sheet = label_export_sheet.append(new_label_info)
                ### NUMBERED COMPOSITE BOTTLES
                    elif no_of_bottles >=1:
                        ## Generate required number of labels
                        for bottle_num in range(1,no_of_bottles+1):
                            ## Format the info for each label
                            new_label_info = pd.DataFrame({'ProjectName':project_name,'ProjectNumber':project_number,'SampleID':sampleID,'Container':info['Container'],'Preservative':info['Pres'],'BottleNumber':str(bottle_num),'NumberOfBottles':str(no_of_bottles),'AnalysisSuite':info['Analysis'],'SampleType':info['Sample Type'],'Matrix':sample_matrix,'Sample Date':sample_date,'Sample Time':sample_time},index=[sampleID+'_'+str(bottle_num)],dtype=object)
                            ## Append label info
                            label_export_sheet = label_export_sheet.append(new_label_info)
                        
                ## For everything else go by the no. of bottles in the chain
                elif 'comp' not in info['Sample Type']:
                    ## Generate required number of labels
                    for bottle_num in range(1,no_of_bottles+1):
                        ## Format the info for each label
                        new_label_info = pd.DataFrame({'ProjectName':project_name,'ProjectNumber':project_number,'SampleID':sampleID,'Container':info['Container'],'Preservative':info['Pres'],'BottleNumber':str(bottle_num),'NumberOfBottles':str(no_of_bottles),'AnalysisSuite':info['Analysis'],'SampleType':info['Sample Type'],'Matrix':sample_matrix,'Sample Date':sample_date,'Sample Time':sample_time},index=[sampleID+'_'+str(bottle_num)],dtype=object)
                        ## Append label info
                        label_export_sheet = label_export_sheet.append(new_label_info)
            
        
        ### Reorder label export sheet columns (used to line up withe Excel Macro, want to keep it just in case)
        label_export_sheet =label_export_sheet[['ProjectName','ProjectNumber','SampleID','Container','Preservative','BottleNumber','NumberOfBottles','AnalysisSuite','SampleType','Matrix','Sample Date','Sample Time']]
        
        #########################################
        #### Mail Merge HERE
                
        document = MailMerge(template_dir+template_file) ## Opens template file downloaded from GitHub above
    
        ## Check which merge fields are present, this is where the label data goes into
        #print(document.get_merge_fields()
        get_fields = document.get_merge_fields()
    
        ## Fields that are required for the label:
        merge_fields = label_export_sheet[['NumberOfBottles','Container','ProjectName','AnalysisSuite','SampleID','BottleNumber','Preservative','Sample Date','Sample Time']]
        
        
        ## build list of data for each label
        row_list = []
        for row in merge_fields.iterrows():
            row_list.append({'NumberOfBottles':row[1]['NumberOfBottles'],'Container':row[1]['Container'],'ProjectName':row[1]['ProjectName'],'AnalysisSuite':row[1]['AnalysisSuite'],'SampleID':row[1]['SampleID'],'BottleNumber':row[1]['BottleNumber'],'Preservative':row[1]['Preservative'],'Sample_Date':row[1]['Sample Date'],'Sample_Time':row[1]['Sample Time']})
            
            
            
        ## Print some extra labels with just the project name
        ## Number is determined by how many to fill a sheet of 10 labels
        if print_extra_labels == 'Y' or print_extra_labels == 'y' or print_extra_labels == 'Yes' or print_extra_labels == 'yes':
            numlabels = len(merge_fields)
            extra_labels = 10 - numlabels%10
            if extra_labels == 10:
                extra_labels = 0
            print('Will fill with '+str(extra_labels ) + ' extra labels')
            
            for row in range(extra_labels):    
                row_list.append({'NumberOfBottles':'__','Container':'','ProjectName':project_name,'AnalysisSuite':'','SampleID':'','BottleNumber':'__','Preservative':''})
            
        #print (row_list
        
        ## Do the Mail Merge
        print ()
        print('Merging fields')
        document.merge_rows('NumberOfBottles',row_list)
        
        ## Write the file
        print
        print('Writing labels to file')
        document.write(maindir+Coc_Excel_File+'-labels_output.docx')
        
        ## Number of Labels
        number_of_labels = len(merge_fields) + extra_labels
        print()
        print()
        print('Wrote '+str(number_of_labels)+' labels: '+str(len(merge_fields))+' project labels, and '+str(extra_labels)+' extra labels')
        print()
        print('Good luck sampling!!')
        print('press any key to exit....')
        input()
        
        
    except:
        raise
        time.sleep(60)
        input()
    
