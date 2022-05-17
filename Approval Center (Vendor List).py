#!/usr/bin/env python
# coding: utf-8

# In[23]:


#Import Library
import csv
import json
import datetime
import shutil

#Pre-set documents destination
log_des = "C:\Approval Center Invoice\Transaction Log"
csv_des = "C:\Approval Center Invoice\BulkVendorList\\"
org_des = "C:\Approval Center Invoice\\"
move_des = "C:\Approval Center Invoice\Completed(Vendor)"

#Open Specific Documents
data = open( org_des + 'approval_vendor.csv', encoding = "utf-8")
csv_data = csv.reader(data)
data_lines = list(csv_data) #Whole documents

#Initialize the excel
runtime = datetime.datetime.now().strftime("%Y%m%d%H%M") #record running time
file_to_output = open(csv_des + 'VendorList_' + runtime + '.csv', 'w', newline = '')
log_file = open( log_des + '\Vendor_Transaction Log_' + runtime + '.txt', 'w', newline = '')
log_file.write("Vendor Title\tRow Number")
csv_writer = csv.writer(file_to_output, delimiter = ',', quoting=csv.QUOTE_ALL)

#Hard Code Header
csv_writer.writerow(['Vendor Requisition','Vendor Name','Vendor Address',
                     'Existing Vendor no(change info)','Contract Person', 'Telephone Number','Email Address'
                     ,'Payment Term', 'Payment Method','Currency', 'Payment Purpose code',
                     
                     'Bank Name','Bank Address','Bank Code', 
                     'Branch Code', 'Bank Account Number', 'Swift code', 'IBAN number', 'Bank beneficiary name',
                     'Intermediary Bank Name', 'Intermediary Bank Address', 'City (Intermediary bank)', 'Country (Intermediary bank)',
                     'Bank acct number (Intermediary)', 'Swift Code (Intermediary bank)', 'IBAN Number (Intermediary bank)' 
                     ])

CustomTabLoop = []
ExpenseLoop = []
count = 0
count_Vendor = 0
for p in data_lines[1:]: #Skipping Title row
    Title = p[0] #Retrieve Title
    Location = p[1] # Retreieve Location
    count_Vendor += 1
    if p[2] not in '':
        CustomTabLoop.append(p[2]) #put it into a list
        a = CustomTabLoop[count] # select specific list = string
        z = json.loads(a) #convert string to json format
        
        #Verification
        if "Bank name (full name)" in z[0]:
            Bank_name = z[0]["Bank name (full name)"]
        else:
            Bank_name = ''
        #
        if "Bank address" in z[0]:
            Bank_address = z[0]["Bank address"]
        else:
            Bank_address = ''
        #
        if "Bank code" in z[0]:
            Bank_code = z[0]["Bank code"]
        else:
            Bank_code = ''
        #    
        if "Branch code" in z[0]:
            Branch_code = z[0]["Branch code"]
        else:
            Branch_code = ''
        #
        if "Bank account number" in z[0]:
            Bank_account_number = z[0]["Bank account number"]
        else:
            Bank_account_number = ''
        #    
        if "Bank beneficiary name" in z[0]:
            Bank_beneficiary_name = z[0]["Bank beneficiary name"]
        else:
            Bank_beneficiary_name = ''
        #    
        if "Swift code (TT only)" in z[0]: #20220517
            Swift_code = z[0]["Swift code (TT only)"]
        else:
            Swift_code = ''
        #
        if "IBAN number (TT only)" in z[0]: #20220517
            IBAN = z[0]["IBAN number (TT only)"]
        else:
            IBAN = ''
        #Approvals Center Fields - Intermediary Bank 
        if "Intermediary bank name (if any)" in z[0]:
            Inter_Bank_Name = z[0]["Intermediary bank name (if any)"]
        else:
            Inter_Bank_Name = ''
            
        if "Intermediary bank address" in z[0]:
            Inter_Bank_Address = z[0]["Intermediary bank address"]
        else:
            Inter_Bank_Address = ''
            
        if "City (Intermediary bank)" in z[0]:
            City = z[0]["City (Intermediary bank)"]
        else:
            City = ''
            
        if "Country (Intermediary bank)" in z[0]:
            Country = z[0]["Country (Intermediary bank)"]
        else:
            Country = ''
       
        if "Bank acct number (Intermediary)" in z[0]:
            Inter_Bank_account_number = z[0]["Bank acct number (Intermediary)"]
        else:
            Inter_Bank_account_number = ''
            
        if "Swift code (Intermediary bank)" in z[0]:
            Inter_Swift = z[0]["Swift code (Intermediary bank)"]
        else:
            Inter_Swift = ''
            
        if "IBAN number (Intermediary bank)" in z[0]:
            Inter_IBAN = z[0]["IBAN number (Intermediary bank)"]
        else:
            Inter_IBAN = ''
            
            
            
        
        log_file.write('\n' + Title + ' \t' +  str(count_Vendor))     
        
        if p[3] not in '':
            ExpenseLoop.append(p[3]) #put it into a list
            c = ExpenseLoop[count] # select specific list = string
            b = json.loads(c)
            Vendor_Requisition = []
            Vendor_name = []
            Vendor_Address = []
            Change_info = []
            Contract_person = []
            Tel_no = []
            Email = []
            Payment_term = []
            Payment_Method = []
            Currency = []
            Payment_Purpose_Code = []

        if len(b) == 1:
            length = len(b)
        else:
            length = len(b)

        for i in range(0, length):
            #Vendor Requisition
            if "CustomField1" in b[i]:
                #Vendor_Requisition = b[i]["CustomField1"]
                Vendor_Requisition.append(b[i]["CustomField1"])
            else:
                #Vendor_Requisition = ''
                Vendor_Requisition.append('')
            #Vendor Name
            if "CustomField3" in b[i]:
                #Vendor_name = b[i]["CustomField3"]
                Vendor_name.append(b[i]["CustomField3"])
            #elif "CustomField3" not in b[0] and "Vendor name (if New)" in b[0]:
            #    Vendor_name = b[0]["Vendor name (if New)"]
            else:
                #Vendor_name = ''
                Vendor_name.append('')
            #Vendor Address
            if "CustomField4" in b[i]:
                #Vendor_Address = b[i]["CustomField4"]
                Vendor_Address.append(b[i]["CustomField4"])
            else:
                #Vendor_Address = ''
                Vendor_Address.append('')

            #Change Info
            if Vendor_Requisition[i] in ['Change Information'] and "CustomField2" not in ['']:
                Change_info.append(b[i]["CustomField2"].split('_')[0]) #Custom Field 2 = Vendor code
            else:
                #Change_info = ''
                Change_info.append('')
            #
            if "CustomField5" in b[i]:
                #Contract_person = b[i]["CustomField5"]
                Contract_person.append(b[i]["CustomField5"])
            else:
                #Contract_person = ''
                Contract_person.append('')
            #
            if "CustomField6" in b[i]:
                #Tel_no = b[i]["CustomField6"]
                Tel_no.append(b[i]["CustomField6"])
            else:
                #Tel_no = ''
                Tel_no.append('')
            #    
            if "CustomField7" in b[i]:
                #Email = b[i]["CustomField7"]
                Email.append(b[i]["CustomField7"])
            else:
                #Email = ''
                Email.append('')

            #Check Payment Term
            if "CustomField8" in b[i]:
                if b[i]["CustomField8"] in ['Others (please specify)'] and "CustomField9" in b[i]:
                    Payment_term.append(b[i]["CustomField9"])
                else:
                    Payment_term.append(b[i]["CustomField8"])
            else:
                #Payment_term = ''
                Payment_term.append('')
            #    
            if "CustomField10" in b[i]:
                #Payment_Method = b[i]["CustomField10"].split('_')[i]
                Payment_Method.append(b[i]["CustomField10"].split('_')[i])
            else:
                #Payment_Method = ''
                Payment_Method.append('')
            #    
            if "CustomField11" in b[i]:
                #Currency = b[i]["CustomField11"]
                Currency.append(b[i]["CustomField11"])
            else:
                #Currency = ''
                Currency.append('')
            #    
            if "CustomField12" in b[i]:
                #Payment_Purpose_Code = b[i]["CustomField12"]
                Payment_Purpose_Code.append(b[i]["CustomField12"])
            else:
                #Payment_Purpose_Code = ''
                Payment_Purpose_Code.append('')
            
            if Vendor_Requisition[i] not in ['']: #Vendor Requisition
                csv_writer.writerow([Vendor_Requisition[-1], Vendor_name[-1], Vendor_Address[-1], Change_info[-1], Contract_person[-1], Tel_no[-1]
                                , Email[-1], Payment_term[-1], Payment_Method[-1], Currency[-1], Payment_Purpose_Code[-1]
                                , Bank_name, Bank_address, Bank_code , Branch_code, Bank_account_number, Swift_code
                                , IBAN, Bank_beneficiary_name
                                , Inter_Bank_Name, Inter_Bank_Address, City, Country, Inter_Bank_account_number, Inter_Swift, Inter_IBAN
                                , ])
            
    count += 1


    
    

file_to_output.close() #Close File
log_file.close()
data.close()

#Rename Source data file
original_des = org_des + 'approval_vendor.csv'
target_des = 'C:\Approval Center Invoice\Completed\\source_approval_' + Location + '_' + runtime + '.csv'
shutil.move(original_des, target_des)

#Rename Exported data file
org_csv_des = csv_des + 'VendorList_' + runtime + '.csv'
csv_target_des = 'C:\Approval Center Invoice\BulkVendorList\\approval_' + Location + '_' + runtime + '.csv'
shutil.move(org_csv_des, csv_target_des)

#Rename Log File
org_log_des = 'C:\Approval Center Invoice\Transaction Log\Vendor_Transaction Log_' + runtime + '.txt'
log_target_des = 'C:\Approval Center Invoice\Transaction Log\\Log_' + Location + '_' + runtime + '.txt'
shutil.move(org_log_des, log_target_des)

