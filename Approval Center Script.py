#!/usr/bin/env python
# coding: utf-8

# In[5]:


#Import Library
import csv
import json
import datetime
import shutil

#Pre-set documents destination
log_des = "C:\Approval Center Invoice\Transaction Log"
csv_des = "C:\Approval Center Invoice\BulkInvoiceData\\"
org_des = "C:\Approval Center Invoice\\"
move_des = "C:\Approval Center Invoice\Completed"
#Open Specific Documents
data = open( org_des + 'approval_invoice.csv', encoding = "utf-8")
csv_data = csv.reader(data)
data_lines = list(csv_data) #Whole documents

#Initialize the excel
runtime = datetime.datetime.now().strftime("%Y%m%d%H%M") #record running time
file_to_output = open(csv_des + 'Approval_invoice_' + runtime + '.csv', 'w', newline = '')
log_file = open( log_des + '\Transaction Log_' + runtime + '.txt', 'w', newline = '')
log_file.write("Invoice Title\tRow Number")
csv_writer = csv.writer(file_to_output, delimiter = ',')
#Hard Code Header
csv_writer.writerow(['Vendor number','Vendor Name', 'Invoice Date','Invoice Number', 'Type', 'Dim 1', 'Dim 2', 'Dim 3', 'Account number', 'Purchase item/Description', 'Delivery quantity','Unit price'])
CustomTabLoop = []
ExpenseLoop = []
count = 0
count_invoice = 0
for p in data_lines[1:]: #Skipping Title row
    Title = p[0] #Retrieve Title, Payment template, Good Receving template
    Location = p[1] #Retrieve Location
    count_invoice += 1
    if p[2] not in '': #Remove blank field (CustomTabJSON)
        CustomTabLoop.append(p[2]) #put it into a list
        a = CustomTabLoop[count] # select specific list = string
        z = json.loads(a) #convert string to json format
        code = []
        vendorname = []
        inv_date = [] #define an empty list for invoice date
        inv_no = []
        for i in range(0,len(z)-1):
            #Check Vendor number(Captial Letter)
            if "Existing Vendor number" in z[i]:
                code.append(z[i]["Existing Vendor number"].split('_')[0]) #Extract Vendor Code
                vendorname.append(z[i]["Existing Vendor number"].split('_')[1]) #Extract Vendor Name
            elif "Existing vendor number" in z[i]: #Small letter
                code.append(z[i]["Existing vendor number"].split('_')[0]) #Extract Vendor Code
                vendorname.append(z[i]["Existing vendor number"].split('_')[1]) #Extract Vendor Name
            elif "Vendor number" in z[i]:
                code.append(z[i]["Vendor number"][:5])
            else:
                code.append('--')
            #Invoice Date    
            if "Invoice date" in z[i]:
                inv_date.append(z[i]["Invoice date"])
            else:
                inv_date.append('')
            #Invoice Number    
            if "Invoice number" in z[i]:
                inv_no.append(z[i]["Invoice number"])
            else:
                inv_no.append('')

        #Invoice_Date = z[0]["Invoice date"] #invoice date
        #Invoice_number = z[0]["Invoice number"] #vendor number
        #log_file.write('\n' + Title + ' \t' +  str(count_invoice))


    if p[3] not in '':
        txt = json.loads(p[3])
        Type = []
        Dim_1 = []
        Dim_2 = []
        Dim_3 = []
        Purchase_item_Description = []
        Account_number = []
        Purchase_qty = []
        Unit_Price = []
        #20220311 if only 1 detail record vs >1 records        
        if len(txt) == 1:
            length = len(txt)
        else:
            length = len(txt)
        
        for i in range(0, length):
            #Check Type
            if "Type" in txt[i]: #Type = Good receiving
                Type.append(txt[i]["Type"])
            elif "CustomField12" in txt[i]: # CustomField12 = Payment 
                Type.append(txt[i]["CustomField12"])
            else:
                Type.append('')
            #Check Dim 1    
            if "Dim1" in txt[i]:
                Dim_1.append(txt[i]["Dim1"])
            else:
                Dim_1.append('')
            #Check Dim 2
            if "Dim2" in txt[i]:
                Dim_2.append(txt[i]["Dim2"])
            else:
                Dim_2.append('')
            #Check Dim 3
            if "Dim3 (BPL code)" in txt[i]:
                Dim_3.append(txt[i]["Dim 3 (BPL code)"][0:5])
            else:
                Dim_3.append('')
            #Check Account Number Updated on 20220223: Account Number > GL Account Name    
            if "GLAccountName" in txt[i]: #Good receiving
                Account_number.append(txt[i]["GLAccountName"])
            elif "CustomField4" in txt[i]: #Payment
                Account_number.append(txt[i]["CustomField4"])
            else:
                Account_number.append('')
            #Check Purchase Item/Description
            if "PurchaseItem" in txt[i]: #Good receiving
                Purchase_item_Description.append(txt[i]["PurchaseItem"])
            elif "CustomField5" in txt[i]: #Payment
                Purchase_item_Description.append(txt[i]["CustomField5"])
            else:
                Purchase_item_Description.append('')
            #Check Delivery Qty
            if "DeliveryQuantity" in txt[i]: #Good receiving
                Purchase_qty.append(txt[i]["DeliveryQuantity"])
            elif "CustomField10" in txt[i]: #Payment
                Purchase_qty.append(txt[i]["CustomField10"])
            else:
                Purchase_qty.append('')
            #Check Unit price
            if "UnitPrice" in txt[i]:
                Unit_Price.append(txt[i]["UnitPrice"])
            elif "CustomField7" in txt[i]: #Payment
                Unit_Price.append(txt[i]["CustomField7"])
            else:
                Unit_Price.append('')    

            if Type[i] in ['Opex', 'Capex'] and Purchase_qty[i] not in ['',0]:
                csv_writer.writerow([code[-1],vendorname[-1], inv_date[-1], inv_no[-1], Type[i], Dim_1[i], Dim_2[i], Dim_3[i], Account_number[i][0:7], Purchase_item_Description[i], Purchase_qty[i], Unit_Price[i]])               
                #PS: code[-1], inv_date[-1] means always take last element from a list 20220309
    count += 1
    #z.pop(0) #remove first dictionary
    
file_to_output.close() #Close File
log_file.close()
data.close()


#source File move
original_des = org_des + 'approval_invoice.csv'
target_des = 'C:\Approval Center Invoice\Completed\\source_approval_' + Location + '_' + runtime + '.csv'
shutil.move(original_des, target_des)

#data file move
csv_org_des = csv_des + 'Approval_invoice_' + runtime + '.csv'
target_csv_des = csv_des + 'approval_' + Location + '_' + runtime + '.csv'
shutil.move(csv_org_des, target_csv_des)

#transaction file
log_org_des = log_des + '\Transaction Log_' + runtime + '.txt'
target_log_des = log_des + '\Log_' + Location + runtime + '.txt'
shutil.move(log_org_des, target_log_des)


#if n["Type"] in ['Opex','Capex']: #Check invoice Type
    #print(p[1])
    #Cust = data_lines[1][1] #First row, second coloumn
    #Expense = data_lines[1][2] #First row, third column
    #DetailExpense = json.loads(Expense) #Loading Json Format
    # print(Filename)
    # ----- For Own Reference ----- #
    # ----- print(type(DetailExpense)) #list
    # ----- print(type(DetailExpense[0])) #dictionary
    # ----- For Own Reference ----- #


# In[ ]:




