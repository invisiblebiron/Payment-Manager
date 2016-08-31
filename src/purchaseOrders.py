'''
Created on 2016/08/23

@author: Brian

This module processes all Purchase Orders in a designated folder.
'''
import openpyxl
import os
import datetime

#Change this to the desired folder path where all purhcase orders are stored. Using invoices temporarily for testing.
purchaseOrderDirectory = 'D:\\AGM\\Documents\\Payments\\POs\\Unsent'

class PurchaseOrder:
    
    language = 'N/A'
    
    def __init__(self, firstName, lastName, email, date, departmentName, PONumber, itemName, orderAmount, projectManager):
        
        self.firstName = firstName 
        self.lastName = lastName
        self.email = email
        self.date = date
        self.departmentName = departmentName
        self.PONumber = PONumber
        self.itemName = itemName
        self.orderAmount = orderAmount
        self.projectManager = projectManager

def createPurchaseOrders():
    print("Creating new purchase order...\n")
    #Request necessary info
    firstName_new = input('First Name: ')
    lastName_new = input('Last Name: ')
    email_new = input('E-mail: ')
    date_new = str(datetime.date.today())
    departmentName_new = input('Department: ')
    PONumber_new = input('PO Number: ')
    itemName_new = input('Project Name: ')
    orderAmount_New = input('Order Amount: ')
    projectManager_New = input('Project Manager: ')
    
    #Open template
    templateFolder = 'D:\\AGM\\Documents\\Payments\\Templates'
    templateFile = templateFolder  + '\\' + 'PO Template.xlsx'
    
    purchaseOrderTemplate = openpyxl.load_workbook(templateFile)
    purchaseOrderTemplateSheet = purchaseOrderTemplate.get_sheet_by_name('Sheet1')
    #Assign information
    purchaseOrderTemplateSheet['B3'] = firstName_new
    purchaseOrderTemplateSheet['B4'] = lastName_new
    purchaseOrderTemplateSheet['B5'] = email_new
    purchaseOrderTemplateSheet['I2'] = date_new
    purchaseOrderTemplateSheet['H10'] = departmentName_new
    purchaseOrderTemplateSheet['A14'] = PONumber_new
    purchaseOrderTemplateSheet['B14'] = itemName_new
    purchaseOrderTemplateSheet['J14'] = orderAmount_New 
    purchaseOrderTemplateSheet['H11'] = projectManager_New
    purchaseOrderTemplateSheet['F14'] = PurchaseOrder.language
    
    newFileSaveName = 'PO_' + PONumber_new + '_' + lastName_new + '_' + firstName_new + '.xlsx'
    
    purchaseOrderTemplate.save(purchaseOrderDirectory + '\\' + newFileSaveName)
    
def listPurchaseOrders():
    i = 1
    for purchaseOrder in purchaseOrderList:
        print(str(i) + '. ' + purchaseOrder.firstName + ' ' + purchaseOrder.lastName + ' - ' + purchaseOrder.itemName + ' - ' + str(purchaseOrder.orderAmount))
        i += 1

def getPurchaseOrderTotal():
    total = 0
    
    for purchaseOrder in purchaseOrderList:
        total += purchaseOrder.orderAmount
        
    print('Purchase Order Total: ' + str(total))
        
   

#Create list for purchaseOrder Objects
purchaseOrderList = []

for root, subfolder, filenames in (os.walk(purchaseOrderDirectory)):
    for file in filenames:  
       #Create new workbook object
       if file.endswith('.xlsx'):
            purchaseOrderWB = openpyxl.load_workbook(purchaseOrderDirectory + '\\' + file)
            purchaseOrderWorksheet = purchaseOrderWB.get_sheet_by_name('Sheet1')
        
             #Edit these before using Purchase Orders. Currently set to invoice values.
            firstName = purchaseOrderWorksheet['B3'].value
            lastName = purchaseOrderWorksheet['B4'].value
            email = purchaseOrderWorksheet['B5'].value
            date = purchaseOrderWorksheet['I2'].value
            departmentName = purchaseOrderWorksheet['H10'].value
            PONumber = purchaseOrderWorksheet['A14'].value
            itemName = purchaseOrderWorksheet['B14'].value
            orderAmount = purchaseOrderWorksheet['J14'].value
            projectManager = purchaseOrderWorksheet['H11'].value
            
            #Instantiate Purchase Order Object
            purchaseOrderObject = PurchaseOrder(firstName, lastName, email, date, departmentName, PONumber, itemName, orderAmount, projectManager)
        
            #Append to list for later.
            purchaseOrderList.append(purchaseOrderObject)
        