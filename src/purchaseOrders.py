'''
Created on 2016/08/23

@author: Brian

This module processes all Purchase Orders in a designated folder.
'''
import openpyxl
import os
import datetime

#Change this to the desired folder path where all purchase orders are stored. Using invoices temporarily for testing.
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
        
def addNewPurchaseOrderObject(newPurchaseOrderObject):
    purchaseOrderList.append(newPurchaseOrderObject)
            
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
    
    print('Information to output to Purchase Order:\n\n')
    print('First Name: ' +  firstName_new)
    print('Last Name: ' + lastName_new)
    print('E-mail: ' + email_new)
    print('Date: ' + date_new)
    print('Department Name: ' + departmentName_new)
    print('PO Number: ' + PONumber_new)
    print('Project Name: ' + itemName_new)
    print('Order Name: ' + orderAmount_New)
    print('Project Manager: ' + projectManager_New)
    print('\n\n')
    
    verifyInfo = input('Would you like to save with this information?(y/n)')
    
    if verifyInfo == 'y':
        
        newPurchaseOrderObject = PurchaseOrder(firstName_new, lastName_new, email_new, date_new, departmentName_new, PONumber_new, itemName_new, orderAmount_New, projectManager_New)
        addNewPurchaseOrderObject(newPurchaseOrderObject)
        
        newFileSaveName = 'PO_' + PONumber_new + '_' + lastName_new + '_' + firstName_new + '.xlsx'
        print('Saving to ' + newFileSaveName + '...')
        purchaseOrderTemplate.save(purchaseOrderDirectory + '\\' + newFileSaveName)
        
    else:
        print('Restarting creation of Purchase Order...')
        createPurchaseOrders()     
    
    
    
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

def deletePurchaseOrderList():
    for item in purchaseOrderList:
        purchaseOrderList.remove(item)
    
#Create list for purchaseOrder Objects
purchaseOrderList = []

def initializePurchaseOrderList():
    
    
    deletePurchaseOrderList()
    
    for root, subfolder, filenames in (os.walk(purchaseOrderDirectory)):
        for file in filenames:  
        #Create new workbook object
            if file.endswith('.xlsx'):
                purchaseOrderWB = openpyxl.load_workbook(purchaseOrderDirectory + '\\' + file)
                purchaseOrderWorksheet = purchaseOrderWB.get_sheet_by_name('Sheet1')
            
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
                
                addNewPurchaseOrderObject(purchaseOrderObject)