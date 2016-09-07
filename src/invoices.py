'''
Created on 2016/08/23

@author: Brian
'''
import os

#invoiceFolderPath = 'C:\\Users\\Ottish\\Documents\\Invoices'
invoiceFolderPath = 'D:\\AGM\\Documents\\Payments\\Invoices'

def displayInvoices():
    print("Listing invoices...\n")
    count = 1
    for file in os.listdir(invoiceFolderPath):
        if os.path.isfile(invoiceFolderPath + '\\' + file):
            print(str(count) +  '. ' + file)
        count += 1

def getInvoiceFileAmount():
    invoiceList = []
    for file in os.listdir(invoiceFolderPath):
        if file.endswith('.xlsx'):
                invoiceList.append(file)
    
    print('\t\t\t\t' + str(len(invoiceList)) + ' Invoices Available')
    
