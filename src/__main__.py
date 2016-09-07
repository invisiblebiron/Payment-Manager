'''
Created on 2016/08/22

@author: Brian
'''


import purchaseOrders #Handles all Purchase Order processes.
import invoices #Handles all Invoice processes.

def main():
#Display options for the user.
    def displayOption():
        purchaseOrders.initializePurchaseOrderList() 
        #Retrieve current invoice amount.
        invoices.getInvoiceFileAmount()
        
        #Display all available options for the user.
        print('Choose an option:')
        print('1. Display Purchase Orders')
        print('2. Display Invoices')
        print('3. Create Purchase Order')
        
        #Enter a number which will run a particular function.
        option = input('\nChoice: ')
        print("You've enterted option " + option + ".\n")
        
        #Displays a list of all Purchase Orders (Names, Amounts, Project Name)
        if option == '1':
            purchaseOrders.listPurchaseOrders()
        
        #Displays a list of all Invoices (Names, Amounts, Project Name)    
        elif option == '2':
            invoices.displayInvoices()
        
        #Create a new Purchase Order.   
        elif option == '3':
            purchaseOrders.createPurchaseOrders()
        else:
            print('Not an available option!')
            print('Returing to selection...')
            displayOption()
    
    print('Project Manager Module')
    print('======================\n')
    
    
    
    displaySelection = True
    
    while displaySelection != False: 
        displayOption()
               
        displayChoice = input('Return to selection? Enter y or n: ')
        
        if displayChoice.lower() == 'y':
            displaySelection = True
           
            
        elif displayChoice.lower() == 'n':
            displaySelection = False
            print('Program terminated.')
        
        
           
            
        print('\n\n')
        
if __name__ == '__main__':
    main()



