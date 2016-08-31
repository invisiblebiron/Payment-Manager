'''
Created on 2016/08/22

@author: Brian
'''
import purchaseOrders
import invoices

def main():
#Display options for the user.
    def displayOption():
        invoices.getInvoiceFileAmount()
        print('Choose an option:')
        print('1. Display Purchase Orders')
        print('2. Display Invoices')
        print('3. Create Purchase Order')
        
        option = input('\nChoice: ')
        
        if option == '1':
            purchaseOrders.listPurchaseOrders()
            
        elif option == '2':
            invoices.getInvoiceFileAmount()
            
        elif option == '3':
            purchaseOrders.createPurchaseOrders()
        else:
            print('Not an available option!')
    
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



