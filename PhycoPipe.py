
## ========================================== ##
##                                            ##
##   Intellectual Property Rights (IPR):      ##
##                                            ##
##   João Paulo Andrade Barbosa (creator)     ##
##   2025-01-07 (date)                        ##
##                                            ##
## ========================================== ##


# ============================ #
#                              #
#      PhycoPipe (v1.0.0)      #
#                              #
# ============================ #

# Lybrary imports


print(f'\nWelcome to the PhycoPipe, a simple toolbox for works with macroalgae communities\n')

## custom functions ##

# Main menu
def tools_menu():
    print(f'\n---- Tools Menu ----')
    print(f'(0) Exit')
    print(f'(1) ')
    print(f'(2) ')
    print(f'(3) ')
    print(f'(4) ')

def tools_menu_loop():    
    while True:
        tools_menu()
        choice = input("\nChoose one of my tools: ").strip()
        
        if  choice == '0':
            print(f'\nExiting...\n')
            break

        elif choice == '1':
            print(f'\n\n-------------  -------------- \n')
            tool_1()
        
        elif choice == '2':
            print(f'\n\n-------------  -------------- \n')
            tool_2()
        
        elif choice == '3':
            print(f'\n\n-------------  -------------- \n')
            tool_3()
        
        elif choice == '4':
            print(f'\n\n-------------  -------------- \n')
            tool_4()
        
        else:
            print(f'\n\nInvalid choice. Please, try again.\n')
