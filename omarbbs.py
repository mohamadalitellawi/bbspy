import helpercad

class App:
    def __init__(self) -> None:
        pass

    def startup(self):        
        # print the greeting at startup
        self.greeting()
        print()

    def greeting(self):
        print("-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-")
        print("~~~~~~ Welcome to Omar BBS App (by Abo Akram)! ~~~~~~")
        print("-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-")
        print()


    def menu_header(self):
        print("--------------------------------")
        print("Please make a selection:")
        print("(M): repeat this Menu")

        print("(A): Add Bar info and Shape")
        print("(B): Create / Update Excel BBS")

        print("(S): Apply Scale to bar info and shape")
        print("(H): Get Object Handle")
        print("(E): Export Image from Autocad Selection")

        print("(X): eXit program")

    def menu_error(self):
        print("That's not a valid selection. Please try again.")

    def goodbye(self):
        print("-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~")
        print(f"-~-~-~ Thanks for using Omar BBS! ~-~-~-~")
        print("-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~")


    def run(self):
        # Execute the startup routine - ask for name, print greeting, etc
        self.startup()
        # Start the main program menu and run until the user exits
        self.menu()



    def menu(self):
        self.menu_header()

        # get the user's selection and act on it. This loop will
        # run until the user exits the app
        selection = ""
        while (True):
            selection = input("Selection? ")

            if len(selection) == 0:
                self.menu_error()
                continue

            selection = selection.capitalize()
            if selection[0] == 'X':
                self.goodbye()
                break
            elif selection[0] == 'M':
                self.menu_header()
                continue


            elif selection[0] == 'E':
                try:
                    imagefilename = input("Enter Image File Name with path: ")
                    if len(imagefilename) == 0: continue
                    if imagefilename.capitalize()[0] == "X": continue

                    helpercad.export_image(helpercad.get_cad_active_doc(),imagefilename)
                    
                except Exception as e:
                    self.menu_error()
                    raise e

                continue







if __name__ == "__main__":
    app = App()
    app.run()