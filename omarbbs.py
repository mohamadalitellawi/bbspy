import helpercad
import buisness

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

        print("(A): Add bar info and shape")
        print("(B): create / update excel BBS")
        print("(C): Check bar bbs in autocad only")
        print("(S): apply Scale to bar info and shape")
        print("(E): Export image from autocad selection")

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

                    active_doc = helpercad.get_cad_active_doc()
                    if active_doc is not None:
                        helpercad.export_image(active_doc,imagefilename)
                    
                except Exception as e:
                    self.menu_error()
                    #raise e
                continue


            elif selection[0] == 'S':
                try:
                    scale = input("Enter Shape Block Scale Factor: ")
                    if len(scale) == 0: continue
                    if scale.capitalize()[0] == "X": continue
                    scale = float(scale)
                    buisness.SHAPE_SCALE_FACTOR = scale
                except Exception as e:
                    self.menu_error()
                    #raise e
                continue

            elif selection[0] == 'A':
                try:
                    buisness.link_Bar_Info()
                except Exception as e:
                    self.menu_error()
                    #raise e
                continue


            elif selection[0] == 'B':
                try:
                    buisness.send_selectedbars_to_excel()
                except Exception as e:
                    self.menu_error()
                    #raise e
                continue

            elif selection[0] == 'C':
                try:
                    buisness.check_bbs()
                except Exception as e:
                    self.menu_error()
                    #raise e
                continue



if __name__ == "__main__":
    app = App()
    app.run()