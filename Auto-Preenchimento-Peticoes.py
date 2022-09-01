from functions.aux_functions import createDesktopShortcut
from classes.autopreenchimentogui import AutoPreenchimentoGUI


def main():
    createDesktopShortcut()
    programGUI = AutoPreenchimentoGUI()
    programGUI.initGUI()



if __name__ == "__main__":
    main()
