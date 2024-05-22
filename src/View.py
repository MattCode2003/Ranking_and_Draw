import pyfiglet
import rich
from utils.text_colour import bcolors


class View:
    def __init__(self):
        t1 = pyfiglet.figlet_format('Referee', font='starwars')
        t2 = pyfiglet.figlet_format('Helper', font='starwars')
        rich.print(t1)
        rich.print(t2)



    def main_menu(self) -> int:
        while True:
            print("\n Main Menu")
            print("----------------------")
            print("[1] Get Rankings")
            print("[2] Create Groups")
            print("[3] Exit")

            try:
                user_input = int(input("> "))
            except ValueError:
                print("Error: Please Select a number from 1-3")

            if 1 <= user_input <= 3:
                break

        return user_input



#====================================================== Rankings =======================================================
    def ranking_menu(self) -> int:
        while True:
            print("\n Ranking")
            print("----------------------")
            print("[1] Single Player Ranking")
            print("[2] Entire Tournament")
            print("[3] Back")

            try:
                user_input = int(input("> "))
            except ValueError:
                print("Error: Please Select a number from 1-3")

            if 1 <= user_input <= 3:
                break

        return user_input



    def single_ranking_id(self) -> str:
        print("\nEnter Player Licence")
        user_input = str(input("> "))
        return user_input


    def output_single_player_ranking(self, rankings) -> None:
        if rankings != []:
            for ranking in rankings:
                print(f"{ranking[0]}: {ranking[1]}")
        else:
            print(f"{bcolors.WARNING}Error: There is no player with that licence number.{bcolors.ENDC}")



#======================================================= Groups ========================================================



    def competition_name(self) -> str:
        print("")
        while True:
            print("Enter Competition Name")
            user_input = str(input("> "))

            if user_input != "":
                break

        return user_input



    def date(self, event) -> str:
        while True:
            print(f"What date is the {event} taking place? (DD/MM/YYYY)")
            user_input = str(input("> "))

            if user_input != "":
                break
        return user_input



    def group_size(self) -> int:

        user_input = 0
        while True:
            print("Enter Preferred Group Size")
            try:
                user_input = int(input("> "))
            except ValueError:
                self.error_messages("Error: Please Select a Number")
                continue

            if user_input != 0:
                break

        return user_input



    def change_group_size(self) -> str:
        while True:
            print("It is not possible to have the same number of people in each group. Would you like to go above the preferred group size? (y/n)")
            user_input = str(input("> ")).lower()

            if user_input == "y" or user_input == "n" or user_input == "yes" or user_input == "no":
                break
            else:
                self.error_messages("Error: Please Select yes or no")

        return user_input



#=================================================== Error Messages ====================================================



    def message(self, message: str) -> None:
        print(message)



    def error_messages(self, message: str) -> None:
        print(f"{bcolors.WARNING}{message}{bcolors.ENDC}")
