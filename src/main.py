'''
Created on March 27, 2024
@author: Matthew Pryke
'''


import pyfiglet
import rich
import Groups
import ranking_collection
from utils.text_colour import bcolors



class Main():

    def main(self) -> None:
        t1 = pyfiglet.figlet_format('Referee', font='starwars')
        t2 = pyfiglet.figlet_format('Helper', font='starwars')
        rich.print(t1)
        rich.print(t2)
        self.menu()

    def menu(self) -> None:
        print("")
        print("1: Get Player Ranking")
        print("2: Create Groups")
        print("3: Create desk charts")
        print("4: Exit")

        try:
            user_input = int(input())
        except ValueError:
            print(f"{bcolors.WARNING}Error: Please enter an integer from 1-4{bcolors.ENDC}")
            print(self.menu())

        match user_input:
            case 1:
                ranking_collection.RankingCollection().main()
            case 2:
                Groups.GroupCreation().main()
            case 3:
                print("inop")
            case 4:
                exit(0)


if __name__ == "__main__":
    Main().main()