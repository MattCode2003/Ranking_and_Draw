'''
Created on March 27, 2024
@author: Matthew Pryke
'''

from ranking_collection import RankingCollection
from Groups import GroupCreation

def main() -> None:
    # Ranking Collection
    while True:
        ranking = input("Would you like to get everyones ranking? (y/n): ")
        if ranking.lower() == "y" or ranking.lower() == "yes":
            RankingCollection().main()
            break
        elif ranking.lower() == "n" or ranking.lower() == "no":
            break
        else:
            print("Sorry, I didn't understand")

    # Group Creation
    while True:
        groups = input("Would you like to create the groups? (y/n): ")
        if groups.lower() == "y" or groups.lower() == "yes":
            group_creation = GroupCreation()
            group_creation.main()
            break
        elif groups.lower() == "n" or groups.lower() == "no":
            break
        else:
            print("Sorry, I didn't understand")


if __name__ == "__main__":
    main()