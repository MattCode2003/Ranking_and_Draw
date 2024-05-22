from utils.text_colour import bcolors
import time
import os

class Model:
    def __init__(self):
        self.player_list = "input files/player_list.xlsx"
        self.alpha_list = ""
        self.players = []
        self.competition_name = ""
        self.short_date = ""
        self.long_date = ""
        self.groups = []
        self.forward = True
        self.end = 1
        self.player_numbers = False
        self.row = 1

        input_files_directory = str(os.path.join(os.getcwd(), "input files"))

        for filename in os.listdir(input_files_directory):
            if "Alpha".lower() in filename.lower():
                self.set_alpha_list(f"{input_files_directory}\\{filename}")
        if self.get_alpha_list() is None:
            print(f"{bcolors.WARNING}Error 1: Alpha list not Found{bcolors.ENDC}")
            time.sleep(3)
            exit(1)


    def get_player_list(self) -> str:
        return self.player_list

    def get_alpha_list(self) -> str:
        return self.alpha_list

    def set_player_list(self, player_list) -> None:
        self.player_list = player_list

    def set_alpha_list(self, alpha_list) -> None:
        self.alpha_list = alpha_list

    def get_players(self) -> list:
        return self.players

    def set_players(self, players) -> None:
        self.players = players

    def get_competition_name(self) -> str:
        return self.competition_name

    def set_competition_name(self, competition_name) -> None:
        self.competition_name = competition_name

    def get_short_date(self) -> str:
        return self.short_date

    def set_short_date(self, short_date) -> None:
        self.short_date = short_date

    def get_long_date(self) -> str:
        return self.long_date

    def set_long_date(self, long_date) -> None:
        self.long_date = long_date

    def get_groups(self) -> list:
        return self.groups

    def set_groups(self, groups) -> None:
        self.groups = groups

    def get_group_number(self) -> int:
        return self.group_number

    def set_group_number(self, group_number) -> None:
        self.group_number = group_number

    def get_forward(self) -> bool:
        return self.forward

    def set_forward(self, forward) -> None:
        self.forward = forward

    def get_end(self):
        return self.end

    def set_end(self, end) -> None:
        self.end = end

    def get_player_numbers(self) -> bool:
        return self.player_numbers

    def set_player_numbers(self, player_numbers) -> None:
        self.player_numbers = player_numbers

    def get_rows(self) -> int:
        return self.row

    def set_rows(self, row) -> None:
        self.row = row