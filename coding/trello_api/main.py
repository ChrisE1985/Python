from reusables.trelloapi import *
from reusables.worddocx import *

if __name__ == '__main__':
    # Trello API Variables
    # My TRELLO API KEY and TOKEN. You can get your own by going to "https://trello.com/app-key"
    API_KEY = "333a211aee74136dcb91a5d8ebcd1abf"
    OAUTH_TOKEN = "138d8c4256affe1e6d1bd8bdf1227aa9c720d7b604e8c579694b6751761cf43e"

    list_names = ["Team Has", "Team Chris", "Team Monz", "Team Callum", "Floating with Steve & Ryan", "Incoming", "Approved / Imminent"]
    full_team_cards = []
    rec_comp_cards = []
    full_app_cards = []
    full_in_cards = []

    board_id = findBoard(API_KEY, OAUTH_TOKEN, "Testing Activities")

    for list_name in list_names:
        list_info = findList(API_KEY, OAUTH_TOKEN, board_id, list_name)
        list_id = list_info[0]
        list_name = list_info[1]
        team_cards = findCards(API_KEY, OAUTH_TOKEN, list_id, list_name)
        if list_name == "Incoming":
            for team_card in team_cards:
                full_in_cards.append(team_card)
        elif list_name == "Approved / Imminent":
            for team_card in team_cards:
                full_app_cards.append(team_card)
            print(len(full_app_cards))
