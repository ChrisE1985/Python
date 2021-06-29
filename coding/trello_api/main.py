from reusables.trelloapi import *
from reusables.tw_worddocx import *

if __name__ == '__main__':
    # Trello API Variables
    # My TRELLO API KEY and TOKEN. You can get your own by going to "https://trello.com/app-key"
    API_KEY = "333a211aee74136dcb91a5d8ebcd1abf"
    OAUTH_TOKEN = "138d8c4256affe1e6d1bd8bdf1227aa9c720d7b604e8c579694b6751761cf43e"

    # List all of the lists in "Team Activities" trello board that you wish to loop through
    list_names = ["Team Has", "Team Chris", "Team Monz", "Team Callum", "Floating with Steve & Ryan", "Incoming", "Approved / Imminent"]
    # Create empty lists that will be populated for later use in the report generation
    full_team_cards = []
    full_comp_cards = []
    full_app_cards = []
    full_in_cards = []

    # board_id is used to return all lists
    board_id = findBoard(API_KEY, OAUTH_TOKEN, "Testing Activities")

    # Using the above list of list names loop through and populate each of full_team_cards e.g.  Team Has,
    # full_comp_cards e.g. cards with complete label, full_app_cards e.g. approved/imminent list and finally all
    # incoming cards
    for list_name in list_names:
        list_info = findList(API_KEY, OAUTH_TOKEN, board_id, list_name)# findList returns 2 items, list id and list name
        list_id = list_info[0]
        list_name = list_info[1]
        # team_cards will always return a list of 2 (the first is the one we want and the second is just a list of
        # completed cards in each list which might be empty)
        team_cards = findCards(API_KEY, OAUTH_TOKEN, list_id, list_name)
        if list_name == "Incoming":
            for team_card in team_cards[0]:
                full_in_cards.append(team_card)
        elif list_name == "Approved / Imminent":
            for team_card in team_cards[0]:
                full_app_cards.append(team_card)
        else:
            if len(team_cards) > 1:
                for team_card in team_cards[0]:
                    full_team_cards.append(team_card)
                for team_card in team_cards[1]:
                    full_comp_cards.append(team_card)
            else:
                for team_card in team_cards:
                    full_team_cards.append(team_card)

    # Report template and output locations
    templateLoc = r"templates/TestingWeeklyReport.docx"
    outputLoc = r"C:/Users/Chris E/Desktop/"

    # Generate Report
    generateWordDoc(full_team_cards, full_comp_cards, full_app_cards, full_in_cards, templateLoc, outputLoc)
