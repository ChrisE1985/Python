import requests  # pip install requests (Requests can be used to call APIs)
import arrow  # pip install arrow
import datetime  # comes with normal python so no need to install
from datetime import datetime

# Interrogates TRELLO to find your board
def findBoard(API_KEY, OAUTH_TOKEN, BOARD_NAME=None):
    get_boards_url = "https://api.trello.com/1/members/me/boards?key=" + API_KEY + "&token=" + OAUTH_TOKEN + "&response_type=token"

    r = requests.get(get_boards_url)

    for boards in r.json():
        board_id = ""
        board_name = ""
        for key, value in boards.items():
            if key == "id":
                board_id = value
            elif key == "name":
                board_name = value
        if board_name == BOARD_NAME:
            print(f"Found the '{BOARD_NAME}' board.")
            return board_id


# Interrogates TRELLO to find every list on the found board
def findList(API_KEY, OAUTH_TOKEN, board_id, LIST_NAME=None):
    get_lists_url = "https://api.trello.com/1/boards/" + board_id + "/lists?key=" + API_KEY + "&token=" + OAUTH_TOKEN + "&response_type=token"

    list_of_lists = []

    r = requests.get(get_lists_url)

    for lists in r.json():
        list_id = ""
        list_name = ""
        for key, value in lists.items():
            if key == "id":
                list_id = value
            elif key == "name":
                list_name = value
        if LIST_NAME is None:
            if list_name in ("Incoming", "Approved / Imminent", "Recently Complete", "Out of Scope"):
                continue
            else:
                list_of_lists.append([list_id, list_name])
        elif list_name == LIST_NAME:
            print(f"Found '{LIST_NAME}' list.")
            return list_id, list_name
        else:
            continue
    return list_of_lists


# Interrogates TRELLO to find the members on a card
def findMembers(API_KEY, OAUTH_TOKEN, card_id):
    get_members_url = f"https://api.trello.com/1/cards/{card_id}/members?key={API_KEY}&token={OAUTH_TOKEN}&response_type=token"

    list_of_members = []

    m = requests.get(get_members_url)

    for member in m.json():
        card_member = ""
        for key, value in member.items():
            if key == "fullName":
                if value == "monzahmed":
                    card_member = "Monz"
                else:
                    card_member = value.replace(".", " ")
                    card_member = card_member.split(" ")[0]
        list_of_members.append(card_member.title())
    return list_of_members


# Interrogates TRELLO to find the labels on a card
def findLabels(API_KEY, OAUTH_TOKEN, card_id):
    get_labels_url = f"https://api.trello.com/1/cards/{card_id}/labels?key={API_KEY}&token={OAUTH_TOKEN}&response_type=token"

    list_of_labels = []

    l = requests.get(get_labels_url)

    for label in l.json():
        card_label = ""
        for key, value in label.items():
            if key == "name":
                card_label = value
        list_of_labels.append(card_label)
    return list_of_labels


# Interrogates TRELLO to find the latest ":warning:" comment on a card
def getLatestComment(API_KEY, OAUTH_TOKEN, card_id):
    get_comments_url = f"https://api.trello.com/1/cards/{card_id}/actions?filter=commentCard&key={API_KEY}&token={OAUTH_TOKEN}&response_type=token"

    list_of_comments = []

    c = requests.get(get_comments_url)
    # print(c.json())
    for comment in c.json():
        card_comment = ""
        for key, value in comment.items():
            if key == "date":
                comment_date = arrow.get(value)
                comment_date = comment_date.datetime.strftime("%d/%m/%Y")
                comment_date = datetime.strptime(comment_date, "%d/%m/%Y").date()
            if key == "data":
                card_comment = value['text']
        if ":warning: " in card_comment:
            card_comment = card_comment.replace(":warning: ", "")
            list_of_comments.append(card_comment)
    return list_of_comments


# Interrogates TRELLO to find every card in a list
def findCards(API_KEY, OAUTH_TOKEN, list_id, list_name=None):
    get_cards_url = f"https://api.trello.com/1/lists/{list_id}/cards?key={API_KEY}&token={OAUTH_TOKEN}&response_type=token"

    list_of_all_cards = []
    list_of_rec_complete_cards = []

    r = requests.get(get_cards_url)

    for cards in r.json():
        card_id = ""
        card_name = ""
        card_start = ""
        card_due = ""
        card_last_activity = ""
        card_desc = ""
        card_labels = ""
        list_of_members = []
        list_of_labels = []
        list_of_comments = []

        for key, value in cards.items():
            if key == "id":
                card_id = value
                list_of_members = findMembers(API_KEY, OAUTH_TOKEN, card_id)
                list_of_labels = findLabels(API_KEY, OAUTH_TOKEN, card_id)
                list_of_comments = getLatestComment(API_KEY, OAUTH_TOKEN, card_id)
            elif key == "name":
                card_name = value
            elif key == "start":
                if value is not None:
                    card_start = value
                    card_start = arrow.get(card_start).datetime
                    card_start = card_start.strftime("%d/%m/%Y")
                    card_start = datetime.strptime(card_start, "%d/%m/%Y").date()
                else:
                    card_start = "N/A"
            elif key == "due":
                if value is not None:
                    card_due = value
                    card_due = arrow.get(card_due).datetime
                    card_due = card_due.strftime("%d/%m/%Y")
                    card_due = datetime.strptime(card_due, "%d/%m/%Y").date()
                else:
                    card_due = "N/A"
            elif key == "dueComplete":
                if value is True:
                    card_complete = "Yes"
                else:
                    card_complete = "No"
            elif key == "dateLastActivity":
                card_last_activity = value
                card_last_activity = arrow.get(card_last_activity).datetime
                card_last_activity = card_last_activity.strftime("%d/%m/%Y")
                card_last_activity = datetime.strptime(card_last_activity, "%d/%m/%Y").date()
            elif key == "desc":
                card_desc = value
        if "Complete" in list_of_labels or card_complete == "Yes":
            if len(list_of_comments) > 1:
                list_of_rec_complete_cards.append([card_name, list_of_labels, list_of_members, card_start, card_due, card_complete, card_last_activity, list_of_comments[0]])
            else:
                list_of_rec_complete_cards.append([card_name, list_of_labels, list_of_members, card_start, card_due, card_complete, card_last_activity, list_of_comments])
        else:
            if len(list_of_comments) > 1:
                list_of_all_cards.append([card_name, list_of_labels, list_of_members, card_start, card_due, card_complete, card_last_activity, list_of_comments[0]])
            else:
                list_of_all_cards.append([card_name, list_of_labels, list_of_members, card_start, card_due, card_complete, card_last_activity, list_of_comments])
    if len(list_of_all_cards) > 0:
        return list_of_all_cards, list_of_rec_complete_cards
