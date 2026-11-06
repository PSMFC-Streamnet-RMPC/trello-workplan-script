import os
from pathlib import Path

import pandas as pd
import requests
from dotenv import load_dotenv

# Parse an Excel workplan and use it to create a Trello board.
load_dotenv()

API_KEY = os.getenv("TRELLO_API_KEY")
TOKEN = os.getenv("TRELLO_TOKEN")
WORKSPACE_ID = os.getenv("WORKSPACE_ID")

required_environment_variables = {
    "TRELLO_API_KEY": API_KEY,
    "TRELLO_TOKEN": TOKEN,
    "WORKSPACE_ID": WORKSPACE_ID,
}

missing_environment_variables = [
    name
    for name, value in required_environment_variables.items()
    if not value
]

if missing_environment_variables:
    raise EnvironmentError(
        "Missing required environment variable(s): "
        + ", ".join(missing_environment_variables)
    )


# Configure the source workbook and worksheet.
workplan_file = Path(
    os.getenv(
        "WORKPLAN_FILE",
        r"C:\Users\tpeterschmidt\OneDrive - PSMFC\Documents\Test_workplan_copy.xlsx",
    )
)
sheet_name = os.getenv("WORKPLAN_SHEET", "Test Trello Version")

if not workplan_file.is_file():
    raise FileNotFoundError(f"Workplan file not found: {workplan_file}")

workplan_df = pd.read_excel(workplan_file, sheet_name=sheet_name)

if workplan_df.empty:
    raise ValueError(
        f"The worksheet '{sheet_name}' in '{workplan_file.name}' contains no rows."
    )

print(f"Loaded {len(workplan_df)} workplan row(s) from '{sheet_name}'.")

# Check that all required columns are present.

required_columns = [
    "BOARD NAME",
    "BOARD MEMBERS",
    "LIST NAME",
    "CARD NAME",
    "CARD LABEL",
    "CARD LABEL COLOR",
    "CARD MEMBERS",
    "CARD START DATE",
    "CARD END DATE",
    "CHECKLIST ITEM DESCRIPTION",
    "CHECKLIST ITEM DUE DATE",
    "CHECKLIST ITEM ASSIGNMENT",
]

missing_columns = [
    column
    for column in required_columns
    if column not in workplan_df.columns
]

if missing_columns:
    raise ValueError(
        "The workplan is missing required column(s): "
        + ", ".join(missing_columns)
    )

# Extract board-level information from the first workplan row.
first_row = workplan_df.iloc[0]

board_name_value = first_row["BOARD NAME"]

if pd.isna(board_name_value) or not str(board_name_value).strip():
    raise ValueError("The first row must include a BOARD NAME.")

new_board_name = str(board_name_value).strip()

board_members_value = first_row["BOARD MEMBERS"]

if pd.isna(board_members_value) or not str(board_members_value).strip():
    new_board_members = []
else:
    new_board_members = [
        member.strip()
        for member in str(board_members_value).split(",")
        if member.strip()
    ]


# Extract unique, non-blank list names.
list_names = (
    workplan_df["LIST NAME"]
    .dropna()
    .astype(str)
    .str.strip()
)

new_lists = list_names[list_names.ne("")].unique().tolist()

if not new_lists:
    raise ValueError("The workplan must include at least one non-blank LIST NAME.")

# ==============================================================================================
# Create or find the board.
# ==============================================================================================

TRELLO_API_BASE = "https://api.trello.com/1"
TRELLO_AUTH = {
    "key": API_KEY,
    "token": TOKEN,
}


def normalize_name(value):
    return str(value).strip().casefold()


try:
    response = requests.get(
        f"{TRELLO_API_BASE}/organizations/{WORKSPACE_ID}/boards",
        params=TRELLO_AUTH,
        timeout=30,
    )
    response.raise_for_status()
except requests.RequestException as error:
    raise RuntimeError("Unable to retrieve existing Trello boards.") from error

existing_board_ids = {
    normalize_name(board["name"]): board["id"]
    for board in response.json()
}

normalized_board_name = normalize_name(new_board_name)

if normalized_board_name in existing_board_ids:
    board_id = existing_board_ids[normalized_board_name]
    print(f"Using existing board '{new_board_name}'.")
else:
    try:
        response = requests.post(
            f"{TRELLO_API_BASE}/boards",
            params={
                **TRELLO_AUTH,
                "name": new_board_name,
                "idOrganization": WORKSPACE_ID,
                "defaultLists": "false",
            },
            timeout=30,
        )
        response.raise_for_status()
    except requests.RequestException as error:
        raise RuntimeError(
            f"Unable to create Trello board '{new_board_name}'."
        ) from error

    board_id = response.json()["id"]
    print(f"Created board '{new_board_name}'.")

# ==============================================================================================
# Add members to the board.
# ==============================================================================================

try:
    response = requests.get(
        f"{TRELLO_API_BASE}/organizations/{WORKSPACE_ID}/members",
        params=TRELLO_AUTH,
        timeout=30,
    )
    response.raise_for_status()
except requests.RequestException as error:
    raise RuntimeError("Unable to retrieve workspace members.") from error

workspace_members_by_name = {
    normalize_name(member["fullName"]): member["id"]
    for member in response.json()
    if member.get("fullName")
}

try:
    response = requests.get(
        f"{TRELLO_API_BASE}/boards/{board_id}/members",
        params=TRELLO_AUTH,
        timeout=30,
    )
    response.raise_for_status()
except requests.RequestException as error:
    raise RuntimeError(
        f"Unable to retrieve members for board '{new_board_name}'."
    ) from error

existing_board_member_ids = {
    member["id"]
    for member in response.json()
}

for member_name in new_board_members:
    member_id = workspace_members_by_name.get(normalize_name(member_name))

    if member_id is None:
        print(f"Member '{member_name}' was not found in the workspace. Skipping.")
        continue

    if member_id in existing_board_member_ids:
        print(f"Member '{member_name}' is already on the board.")
        continue

    try:
        response = requests.put(
            f"{TRELLO_API_BASE}/boards/{board_id}/members/{member_id}",
            params={
                **TRELLO_AUTH,
                "type": "normal",
            },
            timeout=30,
        )
        response.raise_for_status()
    except requests.RequestException as error:
        raise RuntimeError(
            f"Unable to add member '{member_name}' to board '{new_board_name}'."
        ) from error

    existing_board_member_ids.add(member_id)
    print(f"Added member '{member_name}' to the board.")

# ==============================================================================================
# Create labels.
# ==============================================================================================

TRELLO_LABEL_COLORS = {
    "yellow",
    "purple",
    "blue",
    "red",
    "green",
    "orange",
    "black",
    "sky",
    "pink",
    "lime",
}

label_definitions = {}

for row_number, row in workplan_df.iterrows():
    label_name = row["CARD LABEL"]
    label_color = row["CARD LABEL COLOR"]

    if pd.isna(label_name) and pd.isna(label_color):
        continue

    if pd.isna(label_name) or not str(label_name).strip():
        raise ValueError(
            f"Row {row_number + 2}: CARD LABEL COLOR is set, but CARD LABEL is blank."
        )

    if pd.isna(label_color) or not str(label_color).strip():
        raise ValueError(
            f"Row {row_number + 2}: CARD LABEL is set, but CARD LABEL COLOR is blank."
        )

    label_name = str(label_name).strip()
    label_color = str(label_color).strip().lower()
    label_key = normalize_name(label_name)

    if label_color not in TRELLO_LABEL_COLORS:
        allowed_colors = ", ".join(sorted(TRELLO_LABEL_COLORS))
        raise ValueError(
            f"Row {row_number + 2}: '{label_color}' is not a valid Trello label "
            f"color. Use one of: {allowed_colors}."
        )

    if (
        label_key in label_definitions
        and label_definitions[label_key]["color"] != label_color
    ):
        raise ValueError(
            f"Label '{label_name}' is assigned more than one color in the workplan."
        )

    label_definitions[label_key] = {
        "name": label_name,
        "color": label_color,
    }


try:
    response = requests.get(
        f"{TRELLO_API_BASE}/boards/{board_id}/labels",
        params=TRELLO_AUTH,
        timeout=30,
    )
    response.raise_for_status()
except requests.RequestException as error:
    raise RuntimeError(
        f"Unable to retrieve labels for board '{new_board_name}'."
    ) from error

label_dict = {
    normalize_name(label["name"]): label["id"]
    for label in response.json()
    if label.get("name")
}

for label_key, definition in label_definitions.items():
    label_name = definition["name"]
    label_color = definition["color"]

    if label_key in label_dict:
        print(f"Using existing label '{label_name}'.")
        continue

    try:
        response = requests.post(
            f"{TRELLO_API_BASE}/boards/{board_id}/labels",
            params={
                **TRELLO_AUTH,
                "name": label_name,
                "color": label_color,
            },
            timeout=30,
        )
        response.raise_for_status()
    except requests.RequestException as error:
        raise RuntimeError(
            f"Unable to create label '{label_name}' on board '{new_board_name}'."
        ) from error

    label_dict[label_key] = response.json()["id"]
    print(f"Created label '{label_name}' ({label_color}).")

# ==============================================================================================
# Create lists.
# ==============================================================================================

try:
    response = requests.get(
        f"{TRELLO_API_BASE}/boards/{board_id}/lists",
        params=TRELLO_AUTH,
        timeout=30,
    )
    response.raise_for_status()
except requests.RequestException as error:
    raise RuntimeError(
        f"Unable to retrieve lists for board '{new_board_name}'."
    ) from error

list_dict = {}

for trello_list in response.json():
    list_key = normalize_name(trello_list["name"])

    if list_key in list_dict:
        raise RuntimeError(
            f"The board '{new_board_name}' has multiple lists named "
            f"'{trello_list['name']}'. Rename or remove the duplicates before "
            "running this script."
        )

    list_dict[list_key] = trello_list["id"]

for list_name in new_lists:
    list_key = normalize_name(list_name)

    if list_key in list_dict:
        print(f"Using existing list '{list_name}'.")
        continue

    try:
        response = requests.post(
            f"{TRELLO_API_BASE}/lists",
            params={
                **TRELLO_AUTH,
                "idBoard": board_id,
                "name": list_name,
            },
            timeout=30,
        )
        response.raise_for_status()
    except requests.RequestException as error:
        raise RuntimeError(
            f"Unable to create list '{list_name}' on board '{new_board_name}'."
        ) from error

    list_dict[list_key] = response.json()["id"]
    print(f"Created list '{list_name}'.")

# ==============================================================================================
# Create cards.
# ==============================================================================================

def to_trello_timestamp(value, row_number, column_name):
    if pd.isna(value) or not str(value).strip():
        return None

    timestamp = pd.to_datetime(value, errors="coerce")

    if pd.isna(timestamp):
        raise ValueError(
            f"Row {row_number}: '{value}' is not a valid {column_name}."
        )

    if timestamp.tzinfo is None:
        timestamp = timestamp.tz_localize("UTC")
    else:
        timestamp = timestamp.tz_convert("UTC")

    return timestamp.isoformat().replace("+00:00", "Z")


try:
    response = requests.get(
        f"{TRELLO_API_BASE}/boards/{board_id}/cards",
        params=TRELLO_AUTH,
        timeout=30,
    )
    response.raise_for_status()
except requests.RequestException as error:
    raise RuntimeError(
        f"Unable to retrieve cards for board '{new_board_name}'."
    ) from error

card_ids_by_location = {}

for card in response.json():
    card_key = (
        card["idList"],
        normalize_name(card["name"]),
    )

    if card_key in card_ids_by_location:
        raise RuntimeError(
            f"The board has multiple cards named '{card['name']}' "
            "in the same list. Rename or remove the duplicates before running "
            "this script."
        )

    card_ids_by_location[card_key] = card["id"]


for row_number, row in workplan_df.iterrows():
    spreadsheet_row = row_number + 2

    list_name_value = row["LIST NAME"]
    card_name_value = row["CARD NAME"]

    has_list_name = (
        pd.notna(list_name_value) and str(list_name_value).strip()
    )
    has_card_name = (
        pd.notna(card_name_value) and str(card_name_value).strip()
    )

    if not has_list_name and not has_card_name:
        continue

    if not has_list_name or not has_card_name:
        raise ValueError(
            f"Row {spreadsheet_row}: LIST NAME and CARD NAME must both be provided."
        )

    list_name = str(list_name_value).strip()
    card_name = str(card_name_value).strip()
    list_id = list_dict.get(normalize_name(list_name))

    if list_id is None:
        raise ValueError(
            f"Row {spreadsheet_row}: list '{list_name}' could not be found."
        )

    card_key = (list_id, normalize_name(card_name))

    if card_key in card_ids_by_location:
        print(f"Using existing card '{card_name}' in '{list_name}'.")
        continue

    members_raw = row["CARD MEMBERS"]

    if pd.isna(members_raw) or not str(members_raw).strip():
        member_ids = []
    else:
        member_ids = []

        for member_name in str(members_raw).split(","):
            member_name = member_name.strip()

            if not member_name:
                continue

            member_id = workspace_members_by_name.get(
                normalize_name(member_name)
            )

            if member_id is None:
                print(
                    f"Row {spreadsheet_row}: card member '{member_name}' "
                    "was not found in the workspace. Skipping that assignment."
                )
                continue

            member_ids.append(member_id)

    start_date = to_trello_timestamp(
        row["CARD START DATE"],
        spreadsheet_row,
        "CARD START DATE",
    )
    due_date = to_trello_timestamp(
        row["CARD END DATE"],
        spreadsheet_row,
        "CARD END DATE",
    )

    label_value = row["CARD LABEL"]
    label_id = None

    if pd.notna(label_value) and str(label_value).strip():
        label_name = str(label_value).strip()
        label_id = label_dict.get(normalize_name(label_name))

        if label_id is None:
            raise ValueError(
                f"Row {spreadsheet_row}: label '{label_name}' was not found."
            )

    card_params = {
        **TRELLO_AUTH,
        "idList": list_id,
        "name": card_name,
    }

    if member_ids:
        card_params["idMembers"] = ",".join(member_ids)

    if start_date:
        card_params["start"] = start_date

    if due_date:
        card_params["due"] = due_date
        card_params["dueReminder"] = 120

    if label_id:
        card_params["idLabels"] = label_id

    try:
        response = requests.post(
            f"{TRELLO_API_BASE}/cards",
            params=card_params,
            timeout=30,
        )
        response.raise_for_status()
    except requests.RequestException as error:
        raise RuntimeError(
            f"Unable to create card '{card_name}' in list '{list_name}'."
        ) from error

    card_ids_by_location[card_key] = response.json()["id"]
    print(f"Created card '{card_name}' in '{list_name}'.")

# ==============================================================================================
# Create checklists and checklist items.
# ==============================================================================================

# The script uses one checklist per card, with the checklist named after the card.
source_cards = {}

for row_number, row in workplan_df.iterrows():
    list_name_value = row["LIST NAME"]
    card_name_value = row["CARD NAME"]

    if (
        pd.isna(list_name_value)
        or not str(list_name_value).strip()
        or pd.isna(card_name_value)
        or not str(card_name_value).strip()
    ):
        continue

    list_name = str(list_name_value).strip()
    card_name = str(card_name_value).strip()
    list_id = list_dict[normalize_name(list_name)]
    card_key = (list_id, normalize_name(card_name))
    card_id = card_ids_by_location.get(card_key)

    if card_id is None:
        raise RuntimeError(
            f"Could not find Trello card '{card_name}' in list '{list_name}'."
        )

    source_cards[card_key] = {
        "card_id": card_id,
        "card_name": card_name,
        "list_name": list_name,
    }


try:
    response = requests.get(
        f"{TRELLO_API_BASE}/boards/{board_id}/cards",
        params={
            **TRELLO_AUTH,
            "checklists": "all",
        },
        timeout=30,
    )
    response.raise_for_status()
except requests.RequestException as error:
    raise RuntimeError(
        f"Unable to retrieve checklists for board '{new_board_name}'."
    ) from error

checklist_ids_by_card_key = {}

for card in response.json():
    card_key = (
        card["idList"],
        normalize_name(card["name"]),
    )

    matching_checklists = [
        checklist
        for checklist in card.get("checklists", [])
        if normalize_name(checklist["name"]) == normalize_name(card["name"])
    ]

    if len(matching_checklists) > 1:
        raise RuntimeError(
            f"Card '{card['name']}' has multiple checklists with the same name."
        )

    if matching_checklists:
        checklist_ids_by_card_key[card_key] = matching_checklists[0]["id"]


# Create a checklist for each workplan card when one does not already exist.
for card_key, card_details in source_cards.items():
    if card_key in checklist_ids_by_card_key:
        continue

    try:
        response = requests.post(
            f"{TRELLO_API_BASE}/cards/{card_details['card_id']}/checklists",
            params={
                **TRELLO_AUTH,
                "name": card_details["card_name"],
            },
            timeout=30,
        )
        response.raise_for_status()
    except requests.RequestException as error:
        raise RuntimeError(
            f"Unable to create a checklist for card "
            f"'{card_details['card_name']}'."
        ) from error

    checklist_ids_by_card_key[card_key] = response.json()["id"]

    print(
        f"Created checklist for card '{card_details['card_name']}' "
        f"in '{card_details['list_name']}'."
    )


# Fetch existing checklist items once per checklist.
existing_checklist_items = {}

for checklist_id in checklist_ids_by_card_key.values():
    try:
        response = requests.get(
            f"{TRELLO_API_BASE}/checklists/{checklist_id}/checkItems",
            params=TRELLO_AUTH,
            timeout=30,
        )
        response.raise_for_status()
    except requests.RequestException as error:
        raise RuntimeError(
            f"Unable to retrieve items for checklist '{checklist_id}'."
        ) from error

    existing_checklist_items[checklist_id] = {
        normalize_name(item["name"])
        for item in response.json()
    }


for row_number, row in workplan_df.iterrows():
    spreadsheet_row = row_number + 2
    checklist_item_value = row["CHECKLIST ITEM DESCRIPTION"]

    if pd.isna(checklist_item_value) or not str(checklist_item_value).strip():
        continue

    list_name = str(row["LIST NAME"]).strip()
    card_name = str(row["CARD NAME"]).strip()
    checklist_item = str(checklist_item_value).strip()

    list_id = list_dict[normalize_name(list_name)]
    card_key = (list_id, normalize_name(card_name))
    checklist_id = checklist_ids_by_card_key[card_key]
    checklist_item_key = normalize_name(checklist_item)

    if checklist_item_key in existing_checklist_items[checklist_id]:
        print(
            f"Checklist item '{checklist_item}' already exists on "
            f"card '{card_name}'."
        )
        continue

    due_date = to_trello_timestamp(
        row["CHECKLIST ITEM DUE DATE"],
        spreadsheet_row,
        "CHECKLIST ITEM DUE DATE",
    )

    assignee_value = row["CHECKLIST ITEM ASSIGNMENT"]
    assigned_member_id = None

    if pd.notna(assignee_value) and str(assignee_value).strip():
        assignee_name = str(assignee_value).strip()
        assigned_member_id = workspace_members_by_name.get(
            normalize_name(assignee_name)
        )

        if assigned_member_id is None:
            print(
                f"Row {spreadsheet_row}: checklist assignee '{assignee_name}' "
                "was not found in the workspace. Creating the item unassigned."
            )

    checklist_item_params = {
        **TRELLO_AUTH,
        "name": checklist_item,
    }

    if due_date:
        checklist_item_params["due"] = due_date

    if assigned_member_id:
        checklist_item_params["idMember"] = assigned_member_id

    try:
        response = requests.post(
            f"{TRELLO_API_BASE}/checklists/{checklist_id}/checkItems",
            params=checklist_item_params,
            timeout=30,
        )
        response.raise_for_status()
    except requests.RequestException as error:
        raise RuntimeError(
            f"Unable to create checklist item '{checklist_item}' "
            f"for card '{card_name}'."
        ) from error

    existing_checklist_items[checklist_id].add(checklist_item_key)

    print(f"Created checklist item '{checklist_item}' for card '{card_name}'.")
    



print("Workplan import complete.")




