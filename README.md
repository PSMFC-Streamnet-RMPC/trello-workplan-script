# Trello Workplan Importer

Create a Trello board from an Excel workplan, including members, lists, labels, cards, checklists, and checklist items.

The script can also add missing items to an existing board. **Existing cards and checklist items are skipped rather than updated.**

## Requirements

- Python with `pip` available.
- A Trello account with permission to create boards and add members in the target workspace.
- A Trello API key, API token, and workspace ID.
- An Excel workplan using the structure in [workplan_template.xlsx](workplan_template.xlsx).

Checklist item assignments and due dates use Trello’s advanced checklist features. See [Trello’s checklist requirements](https://support.atlassian.com/trello/docs/how-to-use-advanced-checklists-to-set-due-dates).

## 1. Install dependencies

Download or clone this repository, then open a terminal in the project’s root folder.

Install the required packages:

```powershell
python -m pip install pandas requests python-dotenv openpyxl
```

For an isolated Python environment, you can instead run:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install pandas requests python-dotenv openpyxl
```

If you use this environment, run the script with `.\.venv\Scripts\python.exe` in place of `python`.

## 2. Get your Trello credentials

Follow [Trello’s REST API setup instructions](https://support.atlassian.com/trello/docs/getting-started-with-trello-rest-api/) to create an API key and generate a token. The token needs read and write access for the operations performed by this script.

You also need the ID of the workspace where the board will be created. Trello’s [Get an Organization endpoint](https://developer.atlassian.com/cloud/trello/rest/api-group-organizations/#api-organizations-id-get) accepts a workspace ID or name:

```text
GET https://api.trello.com/1/organizations/{workspace_name}
```

Authenticate the request with your API key and token, then copy the `id` value from the response.

## 3. Configure the script

Create a file named `.env` in the project’s root folder, alongside `README.md`.

Add the following settings, replacing the example values:

```dotenv
TRELLO_API_KEY=your_api_key
TRELLO_TOKEN=your_api_token
WORKSPACE_ID=your_workspace_id

WORKPLAN_FILE='C:\Users\YourName\Documents\My_workplan.xlsx'
WORKPLAN_SHEET='My Workplan'
```

| Setting | Purpose |
| --- | --- |
| `TRELLO_API_KEY` | Your Trello API key. Required. |
| `TRELLO_TOKEN` | Your Trello API token. Required. |
| `WORKSPACE_ID` | Target Trello workspace ID. Required. |
| `WORKPLAN_FILE` | Path to the Excel workbook. Set this to your own file. |
| `WORKPLAN_SHEET` | Exact worksheet tab name. Defaults to `Test Trello Version` when unset. |

Use an absolute workbook path to avoid confusion about which folder the script runs from.

The script currently includes a machine-specific fallback workbook path. Setting `WORKPLAN_FILE` avoids needing to edit that path in the Python file.

Keep `.env` private. The repository’s `.gitignore` already excludes it.

### When configuration changes do not take effect

Save `.env` and restart the script after making changes.

The current script uses `load_dotenv()`, which preserves environment variables already supplied by the terminal or operating system. An older terminal value can therefore take precedence over your edited `.env` file.

Close the terminal and open a new one if old values persist. If you want the project’s `.env` file to take precedence, replace `load_dotenv()` with:

```python
env_file = Path(__file__).resolve().parent.parent / ".env"
load_dotenv(env_file, override=True)
```

A `.vscode` folder is optional; the script loads `.env` itself.

## 4. Prepare the Excel workplan

Start with a copy of `workplan_template.xlsx`. Put the column headings in the first worksheet row and preserve their exact spelling.

All of the following columns must exist, even when their cells are blank:

| Column | Instructions |
| --- | --- |
| `BOARD NAME` | Required in the first data row. Used as the board name for the entire worksheet. |
| `BOARD MEMBERS` | Comma-separated Trello full names. Read from the first data row only. |
| `LIST NAME` | List containing the card. |
| `CARD NAME` | Card title. |
| `CARD LABEL` | Optional label name. One label per row. |
| `CARD LABEL COLOR` | Required whenever a label name is provided. |
| `CARD MEMBERS` | Optional comma-separated Trello full names. |
| `CARD START DATE` | Optional card start date. |
| `CARD END DATE` | Optional card due date. |
| `CHECKLIST ITEM DESCRIPTION` | Optional checklist item text. One item per row. |
| `CHECKLIST ITEM DUE DATE` | Optional checklist item due date. |
| `CHECKLIST ITEM ASSIGNMENT` | Optional Trello full name. One person per item. |

### Row structure

Enter both `LIST NAME` and `CARD NAME` on every row describing a card or checklist item.

For multiple checklist items on the same card, repeat the list and card names on separate rows:

| LIST NAME | CARD NAME | CHECKLIST ITEM DESCRIPTION |
| --- | --- | --- |
| Planning | Prepare workplan | Draft scope |
| Planning | Prepare workplan | Review budget |
| Planning | Prepare workplan | Approve schedule |

This creates one card named **Prepare workplan**, with a checklist of the same name containing three items.

For repeated cards, keep the card-level values consistent. The first occurrence supplies the settings when a new card is created; later occurrences contribute checklist items.

### Members

Use members’ Trello **full names**, rather than usernames or email addresses. Members must already belong to the target workspace.

Include card members and checklist assignees in `BOARD MEMBERS` so they are added to the board before assignments are made. Use unique full names to avoid ambiguous matches.

### Labels

Supported colors are:

```text
yellow, purple, blue, red, green, orange, black, sky, pink, lime
```

Provide both the label name and color, or leave both blank. Use the same color every time a label name appears.

### Dates

Use Excel date cells or unambiguous values such as `2026-11-05`.

Dates without a timezone are interpreted as UTC. A date without a time becomes midnight UTC, which may appear as the previous calendar day in Trello depending on your local timezone.

## 5. Run the import

Save the workbook and `.env`, then run this command from the project’s root folder:

```powershell
python script/parse_workplan.py
```

Or, if you created the virtual environment above:

```powershell
.\.venv\Scripts\python.exe script/parse_workplan.py
```

The script prints progress as it creates or reuses Trello objects. A completed import ends with:

```text
Workplan import complete.
```

Running the script writes directly to Trello. There is no preview mode or automatic rollback; an error can leave a partially completed import.

## Running the script again

The script compares names without regard to capitalization or surrounding spaces.

| Object | Behavior on subsequent runs |
| --- | --- |
| Board | Reuses a matching board in the target workspace. |
| Board members | Adds members who are not already on the board. |
| Labels | Reuses matching label names without updating their colors. |
| Lists | Reuses matching list names. |
| Cards | Reuses matching card names within the same list without updating their settings. |
| Checklists | Reuses the checklist matching the card’s name. |
| Checklist items | Skips matching item names within the same checklist. |

You can add rows and rerun the script to create missing objects. Changes to existing card dates, labels, members, or checklist item assignments are not synchronized.

Renaming or moving objects can cause additional objects to be created because matching uses names and locations rather than persistent spreadsheet IDs.

## Troubleshooting

| Problem | What to check |
| --- | --- |
| Import warning or `ModuleNotFoundError` | Install dependencies in the same Python environment used to run the script. In VS Code, check **Python: Select Interpreter**. |
| Missing environment variables | Confirm `.env` is saved in the project root and uses the exact variable names shown above. |
| Workbook not found | Check `WORKPLAN_FILE` and confirm the file exists at that location. |
| Worksheet not found | Match `WORKPLAN_SHEET` to the Excel worksheet tab name. |
| Missing required columns | Restore the exact headings from the template. |
| Member not found | Match the member’s Trello full name and confirm workspace membership. Unmatched assignments are skipped. |
| Duplicate list, card, or checklist error | Rename or remove the duplicates identified by the error before rerunning. |
| Trello request failed | Check credentials, token access, workspace permissions, network access, and the underlying error message. |
| Workbook edits do not change an existing card | Existing cards and checklist items are skipped. Edit those objects directly in Trello. |