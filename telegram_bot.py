import os
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from typing import Final
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# Path to your service account key file
TOKEN = 'Path'

# Define the scopes
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

SPREADSHEET_ID = "spreadsheetid"
TELEGRAM_TOKEN: Final = "botToken"
BOT_USERNAME: Final = "@events_attendings_bot"


# Function to update Google Sheet
def update_sheet(spreadsheet_id, event_name, event_date, participants):
    service = setup_google_sheets()
    # Check if the sheet for the event already exists
    try:
        # Check if the event sheet already exists
        sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', '')
        sheet_titles = [sheet.get('properties', {}).get('title', '') for sheet in sheets]

        # Create the sheet if it doesn't exist
        sheet_id = None
        if event_name not in sheet_titles:
            batch_update_request_body = {
                'requests': [
                    {'addSheet': {'properties': {'title': event_name}}}
                ]
            }
            response = service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=batch_update_request_body
            ).execute()
            sheet_id = response['replies'][0]['addSheet']['properties']['sheetId']

        # If the event sheet is new, prepare the header
        if event_name not in sheet_titles:
            header_values = [['Event Name', 'Event Date', 'Participant Name']]
            body = {'values': header_values}
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id, range=f"{event_name}!A1:C1",
                valueInputOption='USER_ENTERED', body=body
            ).execute()

        # Prepare the participant data
        participant_values = [[event_name, event_date, participants[0]]] if participants else []
        participant_values.extend([["", "", participant] for participant in participants[1:]])

        # Define the range and body for the update request
        range_name = f"{event_name}!A1:C"
        body = {'values': participant_values}

        # Append the data to the sheet
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id, range=range_name,
            valueInputOption='USER_ENTERED', body=body, insertDataOption='INSERT_ROWS'
        ).execute()

        # Apply formatting to the header and participant cells
        requests = [
            # Request for the header row formatting
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 3,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 0.827,  # Light gray
                                "green": 0.827,
                                "blue": 0.827
                            },
                            "textFormat": {
                                "foregroundColor": {
                                    "red": 0.627,  # Purple
                                    "green": 0.125,
                                    "blue": 0.941
                                },
                                "fontSize": 12,
                                "bold": True
                            },
                            "horizontalAlignment": "CENTER"
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
                }
            },
            # Request for coloring all participant cells including event name and date
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 1,
                        "endRowIndex": 1 + len(participants),
                        "startColumnIndex": 0,
                        "endColumnIndex": 3,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 0.0,  # Black
                                "green": 0.0,
                                "blue": 0.0
                            },
                            "textFormat": {
                                "foregroundColor": {
                                    "red": 1.0,  # White
                                    "green": 1.0,
                                    "blue": 1.0
                                },
                                "fontSize": 12,
                                "bold": False
                            },
                            "horizontalAlignment": "CENTER"
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
                }
            }
        ]

        requests.extend([
            # Set the width of the 'Event Name' column
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,  # Column A
                        "endIndex": 1  # Column B
                    },
                    "properties": {
                        "pixelSize": 150  # Adjust the width as needed
                    },
                    "fields": "pixelSize"
                }
            },
            # Set the width of the 'Event Date' column
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 1,  # Column B
                        "endIndex": 2  # Column C
                    },
                    "properties": {
                        "pixelSize": 120  # Adjust the width as needed
                    },
                    "fields": "pixelSize"
                }
            },
            # Set the width of the 'Participant Name' column
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 2,  # Column C
                        "endIndex": 3  # The end index is exclusive
                    },
                    "properties": {
                        "pixelSize": 200  # Adjust the width as needed for long names
                    },
                    "fields": "pixelSize"
                }
            }
        ])

        # Send the batchUpdate request to apply the formatting
        if requests:
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": requests}
            ).execute()

        return f"Event '{event_name}' logged successfully."

    except HttpError as error:
        print(f"An error occurred: {error}")
        if error.error_details == "Invalid requests[0].repeatCell: No grid with id: 0":
            return f"Event '{event_name}' logged successfully."
        return "Failed to log the event."

# Telegram bot handlers
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text('Hello! Send me an event name and participants list.')


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = """
Hello! Here's how you can use this bot:
- To log an event, simply send me the event name and date followed by the participants' list with the $ sign between each. And a period between the names. For example:
  'EventName$28/02/2024$John, Jane Doe, Well S'
- Use /start to restart the bot.
- Use /help to display this message again.

Make sure to replace 'EventName' with the actual name of your event and list all participants names separated by spaces.
"""
    await update.message.reply_text(help_text)


async def log_event(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    # Simple parsing: First word is event name, rest are participants
    event_name, event_date, participants = text.split("$")
    participants_list = participants.split(", ")
    update_sheet(SPREADSHEET_ID, event_name, event_date, participants_list)
    await update.message.reply_text('Event logged successfully!')


async def handle_greeting(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_message = update.message.text.lower()
    chat_type = update.message.chat.type
    bot_username = f"@{context.bot.username}"  # Dynamically get bot's username

    # Check if the bot is mentioned in group chats or if it's a private chat
    if bot_username.lower() in user_message or chat_type == "private":
        if any(greeting in user_message for greeting in ['hello', 'hi', 'hey', 'Hello', 'Hi', 'Hey']):
            reply = "Hello! Please send me the event name and date followed by the participants' list with the $ sign between each. And a period between the names, For example: 'EventName$28/02/2024$John, Jane Doe, Well S'."
            print("Bot: ", reply)
            await update.message.reply_text(reply)
        else:
            reply = "I'm not sure how to respond to that. If you need help, type /help."
            print("Bot: ", reply)
            await update.message.reply_text(reply)
    elif chat_type in ["group", "supergroup"]:
        # If the bot is not mentioned in a group chat, do not respond
        return


async def handle_messages(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message_text = update.message.text
    chat_type = update.message.chat.type
    bot_username = f"@{context.bot.username}"  # Dynamically get bot's username
    print(f'User ({update.message.chat.id}) in {update.message.chat.type}: "{update.message.text}"')

    # Check if the bot is mentioned in groups or it's a direct message
    if bot_username.lower() in message_text.lower() or chat_type == "private":
        parts = message_text.split()
        if len(parts) < 2:
            await handle_greeting(update, context)
            return

        try:

            # Remove the bot's username from the message if present
            clean_text = message_text.replace(bot_username, "")
            event_name, event_date, participants = clean_text.split("$")
            participants_list = participants.split(", ")
            response = update_sheet(SPREADSHEET_ID, event_name, event_date, participants_list)
            print("Bot: ", response)
            await update.message.reply_text(response)

        except ValueError:
            await handle_greeting(update, context)
            reply = "Incorrect format. Please use the format 'EventName$Date$Participant1, Participant2, ...'"
            print("Bot: ", reply)
            await update.message.reply_text(reply)


    else:
        # Optional: Handle other messages or ignore them
        await handle_greeting(update, context)


async def error(update: Update, context: ContextTypes.DEFAULT_TYPE):
    print(f"This update {update} caused the error {context.error}")


def setup_google_sheets():

    service = None

    if os.path.exists(TOKEN):
        print("The file exists and is accessible.")
    else:
        print("The file does not exist or is not accessible. Check the path.")

    # Authenticate using the service account file
    credentials = Credentials.from_service_account_file(TOKEN, scopes=SCOPES)

    try:
        # Build the service
        service = build('sheets', 'v4', credentials=credentials)

    except HttpError as error:
        print(f"An error occurred: {error}")

    return service


def main():

    print("Starting the Bot...")
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    # commands
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))

    # messages
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_messages))

    # errors
    app.add_error_handler(error)

    # Polls the bot
    print("Polling...")
    app.run_polling(poll_interval=3)
    app.idle()


if __name__ == "__main__":
    main()