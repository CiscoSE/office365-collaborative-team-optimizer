from dotenv import load_dotenv
import json
import os
import requests
from webexteamsbot import TeamsBot
from webexteamsbot.models import Response
import pprint

load_dotenv()

# Retrieve required details from environment variables
bot_email = os.getenv("OCTO_BOT_EMAIL")
bot_id = os.getenv("OCTO_BOT_ID")
teams_token = os.getenv("OCTO_BOT_TOKEN")
bot_url = os.getenv("OCTO_BOT_URL")
bot_app_name = os.getenv("OCTO_BOT_APP_NAME")

o365_token = os.getenv("O365_TOKEN")

# Create a Bot Object
bot = TeamsBot(
    bot_app_name,
    teams_bot_token=teams_token,
    teams_bot_url=bot_url,
    teams_bot_email=bot_email,
    webhook_resource_event=[{"resource": "messages", "event": "created"},
                            {"resource": "attachmentActions", "event": "created"}]
)


# Create a custom bot greeting function returned when no command is given.
# The default behavior of the bot is to return the '/help' command response
def greeting(incoming_msg):
    # Loopkup details about sender
    sender = bot.teams.people.get(incoming_msg.personId)

    # Create a Response object and craft a reply in Markdown.
    response = Response()
    response.markdown = "Hello {}, I'm a chat bot. ".format(sender.firstName)
    response.markdown += "See what I can do by asking for **/help**."
    return response


def show_calendar_meeting_card(incoming_msg):
    if bot.teams.rooms.get(incoming_msg.roomId).type == "direct":
        return "I'd love to meet, but you should really use me in a collaborative space!"
    attachment = '''
    {
        "contentType": "application/vnd.microsoft.card.adaptive",
        "content": {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.0",
            "body": [
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "ColumnSet",
                            "spacing": "Padding",
                            "columns": [
                                {
                                    "type": "Column",
                                    "width": 10,
                                    "items": [
                                        {
                                            "type": "Image",
                                            "altText": "",
                                            "url": "https://upload.wikimedia.org/wikipedia/commons/thumb/d/df/Microsoft_Office_Outlook_%282018%E2%80%93present%29.svg/1200px-Microsoft_Office_Outlook_%282018%E2%80%93present%29.svg.png",
                                            "size": "Small",
                                            "horizontalAlignment": "Center"
                                        }
                                    ]
                                },
                                {
                                    "type": "Column",
                                    "width": 50,
                                    "items": [
                                        {
                                            "type": "RichTextBlock",
                                            "inlines": [
                                                {
                                                    "type": "TextRun",
                                                    "text": "Schedule a Meeting",
                                                    "size": "Large"
                                                }
                                            ],
                                            "horizontalAlignment": "Left",
                                            "spacing": "None"
                                        }
                                    ],
                                    "verticalContentAlignment": "Center"
                                }
                            ]
                        },
                        {
                            "type": "Input.Text",
                            "placeholder": "Placeholder text",
                            "isVisible": false,
                            "id": "card_type",
                            "value": "calendar_meeting"
                        }
                    ],
                    "verticalContentAlignment": "Center"
                },
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Subject"
                        },
                        {
                            "type": "Input.Text",
                            "placeholder": "Meeting Subject",
                            "value": "'''
    attachment = attachment + bot.teams.rooms.get(incoming_msg.roomId).title + '''",
                            "id": "subject"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Attendees",
                            "spacing": "Padding"
                        },
                        {
                            "type": "Input.ChoiceSet",
                            "placeholder": "Placeholder text",
                            "choices": [
    '''
    attendees_list = ''''''
    for member in bot.teams.memberships.list(roomId=incoming_msg.roomId):
        if "@webex.bot" not in member.personEmail:
            attendees_list += '''
            {
                "title": "''' + member.personDisplayName + '''",
                "value": "''' + member.personDisplayName + ''':''' + member.personEmail + '''"
            },'''
    attachment = attachment + attendees_list[:-1] + '''
                            ],
                            "style": "expanded",
                            "isMultiSelect": true,
                            "id": "attendees"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Meeting Start"
                        },
                        {
                            "type": "ColumnSet",
                            "columns": [
                                {
                                    "type": "Column",
                                    "width": "30",
                                    "items": [
                                        {
                                            "type": "Container",
                                            "items": [
                                                {
                                                    "type": "TextBlock",
                                                    "text": "Date"
                                                },
                                                {
                                                    "type": "Input.Date",
                                                    "id": "startDate",
                                                    "value": "YYYY-MM-DD"
                                                }
                                            ]
                                        }
                                    ]
                                },
                                {
                                    "type": "Column",
                                    "width": "20",
                                    "items": [
                                        {
                                            "type": "Container",
                                            "items": [
                                                {
                                                    "type": "TextBlock",
                                                    "text": "Time"
                                                },
                                                {
                                                    "type": "Input.Time",
                                                    "id": "startTime",
                                                    "value": "HH:MM"
                                                }
                                            ]
                                        }
                                    ]
                                },
                                {
                                    "type": "Column",
                                    "width": "20",
                                    "items": [
                                        {
                                            "type": "Container",
                                            "items": [
                                                {
                                                    "type": "TextBlock",
                                                    "text": "TimeZone"
                                                },
                                                {
                                                    "type": "Input.ChoiceSet",
                                                    "placeholder": "EST",
                                                    "choices": [
                                                        {
                                                            "title": "EST",
                                                            "value": "Eastern Standard Time"
                                                        },
                                                        {
                                                            "title": "PST",
                                                            "value": "Pacific Standard Time"
                                                        }
                                                    ],
                                                    "id": "startTimeZone",
                                                    "spacing": "Padding"
                                                }
                                            ],
                                            "horizontalAlignment": "Left"
                                        }
                                    ]
                                }
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "text": "Meeting End"
                        },
                        {
                            "type": "ColumnSet",
                            "columns": [
                                {
                                    "type": "Column",
                                    "width": "30",
                                    "items": [
                                        {
                                            "type": "Container",
                                            "items": [
                                                {
                                                    "type": "TextBlock",
                                                    "text": "Date"
                                                },
                                                {
                                                    "type": "Input.Date",
                                                    "id": "endDate",
                                                    "value": "YYYY-MM-DD"
                                                }
                                            ]
                                        }
                                    ]
                                },
                                {
                                    "type": "Column",
                                    "width": "20",
                                    "items": [
                                        {
                                            "type": "Container",
                                            "items": [
                                                {
                                                    "type": "TextBlock",
                                                    "text": "Time"
                                                },
                                                {
                                                    "type": "Input.Time",
                                                    "id": "endTime",
                                                    "value": "HH:MM"
                                                }
                                            ]
                                        }
                                    ]
                                },
                                {
                                    "type": "Column",
                                    "width": "20",
                                    "items": [
                                        {
                                            "type": "Container",
                                            "items": [
                                                {
                                                    "type": "TextBlock",
                                                    "text": "TimeZone"
                                                },
                                                {
                                                    "type": "Input.ChoiceSet",
                                                    "placeholder": "EST",
                                                    "choices": [
                                                        {
                                                            "title": "EST",
                                                            "value": "America/New_York"
                                                        },
                                                        {
                                                            "title": "PST",
                                                            "value": "America/Los_Angeles"
                                                        }
                                                    ],
                                                    "id": "endTimeZone",
                                                    "spacing": "Padding"
                                                }
                                            ],
                                            "horizontalAlignment": "Left"
                                        }
                                    ]
                                }
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "text": "Webex",
                            "spacing": "Padding"
                        },
                        {
                            "type": "Input.Toggle",
                            "title": "Include",
                            "value": "false",
                            "wrap": false,
                            "id": "webex"
                        }
                    ],
                    "separator": true,
                    "spacing": "Medium"
                }
            ],
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "Submit"
                }
            ]
        }
    }
    '''
    backupmessage = "This is an example using Adaptive Cards."

    c = create_message_with_attachment(incoming_msg.roomId,
                                       msgtxt=backupmessage,
                                       attachment=json.loads(attachment))
    print(c)
    return ""


def show_onenote_card(incoming_msg):
    attachment = '''
    {
        "contentType": "application/vnd.microsoft.card.adaptive",
        "content": {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.0",
            "body": [
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "ColumnSet",
                            "columns": [
                                {
                                    "type": "Column",
                                    "width": 10,
                                    "items": [
                                        {
                                            "type": "Image",
                                            "altText": "",
                                            "url": "https://upload.wikimedia.org/wikipedia/commons/thumb/1/1c/Microsoft_Office_OneNote_%282018%E2%80%93present%29.svg/1200px-Microsoft_Office_OneNote_%282018%E2%80%93present%29.svg.png",
                                            "size": "Small",
                                            "horizontalAlignment": "Center"
                                        }
                                    ]
                                },
                                {
                                    "type": "Column",
                                    "width": 50,
                                    "items": [
                                        {
                                            "type": "RichTextBlock",
                                            "inlines": [
                                                {
                                                    "type": "TextRun",
                                                    "text": "Create a Shared Note",
                                                    "size": "Large"
                                                }
                                            ],
                                            "horizontalAlignment": "Left",
                                            "spacing": "None"
                                        }
                                    ],
                                    "verticalContentAlignment": "Center"
                                }
                            ]
                        },
                        {
                            "type": "Input.Text",
                            "placeholder": "Placeholder text",
                            "isVisible": false,
                            "id": "card_type",
                            "value": "onenote_note"
                        }
                    ],
                    "verticalContentAlignment": "Center"
                },
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Subject"
                        },
                        {
                            "type": "Input.Text",
                            "value": "'''
    attachment = attachment + bot.teams.rooms.get(incoming_msg.roomId).title + '''",
                            "id": "subject"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Recipients",
                            "spacing": "Padding"
                        },
                        {
                            "type": "Input.ChoiceSet",
                            "placeholder": "Placeholder text",
                            "choices": [
                            '''
    attendees_list = ''''''
    for member in bot.teams.memberships.list(roomId=incoming_msg.roomId):
        if "@webex.bot" not in member.personEmail:
            attendees_list += '''
                {
                    "title": "''' + member.personDisplayName + '''",
                    "value": "''' + member.personDisplayName + ''':''' + member.personEmail + '''"
                },'''
    attachment = attachment + attendees_list[:-1] + '''
                            ],
                            "style": "expanded",
                            "isMultiSelect": true,
                            "id": "attendees"
                        }
                    ],
                    "separator": true,
                    "spacing": "Medium"
                }
            ],
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "Submit"
                }
            ]
        }
    }
    '''
    backupmessage = "This is an example using Adaptive Cards."

    c = create_message_with_attachment(incoming_msg.roomId,
                                       msgtxt=backupmessage,
                                       attachment=json.loads(attachment))
    print(c)
    return ""


def create_outlook_meeting(meeting_info):
    headers = {"Authorization": "Bearer " + o365_token}
    url = 'https://graph.microsoft.com/v1.0/me/events'
    payload = {}
    payload["subject"] = meeting_info["subject"]
    payload["body"] = {"contentType": "HTML", "content": "Meeting body"}
    startDateTime = meeting_info["startDate"] + "T" + meeting_info["startTime"] + ":00.000"
    payload["start"] = {"dateTime": startDateTime, "timeZone": meeting_info["startTimeZone"]}
    endDateTime = meeting_info["endDate"] + "T" + meeting_info["endTime"] + ":00.000"
    payload["end"] = {"dateTime": endDateTime, "timeZone": meeting_info["endTimeZone"]}
    if meeting_info["webex"]:
        payload["location"] = {"displayName": "@webex"}
    payload["attendees"] = []
    for attendee in meeting_info["attendees"].split(","):
        attendee_payload = {"emailAddress": {"address": attendee.split(":")[1], "name": attendee.split(":")[0]}}
        payload["attendees"].append(attendee_payload)
    r = requests.post(url, json=payload, headers=headers)
    return r.status_code


# An example of how to process card actions
def handle_cards(api, incoming_msg):
    """
    Sample function to handle card actions.
    :param api: webexteamssdk object
    :param incoming_msg: The incoming message object from Teams
    :return: A text or markdown based reply
    """
    m = get_attachment_actions(incoming_msg["data"]["id"])
    card_type = m["inputs"]["card_type"]
    if card_type == "calendar_meeting":
        print("Meeting info sent: ")
        print(m["inputs"])
        status_code = create_outlook_meeting(m["inputs"])
        print(status_code)
        if status_code == 201:
            return "Meeting scheduled successfully!"
        else:
            return "Error occurred during scheduling."
    elif card_type == "onenote_note":
        # TODO: Add OneNote Creation and Sharing
        return "Shared OneNote available here: URL"


def create_message_with_attachment(rid, msgtxt, attachment):
    headers = {
        'content-type': 'application/json; charset=utf-8',
        'authorization': 'Bearer ' + teams_token
    }

    url = 'https://api.ciscospark.com/v1/messages'
    data = {"roomId": rid, "attachments": [attachment], "markdown": msgtxt}
    response = requests.post(url, json=data, headers=headers)
    return response.json()


def get_attachment_actions(attachmentid):
    headers = {
        'content-type': 'application/json; charset=utf-8',
        'authorization': 'Bearer ' + teams_token
    }

    url = 'https://api.ciscospark.com/v1/attachment/actions/' + attachmentid
    response = requests.get(url, headers=headers)
    return response.json()


# Set the bot greeting.
bot.set_greeting(greeting)

# Add new commands to the bot.
bot.add_command('attachmentActions', '*', handle_cards)
bot.add_command("/meeting", "Schedule a Meeting", show_calendar_meeting_card)
bot.add_command("/note", "Create a Shared OneNote", show_onenote_card)

# Every bot includes a default "/echo" command.  You can remove it, or any
# other command with the remove_command(command) method.
bot.remove_command("/echo")

if __name__ == "__main__":
    # Run Bot
    bot.run(host="0.0.0.0", port=5000)