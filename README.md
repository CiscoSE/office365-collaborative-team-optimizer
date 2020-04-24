# office365-collaborative-team-optimizer

[![Watch the video](https://img.youtube.com/vi/Ho7OThndf28/hqdefault.jpg)](https://youtu.be/Ho7OThndf28)

*An integration between Webex Teams and Office 365 for an enhanced collaboration experience!*


## Business/Technical Challenge

Cisco employees - as well as Cisco customers - have heavily adopted Webex Teams and Office 365 for collaboration and running day-to-day business. With two critical tools for collaboration teams must constantly mangage both platforms simultaneously in order to effectively work together or important information can be lost. For example, quick messages are often exchanged via an email within Outlook or a conversation on Webex Teams. Meetings might be set up by opening a personal room through Webex Teams or sending a calendar invite within Outlook. Meeting notes are then scattered among team members and their various physical notebooks, OneNote entries, and Webex rooms.

While most people have learned to be adept at juggling both platforms, what if we could further unify the two products and create a truly collaborative experience?


## Proposed Solution

With Office 365 Collaborative Team Optimizer (OCTO), we want to unify the two major collaboration tools that Cisco and its customers' buisinesses utilize daily. OCTO is a Webex Teams bot that a team can add into any room and begin using immediately. Once added, a simple command will tell OCTO what the team would like to accomplish such as scheduling a meeting or creating a OneNote for the team.

Example 1: If the conversation is trending towards setting up a meeting, simply tell OCTO "meeting" and a card will pop up that automatically includes the name of the room as the subject and auto-includes everyone within the room. Pick a date and time and then just like that - a meeting to tackle what ever was being discussed has been set up from within Webex Teams.

Example 2: Perhaps a meeting has just ended. Everyone took their notes and want to compile them together for simplicity going forward. Rather just copying and pasting into an email thread or into a Teams room, tell OCTO "note" and a card will pop up to create a new shared OneNote with everyone in the room. Now the team can have a single document with all their notes in it.


### Cisco Products Technologies/ Services

Our solution will levegerage the following technologies:

* [Cisco Webex Teams](http://cisco.com/go/webexteams)
* [Office 365](https://www.microsoft.com/office365)


## Team Members

* Bradford Ingersoll <bingerso@cisco.com> - US Commercial East
* Eric Scott <eriscott@cisco.com> - US Public Sector Central East


## Solution Components

* [Python 3.7](https://www.python.org/)
* [webexteamsbot](https://github.com/hpreston/webexteamsbot)
* [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0)


## Usage

To use OCTO, simply setup the bot in your environment. Once deployed, add OCTO to any collaborative space with 2 or more people. When you want to utilize its features, just reference the bot and the appropriate command. For example:

```
@OCTO /meeting
```

A card will appear and you can proceed to schedule your team meeting!


## Installation

1. Clone this repository

```
git clone https://github.com/CiscoSE/office365-collaborative-team-optimizer.git
```

2. Configure a .env file in the **code** directory with all the necessary environment variables
```
OCTO_BOT_EMAIL=
OCTO_BOT_ID=
OCTO_BOT_TOKEN=
OCTO_BOT_URL=
OCTO_BOT_APP_NAME=
O365_TOKEN=
```

**NOTE: For full functionality, an OAUTH integration should be added to the bot that enables OCTO to act on behalf of each individual user. Given constraints regarding InfoSec within our production environment, a temporary token was used to demonstrate the functionality.**

3. Create a virtual environment for the python code
```
python -m venv venv
```

4. Activate the newly created virtual environment
```
source venv/bin/activate
```

5. Install the required modules into the virtual environment
```
pip install -r requirements.txt
```

6. Run the bot.py on a publicly acessible URL (using ngrok or other means)
```
python bot.py
```

## Documentation

* [webexteamsbot](https://github.com/hpreston/webexteamsbot)
* [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0)


## License

Provided under Cisco Sample Code License, for details see [LICENSE](./LICENSE.md)

## Code of Conduct

Our code of conduct is available [here](./CODE_OF_CONDUCT.md)

## Contributing

See our contributing guidelines [here](./CONTRIBUTING.md)
