from unittest import TestCase
from flask.testing import FlaskClient
import config_test
import os
import json

import sharepoint_bot as bot

data_dict = {
    "help": {"wh_data": """
{"id": "Y2lzY29zcGFyazovL3VzL1dFQkhPT0svNWQyZmE3YjItYjU4OS00MzJlLWJjNzctNmM0YWY0YmQ0ODQy", "name": "Webhook for Bot", "targetUrl": "http://8ba6c15f.ngrok.io/", "resource": "messages", "event": "created", "orgId": "Y2lzY29zcGFyazovL3VzL09SR0FOSVpBVElPTi8xZWI2NWZkZi05NjQzLTQxN2YtOTk3NC1hZDcyY2FlMGUxMGY", "createdBy": "Y2lzY29zcGFyazovL3VzL1BFT1BMRS84NDY1ODlhMC05NGM2LTRjNTgtOWZjNC1mZDcyODUzNmJlM2U", "appId": "Y2lzY29zcGFyazovL3VzL0FQUExJQ0FUSU9OL0MzMmM4MDc3NDBjNmU3ZGYxMWRhZjE2ZjIyOGRmNjI4YmJjYTQ5YmE1MmZlY2JiMmM3ZDUxNWNiNGEwY2M5MWFh", "ownedBy": "creator", "status": "active", "created": "2019-09-02T14:48:47.718Z", "actorId": "Y2lzY29zcGFyazovL3VzL1BFT1BMRS82MzFlODQ0Mi02YTU3LTQ1ZTAtYjIyNy1jYWQ1Y2FkMmQ5MWQ", "data": {"id": "Y2lzY29zcGFyazovL3VzL01FU1NBR0UvNWMxM2VkYzAtY2Q5Mi0xMWU5LWI5MjAtYTdmYTZmMmEwOTky", "roomId": "Y2lzY29zcGFyazovL3VzL1JPT00vMjM3MTU3YjctMTZhZi0zOWQxLThkOGMtNTZiZWFiYTE1OTYz", "roomType": "direct", "personId": "Y2lzY29zcGFyazovL3VzL1BFT1BMRS82MzFlODQ0Mi02YTU3LTQ1ZTAtYjIyNy1jYWQ1Y2FkMmQ5MWQ", "personEmail": "jmartan@cisco.com", "created": "2019-09-02T15:00:10.524Z"}}
""",
        "expected_result": ["help created", "message created"],
    },
    "registration_form": {"wh_data": """
{"id": "Y2lzY29zcGFyazovL3VzL1dFQkhPT0svNWQyZmE3YjItYjU4OS00MzJlLWJjNzctNmM0YWY0YmQ0ODQy", "name": "Webhook for Bot", "targetUrl": "http://8ba6c15f.ngrok.io/", "resource": "messages", "event": "created", "orgId": "Y2lzY29zcGFyazovL3VzL09SR0FOSVpBVElPTi8xZWI2NWZkZi05NjQzLTQxN2YtOTk3NC1hZDcyY2FlMGUxMGY", "createdBy": "Y2lzY29zcGFyazovL3VzL1BFT1BMRS84NDY1ODlhMC05NGM2LTRjNTgtOWZjNC1mZDcyODUzNmJlM2U", "appId": "Y2lzY29zcGFyazovL3VzL0FQUExJQ0FUSU9OL0MzMmM4MDc3NDBjNmU3ZGYxMWRhZjE2ZjIyOGRmNjI4YmJjYTQ5YmE1MmZlY2JiMmM3ZDUxNWNiNGEwY2M5MWFh", "ownedBy": "creator", "status": "active", "created": "2019-09-02T14:48:47.718Z", "actorId": "Y2lzY29zcGFyazovL3VzL1BFT1BMRS82MzFlODQ0Mi02YTU3LTQ1ZTAtYjIyNy1jYWQ1Y2FkMmQ5MWQ", "data": {"id": "Y2lzY29zcGFyazovL3VzL01FU1NBR0UvMTgzMGFhZjAtY2RiOC0xMWU5LTk1NmYtNTdiMzBiOThhNTc3", "roomId": "Y2lzY29zcGFyazovL3VzL1JPT00vMjM3MTU3YjctMTZhZi0zOWQxLThkOGMtNTZiZWFiYTE1OTYz", "roomType": "direct", "personId": "Y2lzY29zcGFyazovL3VzL1BFT1BMRS82MzFlODQ0Mi02YTU3LTQ1ZTAtYjIyNy1jYWQ1Y2FkMmQ5MWQ", "personEmail": "jmartan@cisco.com", "created": "2019-09-02T19:30:17.503Z"}}
""",
        "expected_result": ['registration form created', 'message created', 'event form reference saved'],
    },
    "event_form": {"wh_data": """
{"id": "Y2lzY29zcGFyazovL3VzL1dFQkhPT0svNWQyZmE3YjItYjU4OS00MzJlLWJjNzctNmM0YWY0YmQ0ODQy", "name": "Webhook for Bot", "targetUrl": "http://8ba6c15f.ngrok.io/", "resource": "messages", "event": "created", "orgId": "Y2lzY29zcGFyazovL3VzL09SR0FOSVpBVElPTi8xZWI2NWZkZi05NjQzLTQxN2YtOTk3NC1hZDcyY2FlMGUxMGY", "createdBy": "Y2lzY29zcGFyazovL3VzL1BFT1BMRS84NDY1ODlhMC05NGM2LTRjNTgtOWZjNC1mZDcyODUzNmJlM2U", "appId": "Y2lzY29zcGFyazovL3VzL0FQUExJQ0FUSU9OL0MzMmM4MDc3NDBjNmU3ZGYxMWRhZjE2ZjIyOGRmNjI4YmJjYTQ5YmE1MmZlY2JiMmM3ZDUxNWNiNGEwY2M5MWFh", "ownedBy": "creator", "status": "active", "created": "2019-09-02T14:48:47.718Z", "actorId": "Y2lzY29zcGFyazovL3VzL1BFT1BMRS82MzFlODQ0Mi02YTU3LTQ1ZTAtYjIyNy1jYWQ1Y2FkMmQ5MWQ", "data": {"id": "Y2lzY29zcGFyazovL3VzL01FU1NBR0UvOGJlZjY3YjAtY2RiOC0xMWU5LWIzNjYtOWRlMTQ4MzdiNTVh", "roomId": "Y2lzY29zcGFyazovL3VzL1JPT00vMjM3MTU3YjctMTZhZi0zOWQxLThkOGMtNTZiZWFiYTE1OTYz", "roomType": "direct", "personId": "Y2lzY29zcGFyazovL3VzL1BFT1BMRS82MzFlODQ0Mi02YTU3LTQ1ZTAtYjIyNy1jYWQ1Y2FkMmQ5MWQ", "personEmail": "jmartan@cisco.com", "created": "2019-09-02T19:33:31.691Z"}}
""",
        "expected_result": ['event form created', 'message created', 'event form reference saved'],
    },
    "registration_form_submit": {"wh_data": """
{"id": "Y2lzY29zcGFyazovL3VzL1dFQkhPT0svYzE3ZGQ5NjUtNTY4Ni00NWQwLWIyNGItZmVjYzM4ZjNkMTAy", "name": "Webhook for Bot", "targetUrl": "http://bd46278f.ngrok.io/", "resource": "attachmentActions", "event": "created", "orgId": "Y2lzY29zcGFyazovL3VzL09SR0FOSVpBVElPTi8xZWI2NWZkZi05NjQzLTQxN2YtOTk3NC1hZDcyY2FlMGUxMGY", "createdBy": "Y2lzY29zcGFyazovL3VzL1BFT1BMRS84NDY1ODlhMC05NGM2LTRjNTgtOWZjNC1mZDcyODUzNmJlM2U", "appId": "Y2lzY29zcGFyazovL3VzL0FQUExJQ0FUSU9OL0MzMmM4MDc3NDBjNmU3ZGYxMWRhZjE2ZjIyOGRmNjI4YmJjYTQ5YmE1MmZlY2JiMmM3ZDUxNWNiNGEwY2M5MWFh", "ownedBy": "creator", "status": "active", "created": "2019-09-04T17:54:23.839Z", "actorId": "Y2lzY29zcGFyazovL3VzL1BFT1BMRS82MzFlODQ0Mi02YTU3LTQ1ZTAtYjIyNy1jYWQ1Y2FkMmQ5MWQ", "data": {"id": "Y2lzY29zcGFyazovL3VzL0FUVEFDSE1FTlRfQUNUSU9OLzI2MzBlNWUwLWQzMWEtMTFlOS1hNTM3LWNmNDA0ODBmNWMwMg", "type": "submit", "messageId": "Y2lzY29zcGFyazovL3VzL01FU1NBR0UvODJhM2M5MzAtZDA2Zi0xMWU5LWE3YTgtZDVlZmI2MDQ3OGM2", "personId": "Y2lzY29zcGFyazovL3VzL1BFT1BMRS82MzFlODQ0Mi02YTU3LTQ1ZTAtYjIyNy1jYWQ1Y2FkMmQ5MWQ", "roomId": "Y2lzY29zcGFyazovL3VzL1JPT00vNGU0YzYyNjAtYzJjNy0xMWU5LWJmMTItODc2YWY1ZWUzNjI0", "created": "2019-09-09T15:54:47.486Z"}}
""",
        "expected_result": ['form received', 'form data saved', 'response form created', 'message created', 'event form reference saved'],
    },
}

class BotFunctionalTest(TestCase):
    
    def setUp(self):
        print("\nsetup {}".format(self.__class__.__name__))
        bot.flask_app.testing = True
        self.client = bot.flask_app.test_client()
        
    def tearDown(self):
        print("tear down {}".format(self.__class__.__name__))
        
    def test_webhook_create(self):
        self.client.get("/")
        
"""
    def test_user_requests(self):
        for req_name in data_dict.keys():
            print("testing {} command".format(req_name))
            self.run_test_command(req_name)
            
    def test_event_command(self):
        self.run_test_command("event_form")
        
    def test_registration_submit(self):
        self.run_test_command("registration_form_submit")
            
    def run_test_command(self, command):
        action_result = self.client.post("/", data=data_dict[command]["wh_data"], content_type="application/json")
        result_data = json.loads(action_result.get_data(as_text=True))
        self.assertListEqual(result_data, data_dict[command]["expected_result"])
"""
        
        
if __name__ == "__main__":
    unittest.main()
    
"""
admin story 1

privileged user creates 1-1 with the bot. Bot detects the communication, verifies the user's OrgId.
Bot checks the O365 authorization status.
Bot creates a form with the information:
1. auth status
2. O365 authorization link for (re)authorization


admin story 2

privileged user adds bot to a Space. Bot detects the new membership and sends a form asking for action:
1. create a file structure for project (enter name of the project/Sharepoint)
2. monitor user's activities (add, remove) and set the permissions
3. moderator request
4. only local users allowed

Once the file structure is created, the privileged user is responsible for linking the structure to the Space. Ideally bot should do this automatically.
If moderation is requested, the bot sets itself and the requestor to the moderator role.
If "only local users", the bot checks person's Org. If a user from a    different Org is added, the bot removes him.


user story 1

user adds/removes a person to the Space.
Bot adjusts the person's permissions

"""
