import os
import sys

from O365 import Account, FileSystemTokenBackend
from o365_db_token_storage import DBTokenBackend

# Postman
"""
https://www.ktskumar.com/2017/01/access-sharepoint-online-using-postman/

id: 365c6d9f-0006-41b0-bdbf-7d75e6767fe1
secret: 9049NaoFblfs8sSs8Jz/yREpGWk1ShQPwVhKTBKQRvY=
<AppPermissionRequests AllowAppOnlyPolicy="true">
    <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web" Right="Read" />
    <AppPermissionRequest Scope="https://microsoft.sharepoint-df.com" Right="Sites.FullControl.All" />
    <AppPermissionRequest Scope="https://microsoft.sharepoint-df.com" Right="Sites.Manage.All" />
    <AppPermissionRequest Scope="https://microsoft.sharepoint-df.com" Right="Sites.Read.All" />
    <AppPermissionRequest Scope="https://microsoft.sharepoint-df.com" Right="Sites.ReadWrite.All" />
    <AppPermissionRequest Scope="https://microsoft.sharepoint-df.com" Right="Sites.TermStore.Read.All" />
    <AppPermissionRequest Scope="https://microsoft.sharepoint-df.com" Right="TermStore.ReadWrite.All" />
    <AppPermissionRequest Scope="https://microsoft.sharepoint-df.com" Right="User.Read.All" />
    <AppPermissionRequest Scope="https://microsoft.sharepoint-df.com" Right="User.ReadWrite.All" />
</AppPermissionRequests>

authenticate:
Bearer realm="0be1f1ba-21e4-4c34-a193-1b44e66afda7",client_id="00000003-0000-0ff1-ce00-000000000000",trusted_issuers="00000001-0000-0000-c000-000000000000@*,D3776938-3DBA-481F-A652-4BEDFCAB7CD8@*,https://sts.windows.net/*/,00000003-0000-0ff1-ce00-000000000000@90140122-8516-11e1-8eff-49304924019b",authorization_uri="https://login.windows.net/common/oauth2/authorize"

{
    "token_type": "Bearer",
    "expires_in": "86399",
    "not_before": "1590151376",
    "expires_on": "1590238076",
    "resource": "00000003-0000-0ff1-ce00-000000000000/collabrocks.sharepoint.com@0be1f1ba-21e4-4c34-a193-1b44e66afda7",
    "access_token": "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IkN0VHVoTUptRDVNN0RMZHpEMnYyeDNRS1NSWSIsImtpZCI6IkN0VHVoTUptRDVNN0RMZHpEMnYyeDNRS1NSWSJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvY29sbGFicm9ja3Muc2hhcmVwb2ludC5jb21AMGJlMWYxYmEtMjFlNC00YzM0LWExOTMtMWI0NGU2NmFmZGE3IiwiaXNzIjoiMDAwMDAwMDEtMDAwMC0wMDAwLWMwMDAtMDAwMDAwMDAwMDAwQDBiZTFmMWJhLTIxZTQtNGMzNC1hMTkzLTFiNDRlNjZhZmRhNyIsImlhdCI6MTU5MDE1MTM3NiwibmJmIjoxNTkwMTUxMzc2LCJleHAiOjE1OTAyMzgwNzYsImlkZW50aXR5cHJvdmlkZXIiOiIwMDAwMDAwMS0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDBAMGJlMWYxYmEtMjFlNC00YzM0LWExOTMtMWI0NGU2NmFmZGE3IiwibmFtZWlkIjoiMzY1YzZkOWYtMDAwNi00MWIwLWJkYmYtN2Q3NWU2NzY3ZmUxQDBiZTFmMWJhLTIxZTQtNGMzNC1hMTkzLTFiNDRlNjZhZmRhNyIsIm9pZCI6IjE4NjgzODhkLTI0OTYtNDBmOC1hYTNiLWI1Zjk1OGYwYTg4MyIsInN1YiI6IjE4NjgzODhkLTI0OTYtNDBmOC1hYTNiLWI1Zjk1OGYwYTg4MyIsInRydXN0ZWRmb3JkZWxlZ2F0aW9uIjoiZmFsc2UifQ.BxzpzMcO_JvXUpDH52k98gmAqFdXO1nOo8uBDLHbsqYd32GjI6PizlJLfOXcrI3M9c-A6p2YwSTqcEdwnBetQGwkvB5O9mexyILQzeib-HQWhIQjYFcWJB4QIZatXZnzbHlbtJmVpqBRzUsxSWhbsEdLdW59pkN7mn4bUfFJoQukVBRSUo9cKzLHt9AX3CYkLksnOCE6iYi7xxayxwwAtOGdFFQFE2UJOMBLIlvcwG4CKHJj3bgxvWJnpDvV7220MeepP9nQriKJ7faKmqbz2Cuk2JuzNtJlVS-PhGRLtuezccwamHG-PveDLshFCX7x1gYNqJitsonn0M3TIys2Yw"
}

create site:
https://collabrocks.sharepoint.com/_api/web/webinfos/add
https://graph.microsoft.com/beta/sites/collabrocks.sharepoint.com/add

accept/content-type: application/json;odata=verbose

{ "parameters":
{ "__metadata":
{ "type": "SP.WebInfoCreationInformation" },
"Url":"test01",
"Title":"Test Site 01",
"Description":"Test 01 space linked to Webex Teams",
"Language":"1033",
"WebTemplate":"STS#0",
"UseUniquePermissions":false
}
}

oauth
https://login.windows.net/common/oauth2/authorize?client_id=365c6d9f-0006-41b0-bdbf-7d75e6767fe1&response_type=code&redirect_uri=https://localhost&scope=Sites.FullControl.All
"""

# GRAPH_SCOPE = ["Files.ReadWrite", "offline_access", "Sites.Manage.All", "Sites.ReadWrite.All", "User.Read"]
# SP_SCOPE = ["Sites.FullControl.All", "Sites.Manage.All", "Sites.Read.All", "Sites.ReadWrite.All", "TermStore.Read.All", "TermStore.ReadWrite.All", "User.Read.All", "User.ReadWrite.All"]

GRAPH_SCOPE = ["https://graph.microsoft.com/Files.ReadWrite", "https://graph.microsoft.com/offline_access", "https://graph.microsoft.com/Sites.Manage.All", "https://graph.microsoft.com/Sites.ReadWrite.All", "https://graph.microsoft.com/User.Read"]
SP_SCOPE = ["https://microsoft.sharepoint-df.com/Sites.FullControl.All", "https://microsoft.sharepoint-df.com/Sites.Manage.All", "https://microsoft.sharepoint-df.com/Sites.Read.All", "https://microsoft.sharepoint-df.com/Sites.ReadWrite.All", "https://microsoft.sharepoint-df.com/User.Read.All", "https://microsoft.sharepoint-df.com/User.ReadWrite.All"]
EXTRA_SCOPE = ["https://microsoft.sharepoint-df.com/TermStore.Read.All", "https://microsoft.sharepoint-df.com/TermStore.ReadWrite.All"]

O365_SCOPE = GRAPH_SCOPE + SP_SCOPE

# GRAPH_SCOPE = ["basic", "sharepoint_dl"]
o365_client_id = os.getenv("O365_CLIENT_ID")
o365_client_secret = os.getenv("O365_CLIENT_SECRET")
o365_credentials = (o365_client_id, o365_client_secret)

bot_id = "jmartan.sp.test@webex.bot"
org_id = "Y2lzY29zcGFyazovL3VzL09SR0FOSVpBVElPTi8xZWI2NWZkZi05NjQzLTQxN2YtOTk3NC1hZDcyY2FlMGUxMGY"

# token_backend = FileSystemTokenBackend(token_path='.', token_filename='o365_token.txt')
token_backend = DBTokenBackend(owner_id=bot_id, storage_id=bot_id, secondary_id=org_id)
account = Account(o365_credentials, token_backend=token_backend)

SAMPLE_AUTH_FORM = {
    "contentType": "application/vnd.microsoft.card.adaptive",
    "content": {
        "type": "AdaptiveCard",
        "version": "1.0",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "body": [
            {
                "type": "TextBlock",
                "text": "Webex Space & Sharepoint",
                "weight": "Bolder"
            },
            {
                "type": "ColumnSet",
                "columns": [
                    {
                        "type": "Column",
                        "width": "stretch",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": "Org. Id:"
                            },
                            {
                                "type": "TextBlock",
                                "text": "Org. Name"
                            },
                            {
                                "type": "TextBlock",
                                "text": "O365 Authorization Status for Org"
                            },
                            {
                                "type": "TextBlock",
                                "text": "O365 Authorization Status for individual"
                            }
                        ]
                    },
                    {
                        "type": "Column",
                        "width": "stretch",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": "1-2-3"
                            },
                            {
                                "type": "TextBlock",
                                "text": "noname"
                            },
                            {
                                "type": "TextBlock",
                                "text": "unknown"
                            },
                            {
                                "type": "TextBlock",
                                "text": "unknown"
                            }
                        ]
                    }
                ]
            },
            {
                "type": "ActionSet",
                "actions": [
                    {
                        "type": "Action.OpenUrl",
                        "title": "Authorize O365",
                        "id": "o365_auth",
                        "url": "{{url}}"
                    }
                ]
            }
        ]
    }
}

PROJECT_FOLDER_STRUCTURE = ["Zadani", "Zapisy", {"Faze": [{"Faze 1": "Subfaze 1-1"}, {"Faze 2": ["2-1", "2-2"]}, "Faze 3"]}]

def nested_replace( structure, original, new ):
    if type(structure) == list:
        return [nested_replace( item, original, new) for item in structure]

    if type(structure) == dict:
        return {key : nested_replace(value, original, new)
                     for key, value in structure.items() }

    return structure.replace("{{"+original+"}}", new)
    
def walk_structure(parent_name, structure, action, action_args={}, level=0):
    indent = level*"-"
    action("parent action {}{}".format(indent, parent_name), **action_args)
    level += 1
    indent = level*"-"
    if isinstance(structure, dict):
        structure = [structure]
    if isinstance(structure, (list, tuple)):
        for item in structure:
            if isinstance(item, dict):
                for key, val in item.items():
                    walk_structure(key, val, action, action_args=action_args, level=level)
            else:
                action("child action {}{}".format(indent, item))
    else:
        action("child action {}{}".format(indent, structure))
