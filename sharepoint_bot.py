#!/bin/#!/usr/bin/env python3

import os
import sys
import re
import uuid
import logging
from dotenv import load_dotenv, find_dotenv
load_dotenv(find_dotenv())

from urllib.parse import urlparse, urlunparse, quote, parse_qsl, urlencode

from distutils.util import strtobool

from webexteamssdk import WebexTeamsAPI, ApiError, AccessToken
webex_api = WebexTeamsAPI()

from ddb_single_table_obj import DDB_Single_Table

import json, requests
from datetime import datetime, timedelta, timezone
import time
from flask import Flask, request, redirect, url_for

from zappa.asynchronous import task

import concurrent.futures
import signal

from O365 import Account, FileSystemTokenBackend
from o365_db_token_storage import DBTokenBackend

flask_app = Flask(__name__)
flask_app.config["DEBUG"] = True
requests.packages.urllib3.disable_warnings()

logger = logging.getLogger()

ddb = None

WEBEX_SCOPE = ["spark-compliance:events_read", "spark-compliance:memberships_read",
    "spark-compliance:memberships_write", "spark-compliance:messages_read", "spark-compliance:messages_write",
    "spark-compliance:rooms_read", "spark-compliance:team_memberships_read", "spark-compliance:team_memberships_write",
    "spark-compliance:teams_read", "spark:people_read"] # "spark:rooms_read", "spark:kms"
STATE_CHECK = "webex is great" # integrity test phrase
SAFE_TOKEN_DELTA = 86400 # safety seconds before access token expires - renew if smaller
DEFAULT_FORM_MSG = "Toto je formulář. Zobrazíte si ho ve webovém klientovi Webex Teams nebo v desktopové aplikaci."
DEFAULT_AVATAR_URL= "http://bit.ly/SparkBot-512x512"
avatar_url = DEFAULT_AVATAR_URL
webhook_url = None
bot_name = None
bot_email = None
bot_id = None

GRAPH_SCOPE = ["Files.ReadWrite", "offline_access", "Sites.Manage.All", "Sites.ReadWrite.All", "User.Read"]
# SP_SCOPE = ["Sites.FullControl.All", "Sites.Manage.All", "Sites.Read.All", "Sites.ReadWrite.All", "TermStore.Read.All", "TermStore.ReadWrite.All", "User.Read.All", "User.ReadWrite.All"]

# GRAPH_SCOPE = ["offline_access", "https://graph.microsoft.com/Files.ReadWrite", "https://graph.microsoft.com/offline_access", "https://graph.microsoft.com/Sites.Manage.All", "https://graph.microsoft.com/Sites.ReadWrite.All", "https://graph.microsoft.com/User.Read"]
SP_SCOPE = ["https://microsoft.sharepoint-df.com/Sites.Manage.All", "https://microsoft.sharepoint-df.com/Sites.Read.All", "https://microsoft.sharepoint-df.com/Sites.ReadWrite.All", "https://microsoft.sharepoint-df.com/User.Read.All", "https://microsoft.sharepoint-df.com/User.ReadWrite.All"]
EXTRA_SCOPE = ["https://microsoft.sharepoint-df.com/TermStore.Read.All", "https://microsoft.sharepoint-df.com/TermStore.ReadWrite.All"]

O365_SCOPE = GRAPH_SCOPE

# GRAPH_SCOPE = ["basic", "sharepoint_dl"]
# o365_client_id = os.getenv("O365_CLIENT_ID")
# o365_client_secret = os.getenv("O365_CLIENT_SECRET")
# o365_credentials = (o365_client_id, o365_client_secret)

# threading part
thread_executor = concurrent.futures.ThreadPoolExecutor()

class AccessTokenAbs(AccessToken):
    def __init__(self, access_token_json):
        super().__init__(access_token_json)
        if not "expires_at" in self._json_data.keys():
            self._json_data["expires_at"] = str((datetime.now(timezone.utc) + timedelta(seconds = self.expires_in)).timestamp())
        flask_app.logger.debug("Access Token expires in: {}s, at: {}".format(self.expires_in, self.expires_at))
        if not "refresh_token_expires_at" in self._json_data.keys():
            self._json_data["refresh_token_expires_at"] = str((datetime.now(timezone.utc) + timedelta(seconds = self.refresh_token_expires_in)).timestamp())
        flask_app.logger.debug("Refresh Token expires in: {}s, at: {}".format(self.refresh_token_expires_in, self.refresh_token_expires_at))
        
    @property
    def expires_at(self):
        return self._json_data["expires_at"]
        
    @property
    def refresh_token_expires_at(self):
        return self._json_data["refresh_token_expires_at"]
        
def sigterm_handler(_signo, _stack_frame):
    "When sysvinit sends the TERM signal, cleanup before exiting."

    flask_app.logger.info("Received signal {}, exiting...".format(_signo))
    
    thread_executor._threads.clear()
    concurrent.futures.thread._threads_queues.clear()
    sys.exit(0)
    
signal.signal(signal.SIGTERM, sigterm_handler)
signal.signal(signal.SIGINT, sigterm_handler)

SAMPLE_AUTH_FORM = {
    "contentType": "application/vnd.microsoft.card.adaptive",
    "content": {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.2",
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
                                "text": "O365 Authentication Valid",
                            },
                            {
                                "type": "TextBlock",
                                "text": "O365 Authorization Status for Org",
                            },
                            {
                                "type": "TextBlock",
                                "text": "O365 Authorization Status for individual",
                            },
                        ]
                    },
                    {
                        "type": "Column",
                        "width": "stretch",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": "{{o365_authenticated}}",
                            },
                            {
                                "type": "TextBlock",
                                "text": "{{o365_auth_status_org}}",
                            },
                            {
                                "type": "TextBlock",
                                "text": "{{o365_auth_status_individual}}",
                            },
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
                        "url": "{{o365_auth_url}}",
                    }
                ]
            }
        ]
    }
}

SAMPLE_SPACE_FORM = {
    "contentType": "application/vnd.microsoft.card.adaptive",
    "content": {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.2",
        "body": [
            {
                "type": "TextBlock",
                "text": "Sharepoint",
                "weight": "Bolder",
            },
            {
                "type": "Input.ChoiceSet",
                "placeholder": "Vyberte Sharepoint",
                "value": "sharepoint",
                "style": "compact",
                "id": "sharepoint_select",
                "choices": [],
            },
            {
                "type": "Input.ChoiceSet",
                "placeholder": "Placeholder text",
                "value": "project",
                "style": "expanded",
                "id": "structure",
                "choices": [
                    {
                        "title": "Projekt",
                        "value": "project"
                    },
                    {
                        "title": "Nabídka",
                        "value": "proposal"
                    },
                    {
                        "title": "Smlouva",
                        "value": "contract"
                    },
                ]
            },
            {
                "type": "Input.Toggle",
                "title": "Moderovat",
                "value": "false",
                "wrap": True,
                "id": "moderate",
            },
            {
                "type": "Input.Toggle",
                "title": "Pouze lokální uživatelé",
                "value": "false",
                "wrap": False,
                "id": "local_only",
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Vytvořit srukturu",
                # "id": "submit",
            }
        ]
    }
}

PROJECT_FOLDER_STRUCTURE = ["Zadani", "Zapisy", {"Faze": ["Faze 1", "Faze 2"]}]
NO_FOLDER_STRUCTURE = []

FOLDER_STRUCTURES = {
    "project": PROJECT_FOLDER_STRUCTURE,
    "proposal": PROJECT_FOLDER_STRUCTURE,
    "contract": NO_FOLDER_STRUCTURE
}

# identification mapping in DB between form and submitted data
FORM_DATA_MAP = {
    "SAMPLE_SPACE_FORM": "SHAREPOINT_SPACE",
}

# type of form data to be saved in database
FORM_DATA_TO_SAVE = ["SHAREPOINT_SPACE"]

FORM_TEMPLATE_MAP = {
}

TARGET_FORM_MAP = {
}


def nested_replace( structure, original, new ):
    if type(structure) == list:
        return [nested_replace( item, original, new) for item in structure]

    if type(structure) == dict:
        return {key : nested_replace(value, original, new)
                     for key, value in structure.items() }

    return structure.replace("{{"+original+"}}", new)

def save_tokens(user_email, tokens):
    flask_app.logger.debug("AT timestamp: {}".format(tokens.expires_at))
    token_record = {
        "access_token": tokens.access_token,
        "refresh_token": tokens.refresh_token,
        "expires_at": tokens.expires_at,
        "refresh_token_expires_at": tokens.refresh_token_expires_at
    }
    ddb.save_db_record(user_email, "TOKENS", tokens.expires_at, **token_record)
    
def get_tokens_for_user(user_email):
    db_tokens = ddb.get_db_record(user_email, "TOKENS")
    
    if db_tokens:
        tokens = AccessTokenAbs(db_tokens)
        flask_app.logger.debug("Got tokens: {}".format(tokens))
        ## TODO: check if token is not expired, generate new using refresh token if needed
        return tokens
    else:
        flask_app.logger.error("No tokens for user {}.".format(user_email))
        return None

def refresh_tokens_for_user(user_email):
    tokens = get_tokens_for_user(user_email)
    integration_api = WebexTeamsAPI()
    client_id = os.getenv("WEBEX_INTEGRATION_CLIENT_ID")
    client_secret = os.getenv("WEBEX_INTEGRATION_CLIENT_SECRET")
    try:
        new_tokens = AccessTokenAbs(integration_api.access_tokens.refresh(client_id, client_secret, tokens.refresh_token).json_data)
        save_tokens(user_email, new_tokens)
        flask_app.logger.info("Tokens refreshed for user {}".format(user_email))
    except ApiError as e:
        flask_app.logger.error("Client Id and Secret loading error: {}".format(e))
        return "Error refreshing an access token. Client Id and Secret loading error: {}".format(e)
        
    return new_tokens
    
def create_webhook(target_url):    
    flask_app.logger.debug("Create new webhook to URL: {}".format(target_url))
    
    webhook_name = "Webhook for Bot {}".format(bot_email)
    event = "created"
    resource_events = {
        "messages": ["created"],
        "memberships": ["created", "deleted"],
        "attachmentActions": ["created"]
    }
    status = None
        
    try:
        check_webhook = webex_api.webhooks.list()
        for webhook in check_webhook:
            flask_app.logger.debug("Deleting webhook {}, '{}', App Id: {}".format(webhook.id, webhook.name, webhook.appId))
            try:
                if not flask_app.testing:
                    webex_api.webhooks.delete(webhook.id)
            except ApiError as e:
                flask_app.logger.error("Webhook {} delete failed: {}.".format(webhook.id, e))
    except ApiError as e:
        flask_app.logger.error("Webhook list failed: {}.".format(e))
        
    for resource, events in resource_events.items():
        for event in events:
            try:
                if not flask_app.testing:
                    webex_api.webhooks.create(name=webhook_name, targetUrl=target_url, resource=resource, event=event)
                status = True
                flask_app.logger.debug("Webhook for {}/{} was successfully created".format(resource, event))
            except ApiError as e:
                flask_app.logger.error("Webhook create failed: {}.".format(e))
            
    return status

def group_info(bot_name):
    return "Nezapomeňte, že je třeba mne oslovit '@{}'".format(bot_name)

def greetings(personal=True):
    
    greeting_msg = """
Tady bude formulář pro vytvoření Space & SP site.
"""
    if not personal:
        greeting_msg += " " + group_info(bot_name)

    return greeting_msg

def help_me(personal=True):

    greeting_msg = """
Dummy help.
"""
    if not personal:
        greeting_msg += group_info(bot_name)

    return greeting_msg

def is_room_direct(room_id):
    try:
        res = webex_api.rooms.get(room_id)
        return res.type == "direct"
    except ApiError as e:
        flask_app.logger.error("Room info request failed: {}".format(e))
        return False

# Flask part of the code

"""
1. initialize database table if needed
2. start event checking thread
"""
@flask_app.before_first_request
def before_first_request():
    global bot_email, bot_name, bot_id, avatar_url, ddb
    
    try:
        me = webex_api.people.me()
        bot_email = me.emails[0]
        bot_name = me.displayName
        bot_id = me.id
        avatar_url = me.avatar
    except ApiError as e:
        avatar_url = DEFAULT_AVATAR_URL
        flask_app.logger.error("Status code: {}, {}".format(e.status_code, e.message))

    if ("@sparkbot.io" not in bot_email) and ("@webex.bot" not in bot_email):
        flask_app.logger.error("""
You have provided access token which does not belong to a bot ({}).
Please review it and make sure it belongs to your bot account.
Do not worry if you have lost the access token.
You can always go to https://developer.ciscospark.com/apps.html 
URL and generate a new access token.""".format(bot_email))

    ddb = DDB_Single_Table()

@task
def handle_webhook_event(webhook):
    action_list = []
    if webhook["data"].get("personEmail") != bot_email:
        flask_app.logger.info(json.dumps(webhook))
        pass
    if webhook["resource"] == "memberships":
        msg = ""
        if webhook["data"]["personEmail"] == bot_email:
            if webhook["event"] == "created":
                personal_room = is_room_direct(webhook["data"]["roomId"])
                if personal_room:
                    flask_app.logger.debug("I was invited to a new 1-1 communication")
                    msg = markdown=greetings(personal_room)
                    action_list.append("invited to a new 1-1 communication")
                else:
                    flask_app.logger.debug("I was invited to a new group Space")
                    msg = "Proběhne kontrola, zda je na Space navázán Sharepoint. Pokud není, bude tu formulář na potvrzení, že se má vytvořit Sharepoint site a prolinkovat na Space."
                    action_list.append("invited to a group Space")
            elif webhook["event"] == "deleted":
                flask_app.logger.info("I was removed from a Space")
                action_list.append("bot removed from a Space")
            else:
                flask_app.logger.info("unhandled membership event '{}'".format(webhook["event"]))
        else:
            if webhook["event"] == "created":
                msg += "Kontrola, zda uživatel {} může být členem space  \n".format(webhook["data"]["personEmail"])
                if user_allowed_to_space(webhook["data"]["roomId"], webhook["data"]["personId"]):
                    msg += "členství povoleno  \n"
                    flask_app.logger.debug("user {} added to a Space".format(webhook["data"]["personEmail"]))
                    msg += "Budou nastavena přístupová práva nového uživatele {}\n".format(webhook["data"]["personEmail"])
                    result = invite_new_user_to_folder(webhook["data"]["roomId"], webhook["data"]["personId"])
                    msg += "  \n{}".format(result.get("message", "no result"))
                    action_list.append("user {} added to a Space".format(webhook["data"]["personEmail"]))
                else:
                    remove_user_from_space(webhook["data"]["roomId"], webhook["data"]["personId"])
                    msg += "uživatel odstraněn  \n"
            elif webhook["event"] == "deleted":
                flask_app.logger.info("user {} removed from a Space".format(webhook["data"]["personEmail"]))
                msg = "Přístupová práva uživatele {} budou zrušena.".format(webhook["data"]["personEmail"])
                result = remove_user_access(webhook["data"]["roomId"], webhook["data"]["personEmail"])
                msg += "  \n{}".format(result.get("message", "no result"))
                action_list.append("user {} removed from a Space".format(webhook["data"]["personEmail"]))
            else:
                flask_app.logger.info("unhandled membership event '{}'".format(webhook["event"]))
        if msg != "":
            webex_api.messages.create(roomId=webhook["data"]["roomId"], markdown=msg)            
        return action_list
        
    msg = ""
    attach = []
    target_dict = {"roomId": webhook["data"]["roomId"]}
    form_type = None
    out_messages = [] # additional messages apart of the standard response
    if webhook["resource"] == "messages":
        if webhook["data"]["personEmail"] == bot_email:
            flask_app.logger.debug("Ignoring my own message")
        else:
            in_msg = webex_api.messages.get(webhook["data"]["id"])
            in_msg_low = in_msg.text.lower()
            in_msg_low = in_msg_low.replace(bot_name.lower() + " ", "") # remove bot"s name from message test to avoid command conflict
            myUrlParts = urlparse(request.url)

            if "help" in in_msg_low:
                personal_room = is_room_direct(webhook["data"]["roomId"])
                msg = help_me(personal_room)
                action_list.append("help created")
            elif "authorize" in in_msg_low:
                msg = DEFAULT_FORM_MSG
                full_auth_uri = secure_scheme(myUrlParts.scheme) + "://" + myUrlParts.netloc + url_for("o365_auth", **{"state": webhook["data"].get("personId")})
                flask_app.logger.debug("O365 authorization URL: {}".format(full_auth_uri))
                
                # org_info = webex_api.organizations.get(webhook["orgId"])
                
                form = nested_replace(SAMPLE_AUTH_FORM, "o365_auth_url", full_auth_uri)
                # form = nested_replace(form, "org_id", webhook["orgId"])
                # form = nested_replace(form, "org_name", org_info.displayName)
                
                account = get_o365_account(webhook["data"]["personId"], webhook["orgId"])
                
                form = nested_replace(form, "o365_authenticated", str(account.is_authenticated))
                
                org_auth_info = account.con.token_backend.secondary_id == webhook["orgId"]
                person_auth_info = account.con.token_backend.owner_id == webhook["data"]["personId"]
                form = nested_replace(form, "o365_auth_status_org", str(org_auth_info))
                form = nested_replace(form, "o365_auth_status_individual", str(person_auth_info))

                attach.append(form)
                form_type = "SAMPLE_FORM"
                action_list.append("sample form created")
            elif "space" in in_msg_low:
                msg = DEFAULT_FORM_MSG
                
                sp_sites = get_sharepoint_sites(webhook["data"]["personId"], webhook["orgId"])
                sp_selection = create_site_selection(sp_sites)
                
                form = SAMPLE_SPACE_FORM
                form["content"]["body"][1]["choices"] = sp_selection
                
                attach.append(form)
                form_type = "SAMPLE_SPACE_FORM"
                action_list.append("space form created")
            elif "notify" in in_msg_low:
                msg = DEFAULT_FORM_MSG
                attach.append(EVENT_NOTIFICATION_FORM_SAMPLE)
                form_type = "EVENT_NOTIFICATION_FORM"
                action_list.append("event notification form created")
            elif "event" in in_msg_low:
                msg, event_form = create_event_form(webhook["data"]["personId"])
                form_type = "EVENT_FORM"
                attach.append(event_form)
                action_list.append("event form created")
            elif "confirm" in in_msg_low:
                msg = DEFAULT_FORM_MSG
                attach.append(REG_CONFIRMATION_FORM)
                action_list.append("registration confirmation form created")
                
            if msg != "" or len(attach) > 0:
                out_messages.append({"message": msg, "attachments": attach, "target": target_dict, "form_type": form_type})
    elif webhook["resource"] == "attachmentActions":
        try:
            spk_headers = {
                "Accept": "application/json",
                "Content-Type": "application/json; charset=utf-8",
                "Authorization": "Bearer " + webex_api.access_token
            }

            in_attach = requests.get(webex_api.base_url+"attachment/actions/"+webhook["data"]["id"], headers=spk_headers)
            in_attach_dict = in_attach.json()
            flask_app.logger.debug("Form received: {}".format(in_attach_dict))
            # flask_app.logger.debug("Form metadata: \nApp Id: {}\nMsg Id: {}\nPerson Id: {}".format(webhook["appId"], webhook["data"]["messageId"], webhook["data"]["personId"]))
            action_list.append("form received")
            in_attach_dict["orgId"] = webhook["orgId"] # orgId is present only in original message, not in attachement
            
            ## TODO: detect the type of form (event, registration, management, ...)
            #  check if it exists 
            #  save form_data with proper identification
            #  generate the subsequent form if needed
            
            parent_record = ddb.get_db_records_by_secondary_key(webhook["data"]["messageId"])[0]
            flask_app.logger.debug("Parent record: {}".format(parent_record))
            form_type = parent_record.get("data")
            form_data_type = FORM_DATA_MAP.get(form_type)
            flask_app.logger.debug("Received form type: {} -> {}".format(form_type, form_data_type))
            form_params = {}
            if form_data_type in FORM_DATA_TO_SAVE:
                additional_data = {}
                primary_key = webhook["data"]["messageId"]
                secondary_key = webhook["data"]["personId"]
                # form_saved = save_form_data(primary_key, secondary_key, in_attach_dict, form_data_type, **additional_data)
                action_list.append("form data complete")
            target_form_type_list = TARGET_FORM_MAP.get(form_data_type)
            if target_form_type_list is not None:
                # we should create another form(s)
                for target_form_type in target_form_type_list:
                    target_dict, response_form = create_form(in_attach_dict, target_form_type, webhook["data"]["personId"])
                    if response_form is not None:
                        out_messages.append({"message": DEFAULT_FORM_MSG, "attachments": response_form.copy(), "target": target_dict.copy(), "form_type": target_form_type, "form_params": form_params})
                        action_list.append("response form created")
            else:
                # this is the leaf form, check the data content only
                target_dict, msg, attach = handle_response(in_attach_dict, form_data_type, parent_record)
                form_type = "UNKNOWN_FORM"
                out_messages.append({"message": msg, "target": target_dict, "attachments": attach})
        except ApiError as e:
            flask_app.logger.error("Form read failed: {}.".format(e))
            action_list.append("form read failed")
                        
    if len(out_messages) > 0:
        for msg_dict in out_messages:
            try:
                target_dict = msg_dict["target"]
                msg = msg_dict["message"]
                attach = msg_dict.get("attachments", [])
                form_type = msg_dict.get("form_type", "UNKNOWN_FORM")
                form_params = msg_dict.get("form_params", {})
                flask_app.logger.debug("Send to space: {}\n\nmarkdown: {}\n\nattach: {}".format(target_dict, msg, attach))
                res_msg = webex_api.messages.create(**target_dict, markdown=msg, attachments=attach)
                flask_app.logger.debug("Message created: {}".format(res_msg.json_data))
                action_list.append("message created")
                if len(attach) > 0 and form_type is not None:
                    save_form_info(webhook["data"]["personId"], res_msg.id, form_type, form_params) # sender is the owner
                    action_list.append("event form reference saved")
                else:
                    flask_app.logger.debug("Not saving, attach len: {}, form type: {}".format(len(attach), form_type))
            except ApiError as e:
                flask_app.logger.error("Message create failed: {}.".format(e))
                action_list.append("message create failed")

    return json.dumps(action_list)

def handle_response(attachment_data, form_data_type, parent_info):
    response = ""
    attachments = []
    target = None
    if attachment_data.get("roomId"):
        target = {"roomId": attachment_data.get("roomId")}
        target_space = attachment_data.get("roomId")
    if target is None and attachment_data.get("personId"):
        target = {"toPersonId": attachment_data.get("personId")}
        target_space = attachment_data.get("personId")
    inputs = attachment_data.get("inputs")
    if form_data_type == "SHAREPOINT_SPACE":
        sender_data = webex_api.people.get(attachment_data.get("personId"))
        response = "creating SharePoint site upon request from {}".format(sender_data.emails[0])
        # TODO: send request to O365, create a site
        result = create_sharepoint_site(sender_data.id, sender_data.orgId, target_space)
        response += "  \n{}".format(result.get("message", "no results"))
        if result["created"]:
            # TODO create folder structure
            # sp_site = result["sharepoint_site"]
            sp_site = inputs.get("sharepoint_select", None)
            
            response += "creating folder structure {} under {}".format(result["room_name"], sp_site)
            result = create_folder_structure(sp_site, inputs["structure"], target_space, sender_data.orgId, sender_data.id)
            response += "  \n{}".format(result.get("message", "no results"))
            # TODO set permissions for existing users
            response += "invite existing Space users to the folder"
            result = invite_existing_users_to_folder(result["folder"], target_space)
            response += "  \n{}".format(result.get("message", "no results"))
            # TODO set moderator
            if strtobool(inputs.get("moderate", "False")):
                response += "  \nsetting moderators to {}, {}".format(bot_email, sender_data.emails[0])
                set_moderator(attachment_data.get("roomId"), [sender_data.id])
                response += "  \nmoderators set"
            # TODO start monitoring space memberships
            response += "  \nstart monitoring this Space membership"
            result = start_monitoring_space_membership(sender_data.id, sender_data.orgId, target_space, sp_site, strtobool(inputs.get("local_only", "False")))
            response += "  \n{}".format(result.get("message", "no results"))
            # TODO: internal API request to WxT for site linking
            response += "  \nlinking SharePoint site to this space"
            result = link_folder_to_space(sp_site, target_space)
            response += "  \n{}".format(result.get("message", "no results"))
    
    return target, response, attachments
    
def create_sharepoint_site(personId, orgId, roomId):
    result = {"created": False}
    
    account = get_o365_account(personId, orgId)
    
    if account.is_authenticated:
        room_info = webex_api.rooms.get(roomId)
        result["message"] = "Sharepoint site \"{}\" created".format(room_info.title)
        result["sharepoint_site"] = None
        result["created"] = True
        result["room_name"] = room_info.title
    else:
        result["message"] = "O365 account is not authenticated. Please authenticate first in 1-1 communication with me."
        
    return result
    
def create_folder_structure(sp_site, structure_name, roomId, orgId, personId):
    result = {"created": False, "message": "no action"}
    
    o365_account = get_o365_account(personId, orgId)
    
    subsite = get_sharepoint_site(sp_site, o365_account)
    if subsite:
        result["message"] = " found site {}".format(subsite.display_name)
        space_info = webex_api.rooms.get(roomId)
        
        doc_library = subsite.get_default_document_library()
        root_folder = doc_library.get_root_folder()
        walk_result = walk_folder_structure(root_folder, space_info.title, FOLDER_STRUCTURES.get(structure_name, []))
        result["message"] = walk_result["message"]
        result["folder"] = walk_result["folder"]
    
    return result

def get_sharepoint_site(sp_site_name, o365_account):
    sharepoint = o365_account.sharepoint()
    sp_root = sharepoint.get_root_site()
    subsites = sp_root.get_subsites()
    for subsite in subsites:
        if subsite.name == sp_site_name:
            flask_app.logger.debug("found site {}".format(subsite.display_name))
            return subsite    
    
def walk_folder_structure(parent_folder, parent_name, structure, level=0):
    indent = level*"-"
    result = {"message": " create {}{}".format(indent, parent_name)}
    flask_app.logger.debug("parent action - create folder {}".format(parent_name))
    new_folder = parent_folder.create_child_folder(parent_name)
    if level == 0:
        result["folder"] = new_folder
    level += 1
    indent = level*"-"
    if isinstance(structure, dict):
        structure = [structure]
    if isinstance(structure, (list, tuple)):
        for item in structure:
            if isinstance(item, dict):
                for key, val in item.items():
                    walk_res = walk_folder_structure(new_folder, key, val, level=level)
                    result["message"] += walk_res["message"]
            else:
                new_folder.create_child_folder(item)
    else:
        new_folder.create_child_folder(str(structure))
        
    return result
    
def invite_existing_users_to_folder(folder, roomId):
    result = {}
    memberships = webex_api.memberships.list(roomId = roomId)
    invite_list = []
    for membership in memberships:
        if not membership.personId == bot_id:
            invite_list.append(membership.personEmail)
    
    if invite_list:
        flask_app.logger.debug("inviting {} to the folder \"{}\"".format(invite_list, folder.name))
        folder.share_with_invite(invite_list, share_type="edit")
        result["message"] = " pozváni uživatelé {}".format(invite_list)
        
        return result
        
def invite_new_user_to_folder(roomId, personId):
    result = {}
    space_info = webex_api.rooms.get(roomId)
    monitored, data = space_is_monitored(roomId)
    if not monitored:
        flask_app.logger.error("space {} is not monitored".format(roomId))
        return "space \"{}\" není monitorován".format(space_info.title)
    person_info = webex_api.people.get(personId)
    
    o365_account = get_o365_account(data["requested_by"], data["org_id"])
    
    subsite = get_sharepoint_site(data["sharepoint_site"], o365_account)
    if subsite:
        folder = find_folder(subsite, space_info.title)
        if folder:            
            folder.share_with_invite(person_info.emails[0], share_type="edit")
            
            result["message"] = "user {} invited to the folder".format(person_info.emails[0])
        else:
            result["message"] = "folder {} not found in document library of {}".format(space_info.title, sp_site)
    
    return result

def remove_user_from_folder(roomId, personId):    
    result = {}
    space_info = webex_api.rooms.get(roomId)
    monitored, data = space_is_monitored(roomId)
    if not monitored:
        flask_app.logger.error("space {} is not monitored".format(roomId))
        return "space \"{}\" není monitorován".format(space_info.title)
        
    person_info = webex_api.people.get(personId)
    
    o365_account = get_o365_account(data["requested_by"], data["org_id"])
    
    subsite = get_sharepoint_site(data["sharepoint_site"], o365_account)
    if subsite:
        folder = find_folder(subsite, space_info.title)
        if folder:            
            folder.share_with_invite(person_info.emails[0], share_type="edit")
            
            result["message"] = "user {} invited to the folder".format(person_info.emails[0])
        else:
            result["message"] = "folder {} not found in document library of {}".format(space_info.title, sp_site)

    return result
    
def find_folder(subsite, folder_name):
    doc_library = subsite.get_default_document_library()
    root_folder = doc_library.get_root_folder()
    subfolders = root_folder.get_child_folders()
    for folder in subfolders:
        if folder.name == folder_name:
            return folder
    
        
def start_monitoring_space_membership(personId, orgId, roomId, sp_site, own_users_only=False):
    result = {"created": False}

    result["message"] = "no action"
    
    monitoring_data = {"own_users_only": own_users_only, "org_id": orgId, "requested_by": personId, "sharepoint_site": sp_site}
    
    ddb.save_db_record(roomId, "MONITORED_SPACE", "", **monitoring_data)
    result["created"] = True
    result["message"] = "space {} monitoring started, {} allowed".format(roomId, "own users only" if own_users_only else "all users")
    
    """
    informace pro DB: roomId, personId (kdo akci vyvolal), povoleni pouze interni uzivatele
    """
    
    return result
    
def space_is_monitored(roomId):
    db_record = ddb.get_db_record(roomId, "MONITORED_SPACE")
    if db_record is not None:
        flask_app.logger.debug("Space {} is monitored, data: {}".format(roomId, db_record["data"]))
        return True, db_record
    else:
        flask_app.logger.debug("Space {} is not monitored.".format(roomId))
        return False, None
        
def set_moderator(roomId, moderator_list):
    my_memberships = webex_api.memberships.list(roomId=roomId, personId=bot_id)
    for membership in my_memberships:
        flask_app.logger.debug("setting primary moderator flag for user {} in membership {}".format(bot_id, membership.id))
        webex_api.memberships.update(membership.id, isModerator=True)
        
    for moderator in moderator_list:
        other_memberships = webex_api.memberships.list(roomId=roomId, personId=moderator)
        for membership in other_memberships:
            flask_app.logger.debug("setting other moderator flag for user {} in membership {}".format(bot_id, membership.id))
            webex_api.memberships.update(membership.id, isModerator=True)
    
def link_folder_to_space(sp_site, roomId):
    result = {"created": False}

    result["message"] = "no action"
    
    """
    informace pro DB: roomId, personId (kdo akci vyvolal), orgId (kvuli autentizaci O365), SP_data (informace o Sharepointu - site, folder name/link)
    """
    
    return result
    
def user_allowed_to_space(roomId, personId):
    result = True
    monitored, data = space_is_monitored(roomId)
    if monitored and data["own_users_only"]:
        user_info = webex_api.people.get(personId)
        result = user_info.orgId == data["org_id"]
        
    return result
    
def remove_user_from_space(roomId, personId):
    memberships = webex_api.memberships.list(roomId = roomId, personId = personId)
    
    for membership in memberships:
        flask_app.logger.debug("deleting membership {}".format(membership.id))
        webex_api.memberships.delete(membership.id)
    
    return True
    
def add_user_access(roomId, person_email):
    result = {}
    
    return result
    
def remove_user_access(roomId, person_email):
    result = {}
    
    return result
        
def get_sharepoint_sites(user_id, org_id):
    result = {}
    o365_account = get_o365_account(user_id, org_id)
    
    sharepoint = o365_account.sharepoint()
    sp_root = sharepoint.get_root_site()
    sites_list = sp_root.get_subsites()
    
    for site in sites_list:
        result[site.name] = site.display_name
        
    flask_app.logger.debug("SP site dict: {}".format(result))
    
    return result
    
def create_site_selection(site_dict):
    result = []
    for key, value in site_dict.items():
        flask_app.logger.debug("create site selection: {} -> {}".format(key, value))
        result.append({"title": value, "value": key})
        
    return result
    
def get_o365_account(user_id, org_id):
    o365_client_id = os.getenv("O365_CLIENT_ID")
    o365_client_secret = os.getenv("O365_CLIENT_SECRET")
    o365_credentials = (o365_client_id, o365_client_secret)

    token_backend = DBTokenBackend(user_id, bot_email, org_id)
    account = Account(o365_credentials, token_backend=token_backend)
    
    flask_app.logger.debug("account {} is{} authenticated".format(user_id, "" if account.is_authenticated else " not"))

    return account
    
def get_o365_account_noauth():
    o365_client_id = os.getenv("O365_CLIENT_ID")
    o365_client_secret = os.getenv("O365_CLIENT_SECRET")
    o365_credentials = (o365_client_id, o365_client_secret)

    account = Account(o365_credentials)
    
    flask_app.logger.debug("get O365 account without authentication")

    return account

def link_sharepoint_site(personId, orgId, roomId, siteInfo):
    pass
    
def save_form_info(creator_id, form_data_id, form_type, params={}):
    return ddb.save_db_record(creator_id, form_data_id, form_type, **params)
    
def get_form_info(form_data_id):
    return ddb.get_db_records_by_secondary_key(form_data_id)[0]
    
def delete_form_info(form_data_id):
    return ddb.delete_db_records_by_secondary_key(form_data_id)
    
def save_form_data(primary_key, secondary_key, registration_data, data_type, **kwargs):
    inputs = registration_data.get("inputs", {})
    optional_data = {**inputs, **kwargs}
    return ddb.save_db_record(primary_key, secondary_key, data_type, **optional_data)
    
def get_form_data(form_data_id):
    return ddb.get_db_record_list(form_data_id)
    
def delete_form_data_for_user(form_id, user_id):
    return ddb.delete_db_record(form_id, user_id)
    
def secure_scheme(scheme):
    return re.sub(r"^http$", "https", scheme)
        
@flask_app.route("/", methods=["GET", "POST"])
def spark_webhook():
    if request.method == "POST":
        webhook = request.get_json(silent=True)
        
        handle_webhook_event(webhook)        
    elif request.method == "GET":
        message = "<center><img src=\"{0}\" alt=\"{1}\" style=\"width:256; height:256;\"</center>" \
                  "<center><h2><b>Congratulations! Your <i style=\"color:#ff8000;\">{1}</i> bot is up and running.</b></h2></center>".format(avatar_url, bot_name)
                  
        message += "<center><b>I'm hosted at: <a href=\"{0}\">{0}</a></center>".format(request.url)
        if webhook_url is None:
            res = create_webhook(request.url)
            if res is True:
                message += "<center><b>New webhook created sucessfully</center>"
            else:
                message += "<center><b>Tried to create a new webhook but failed, see application log for details.</center>"

        return message
        
    flask_app.logger.debug("Webhook handling done.")
    return("OK")

@flask_app.route("/startup")
def startup():
    return "Hello World!"

"""
Webex Teams OAuth grant flow start
"""
@flask_app.route("/authorize", methods=["GET"])
def authorize():
    myUrlParts = urlparse(request.url)
    full_redirect_uri = secure_scheme(myUrlParts.scheme) + "://" + myUrlParts.netloc + url_for("manager")
    flask_app.logger.debug("Authorize redirect URL: {}".format(full_redirect_uri))

    client_id = os.getenv("WEBEX_INTEGRATION_CLIENT_ID")
    redirect_uri = quote(full_redirect_uri, safe="")
    scope = WEBEX_SCOPE
    scope_uri = quote(" ".join(scope), safe="")
    join_url = webex_api.base_url+"authorize?client_id={}&response_type=code&redirect_uri={}&scope={}&state={}".format(client_id, redirect_uri, scope_uri, STATE_CHECK)

    return redirect(join_url)
    
"""
OAuth grant flow redirect url
generate access and refresh tokens using "code" generated in OAuth grant flow
after user successfully authenticated to Webex

See: https://developer.webex.com/blog/real-world-walkthrough-of-building-an-oauth-webex-integration
https://developer.webex.com/docs/integrations
"""   
@flask_app.route("/manager", methods=["GET"])
def manager():
    if request.args.get("error"):
        return request.args.get("error_description")
        
    input_code = request.args.get("code")
    check_phrase = request.args.get("state")
    flask_app.logger.debug("Authorization request \"state\": {}, code: {}".format(check_phrase, input_code))

    myUrlParts = urlparse(request.url)
    full_redirect_uri = secure_scheme(myUrlParts.scheme) + "://" + myUrlParts.netloc + url_for("manager")
    flask_app.logger.debug("Manager redirect URI: {}".format(full_redirect_uri))
    
    try:
        client_id = os.getenv("WEBEX_INTEGRATION_CLIENT_ID")
        client_secret = os.getenv("WEBEX_INTEGRATION_CLIENT_SECRET")
        tokens = AccessTokenAbs(webex_api.access_tokens.get(client_id, client_secret, input_code, full_redirect_uri).json_data)
        flask_app.logger.debug("Access info: {}".format(tokens))
    except ApiError as e:
        flask_app.logger.error("Client Id and Secret loading error: {}".format(e))
        return "Error issuing an access token. Client Id and Secret loading error: {}".format(e)
        
    webex_integration_api = WebexTeamsAPI(access_token=tokens.access_token)
    try:
        user_info = webex_integration_api.people.me()
        flask_app.logger.debug("Got user info: {}".format(user_info))
        save_tokens(user_info.emails[0], tokens)
        
        ## TODO: add periodic access token refresh
    except ApiError as e:
        flask_app.logger.error("Error getting user information: {}".format(e))
        return "Error getting your user information: {}".format(e)
        
    return redirect(url_for("authdone"))
    
"""
OAuth proccess done
"""
@flask_app.route("/authdone", methods=["GET"])
def authdone():
    ## TODO: post the information & help, maybe an event creation form to the 1-1 space with the user
    return "Thank you for providing the authorization. You may close this browser window."
    
"""
Manual token refresh of a single user. Not needed if the thread is running.
"""
@flask_app.route("/tokenrefresh", methods=["GET"])
def token_refresh():
    user_id = request.args.get("user_id")
    if user_id is None:
        return "Please provide a user id"
    
    return refresh_token_for_user(user_id)
    
def refresh_token_for_user(user_id):
    tokens = get_tokens_for_user(user_id)
    integration_api = WebexTeamsAPI()
    client_id = os.getenv("WEBEX_INTEGRATION_CLIENT_ID")
    client_secret = os.getenv("WEBEX_INTEGRATION_CLIENT_SECRET")
    try:
        new_tokens = AccessTokenAbs(integration_api.access_tokens.refresh(client_id, client_secret, tokens.refresh_token).json_data)
        save_tokens(user_id, new_tokens)
    except ApiError as e:
        flask_app.logger.error("Client Id and Secret loading error: {}".format(e))
        return "Error refreshing an access token. Client Id and Secret loading error: {}".format(e)
        
    return "token refresh for user {} done".format(user_id)

"""
Manual token refresh of all users. Not needed if the thread is running.
"""
@flask_app.route("/tokenrefreshall", methods=["GET"])
def token_refresh_all():
    results = ""
    user_tokens = ddb.get_db_records_by_secondary_key("TOKENS")
    for token in user_tokens:
        flask_app.logger.debug("Refreshing: {} token".format(token["pk"]))
        results += refresh_token_for_user(token["pk"])+"\n"
    
    return results

# TODO: manual query of events API
@flask_app.route("/queryevents", methods=["GET"])
def query_events():
    results = ""
    
    return results
    
"""
O365 OAuth grant flow
"""
@flask_app.route('/o365auth')
def o365_auth():
    my_state = request.args.get("state", "local")
    flask_app.logger.debug("input state: {}".format(my_state))
    
    myUrlParts = urlparse(request.url)
    full_redirect_uri = secure_scheme(myUrlParts.scheme) + "://" + myUrlParts.netloc + url_for("o365_do_auth")
    flask_app.logger.debug("Authorize redirect URL: {}".format(full_redirect_uri))

    # callback = quote(full_redirect_uri, safe="")
    callback = full_redirect_uri
    scopes = O365_SCOPE
    
    account = get_o365_account_noauth()

    url, o365_state = account.con.get_authorization_url(requested_scopes=scopes, redirect_uri=callback)
    
    # replace "state" parameter injected by O365 object
    o365_auth_parts = urlparse(url)
    o365_query = dict(parse_qsl(o365_auth_parts.query))
    o365_query["state"] = my_state
    new_o365_auth_parts = o365_auth_parts._replace(query = urlencode(o365_query))
    new_o365_url = urlunparse(new_o365_auth_parts)
    
    flask_app.logger.debug("O365 auth URL: {}".format(new_o365_url))

    # the state must be saved somewhere as it will be needed later
    # my_db.store_state(state) # example...

    return redirect(new_o365_url)

@flask_app.route('/o365doauth')
def o365_do_auth():
    # token_backend = FileSystemTokenBackend(token_path='.', token_filename='o365_token.txt')
    my_state = request.args.get("state", "local")
    flask_app.logger.debug("O365 state: {}".format(my_state))
    
    person_data = webex_api.people.get(my_state)
    flask_app.logger.debug("O365 login requestor data: {}".format(person_data))
    
    account = get_o365_account(my_state, person_data.orgId)
    
    # retreive the state saved in auth_step_one
    # my_saved_state = my_db.get_state()  # example...

    # rebuild the redirect_uri used in auth_step_one
    myUrlParts = urlparse(request.url)
    full_redirect_uri = secure_scheme(myUrlParts.scheme) + "://" + myUrlParts.netloc + url_for("o365_do_auth")
    flask_app.logger.debug("Authorize doauth redirect URL: {}".format(full_redirect_uri))

    # callback = quote(full_redirect_uri, safe="")
    callback = full_redirect_uri
    req_url = re.sub(r"^http:", "https:", request.url)
    
    flask_app.logger.debug("URL: {}".format(req_url))

    result = account.con.request_token(req_url, 
                                       state=my_state,
                                       redirect_uri=callback)
                                       
    flask_app.logger.info("O365 authentication status: {}".format("authenticated" if account.is_authenticated else "not authenticated"))
    
    if account.is_authenticated:
        webex_api.messages.create(toPersonId=my_state, text="Autentizace O365 proběhla úspěšně.")
    
    # if result is True, then authentication was succesful 
    #  and the auth token is stored in the token backend
    if result:
        return redirect(url_for("authdone"))
    else:
        return "Authentication failed: {}".format(result)

"""
Independent thread startup, see:
https://networklore.com/start-task-with-flask/
"""
def start_runner():
    def start_loop():
        not_started = True
        while not_started:
            logger.info('In start loop')
            try:
                r = requests.get('http://127.0.0.1:5050/startup')
                if r.status_code == 200:
                    logger.info('Server started, quiting start_loop')
                    not_started = False
                logger.debug("Status code: {}".format(r.status_code))
            except:
                logger.info('Server not yet started')
            time.sleep(2)

    logger.info('Started runner')
    thread_executor.submit(start_loop)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument('-v', '--verbose', action='count', help="Set logging level by number of -v's, -v=WARN, -vv=INFO, -vvv=DEBUG")
    
    args = parser.parse_args()
    if args.verbose:
        if args.verbose > 2:
            logging.basicConfig(level=logging.DEBUG)
        elif args.verbose > 1:
            logging.basicConfig(level=logging.INFO)
        if args.verbose > 0:
            logging.basicConfig(level=logging.WARN)
            
    flask_app.logger.info("Logging level: {}".format(logging.getLogger(__name__).getEffectiveLevel()))
    
    bot_identity = webex_api.people.me()
    flask_app.logger.info("Bot \"{}\"\nUsing database: {} - {}".format(bot_identity.displayName, os.getenv("DYNAMODB_ENDPOINT_URL"), os.getenv("DYNAMODB_TABLE_NAME")))
    
    start_runner()
    flask_app.run(host="0.0.0.0", port=5050)
