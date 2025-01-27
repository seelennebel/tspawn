#script configuration files
import config
import default_owners

#libraries
import csv
import sys
import requests
from msal import PublicClientApplication

class team_spawner:

    #check if config.py contains MS Azure credentials
    def __init__(self): 
        for value in config.credentials:
            if(config.credentials[value] == ""):
                raise Exception("ERROR IN CONFIG FILE")

    #class instance variables definitions
        self.client_id = config.credentials["client_id"]
        self.tenant_id = config.credentials["tenant_id"]
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.scopes = ["User.Read.All", "Team.Create"]
        self.app = PublicClientApplication(self.client_id, authority=self.authority)
        self.access_token = None
        self.team_id = ""

    #parse team id string
    def parse_team_id(team_id):
        id = team_id[8:]
        id = id[:-2]
        return id

    # receive access token
    def get_access_token(self):
        result = self.app.acquire_token_interactive(self.scopes, port=8080)
        if "access_token" in result:
            self.access_token = result["access_token"]
            return result["access_token"]
        else:
            raise Exception("COULD NOT ACQUIRE ACCESS TOKEN")

    # create a team with a title 
    def initiate_team(self, title):
        team_data = {
            "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('educationClass')",
            "displayName": title,
        }
        url = "https://graph.microsoft.com/v1.0/teams"
        self.access_token = self.get_access_token()
        headers = {
            "authorization": f"bearer {self.access_token}",
            "content-type": "application/json"
        }
        response = requests.post(url, headers=headers, json=team_data)
        if(response.status_code != 202):
            print(response.json())
            raise Exception("COULD NOT CREATE TEAM")
        self.team_id = response.headers["content-location"]
        return self

    #add default owners from default_owners.py    
    def add_default_owners(self):
        data = {"values": []}
        for email in default_owners.emails:
            data["values"].append(
                {
                    "@odata.type": "microsoft.graph.aadUserConversationMember",
                    "roles":["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{email}')" 
                }
            )
        url = f"https://graph.microsoft.com/v1.0/teams/{team_spawner.parse_team_id(self.team_id)}/members/add"
        headers = {
            "authorization": f"bearer {self.access_token}",
            "content-type": "application/json"
        }
        response = requests.post(url, headers=headers, json=data)
        if(response.status_code != 200):
            print(response.json())
            raise Exception("COULD NOT ADD MEMBERS")
        return self

    #add a single owner
    def add_owner(self, email):
        data = {
            "values": [
                {
                    "@odata.type": "microsoft.graph.aadUserConversationMember",
                    "roles":["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{email}')" 
                }
            ]
        }
        url = f"https://graph.microsoft.com/v1.0/teams/{team_spawner.parse_team_id(self.team_id)}/members/add"
        headers = {
            "authorization": f"bearer {self.access_token}",
            "content-type": "application/json"
        }
        response = requests.post(url, headers=headers, json=data)
        if(response.status_code != 200):
            print(response.json())
            raise Exception("COULD NOT ADD MEMBERS")
        return self

    #function that controlls how a single team will be created
    def create_team(self, options):
        title = options["-n"]
        default = options["-d"]
        owner = options["-o"]
        if default == True and owner == "":
            self.initiate_team(title).add_default_owners()
        if default == True and owner != "":
            self.initiate_team(title).add_default_owners().add_owner(owner)
        else:
            if owner == "":
                self.initiate_team(title)
            else:
                self.initiate_team(title).add_owner(owner)

    #funtion that controlls how multiple teams will be created
    def create_multiple_teams(self, default):
        with open("export.csv") as csvfile:
            lines = csvfile.readlines()
            for title in lines:
                team_spawner.create_team(title, default)

#list of options
options = {
    "-n": "",
    "-l": "",
    "-o": "",
    "-d": False
}

def invoke_singular_multiple():
    tmsp = team_spawner()
    if options["-l"] != "":
        tmsp.create_multiple_teams(options["-d"])
    else:
        id = tmsp.create_team(options)

def ext_opts(args):
    for i in range(len(args)):
        if len(args) == 1:
            print_usage()
            exit()
        if args[i] in options.keys() and type(options[args[i]]) == str:
            options[args[i]] = args[i+1]
        if args[i] in options.keys() and type(options[args[i]]) == bool:
            options[args[i]] = True

def print_usage():
    print(
        " -n    specify the name of a team \n",
        "-l    name of *.csv file containing the names of teams \n",
        "-o    add a single owner to a team",   
        "-d    add default owners to the team \n \n",
        "type tspawn for help"
    )
            
if __name__ == "__main__":
    ext_opts(sys.argv)
    invoke_singular_multiple()