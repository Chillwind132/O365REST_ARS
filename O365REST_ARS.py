import requests
import secret_p
from urllib.parse import quote
from colorama import Back, Fore, Style, init

class main():
    def __init__(self):

        self.site_url = 'https://pwceur.sharepoint.com/sites/GBL-xLoS-SPO-Playground/' 

        self.ClientId = ""
        self.ClientSecret = ""
        self.tenant = ""
        self.access_token = ""
        self.full_url_title = ""

        self.auth_header = {}

        self.main()

    def main(self): 

        self.get_access_token()
        self.get_list_by_title("Test_list")
        self.get_list_items(self.full_url_title)
        #self.update_list_item(full_url)
        self.get_ListItemEntityTypeFullName(self.full_url_title)

    def get_access_token(self):
        
        init()
        self.ClientId = secret_p.client["ClientId"]
        self.ClientSecret = quote(secret_p.client["ClientSecret"])
        self.tenant = "pwc"
        url = 'https://' + self.tenant + '.sharepoint.com/_vti_bin/client.svc/'

        test_url = 'https://pwceur.sharepoint.com/sites/GBL-xLoS-SPO-Playground/_api/web/lists'
        
        headers = {'Authorization': 'Bearer'}
        response = requests.request("GET", url, headers=headers, verify=False)

        www_auth = response.headers["WWW-Authenticate"]
        Bearer_realm = www_auth.partition('Bearer realm=')[2].partition(',client_id')[0].strip('"')
        client_id = www_auth.partition('client_id=')[2].partition(',trusted_issuers')[0].strip('"')

        url = "https://accounts.accesscontrol.windows.net/" + Bearer_realm + "/tokens/OAuth/2?="

        payload='grant_type=client_credentials&client_id%20%20=' + self.ClientId + '%40' + Bearer_realm + '&client_secret=' + self.ClientSecret +'&resource=' + client_id + '%2Fpwceur.sharepoint.com%40513294a0-3e20-41b2-a970-6d30bf1546fa'
        headers = {'Content-Type': 'application/x-www-form-urlencoded',}
        access_token_json = {}
        response = requests.request("POST", url, headers=headers, data=payload)
        access_token_json = response.json() 
        self.access_token = access_token_json['access_token']

        print(Back.BLUE + f"\nAccess_token: {self.access_token}\n" + Style.RESET_ALL)

        self.auth_header = {
            'Authorization': "Bearer " + self.access_token,
            'Accept':'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
        }

    def get_list_by_title(self, title):
        
        self.full_url_title = self.site_url + "_api/web/lists/GetByTitle("+ "'" + title + "'" + ")"
        response = requests.get(self.full_url_title, headers=self.auth_header, verify=False)

    def get_list_items(self, full_url):
        items = '/items'
        full_url = full_url + items
        
        response = requests.get(full_url, headers=self.auth_header, verify=False)
        
        json = response.json()
        print(response.json())

        field = json['d']['results'][0]['Rich_x0020_text']
        print(field)
        
    
    #def update_list_item(self, full_url):
        
    def get_ListItemEntityTypeFullName(self, full_url):
        select = '?$select=ListItemEntityTypeFullName'
        full_url = full_url + select
        
        response = requests.get(full_url, headers=self.auth_header, verify=False)
        json = response.json()
        ListItemEntityTypeFullName = json['d']['ListItemEntityTypeFullName']
        print(ListItemEntityTypeFullName)
        
        

if __name__ == "__main__":
    main()
