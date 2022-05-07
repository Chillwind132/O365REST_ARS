import requests
import xmltodict
import re
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

        #self.get_list_by_title("Test_list")
        #self.get_list_items("Test_list")
        #self.update_list_item("REST_TEST", "1")
        #self.get_ListItemEntityTypeFullName("Test_list")
        #self.create_list_item("REST_TEST","Test_name")
        #self.delete_list_item("Test_list", "6")
        #self.create_list("TEST_REST")
        #self.get_list_guid("REST_TEST")
        #self.update_list("REST_TEST")
        #self.create_field_list("REST_UPDATED")
        #self.create_folder()
        self.download_file('Shared%20Documents/Test_rest', 'dummy.pdf')

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
        
        self.full_url_title = self.site_url + "_api/web/lists/GetByTitle("+ "'" + title + "'" + ")" # https://pwceur.sharepoint.com/sites/GBL-xLoS-SPO-Playground/_api/web/lists/GetByTitle('Test_list')
        response = requests.get(self.full_url_title, headers=self.auth_header, verify=False)

    def get_list_items(self, title):
        items = '/items'
        full_url = self.site_url + "_api/web/lists/GetByTitle("+ "'" + title + "'" + ")" + items
        
        response = requests.get(full_url, headers=self.auth_header, verify=False)
        json = response.json()
        #field = json['d']['results'][0]['Rich_x0020_text']
        
    def create_list_item(self, list, item):

        SP_item = self.get_ListItemEntityTypeFullName(list)
        header = {
        'Authorization': "Bearer " + self.access_token,
        'Accept':'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
        }
        body = {
          "__metadata": {
            "type": SP_item
          },
          "Title": item
        }
        response = requests.post("https://pwceur.sharepoint.com/sites/GBL-xLoS-SPO-Playground/_api/web/lists/GetByTitle(" + "'" + list + "'" + ")/items", json=body, headers=header, verify=False) 
       
    def update_list_item(self, list, item_id):
        SP_item = self.get_ListItemEntityTypeFullName(list)
        title = "Item test title"
        header = {
        'Authorization': "Bearer " + self.access_token,
        'Accept':'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-HTTP-Method': 'MERGE',
        'If-Match': '*'
        }
        body = {
          "__metadata": {
            "type": SP_item
          },
          "Title": title
        }
        response = requests.post("https://pwceur.sharepoint.com/sites/GBL-xLoS-SPO-Playground/_api/web/lists/GetByTitle(" + "'" + list + "'" + ")/items(" + "'" + item_id + "'" + ")", json=body, headers=header, verify=False) 
    
    def delete_list_item(self, list, item_id):
        header = {
        'Authorization': "Bearer " + self.access_token,
        'Accept':'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-HTTP-Method': 'DELETE',
        'If-Match': '*'
        }
        response = requests.post("https://pwceur.sharepoint.com/sites/GBL-xLoS-SPO-Playground/_api/web/lists/GetByTitle(" + "'" + list + "'" + ")/items(" + "'" + item_id + "'" + ")", headers=header, verify=False) 

    def get_ListItemEntityTypeFullName(self, title):
        select = '?$select=ListItemEntityTypeFullName'
        full_url = self.site_url + "_api/web/lists/GetByTitle("+ "'" + title + "'" + ")" + select 
        
        response = requests.get(full_url, headers=self.auth_header, verify=False)
        json = response.json()
        ListItemEntityTypeFullName = json['d']['ListItemEntityTypeFullName']
        return ListItemEntityTypeFullName

    def create_list(self, list_title):
        full_url = self.site_url + "_api/web/lists"
        body = {
          "__metadata": {
            "type": "SP.List"
          },
          "AllowContentTypes": True,
          "BaseTemplate": 100,
         "ContentTypesEnabled": True,
         "Description": "My list description",
         "Title": "REST_TEST2"
        }

        response = requests.post(full_url, headers=self.auth_header, json=body, verify=False)

    def get_list_guid(self, list_title):
        full_url = self.site_url + "_api/web/lists/getByTitle(" + "'" + list_title + "'" + ")/Id"
        response = requests.get(full_url, headers=self.auth_header, verify=False)
        GUID = response.text.split('"Id":"',1)[1].removesuffix('"}}') 
        
        return GUID

    def update_list(self, list_title):
        header = {
        'Authorization': "Bearer " + self.access_token,
        'Accept':'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-HTTP-Method': 'MERGE',
        'If-Match': '*'
        }
        list_guid = self.get_list_guid(list_title)
        full_url = self.site_url + "_api/web/lists(guid" + "'" + list_guid + "'" +")"

        body = {
          "__metadata": {
            "type": "SP.List"
          },
          "Title": "REST_UPDATED"
        }

        response = requests.post(full_url, headers=header, json=body, verify=False) 
    
    def create_field_list(self, list_title):
        list_guid = self.get_list_guid(list_title)
        full_url = self.site_url + "_api/web/lists(guid" + "'" + list_guid + "'" +")/Fields"
        body = {
          "__metadata": {
            "type": "SP.Field"
          },
          "Title": "field title",
          "FieldTypeKind": 2,
          "Required": True,
          "EnforceUniqueValues": False,
          "StaticName": "field name"
        }
        response = requests.post(full_url, headers=self.auth_header, json=body, verify=False) 

    def create_folder(self):
      full_url = self.site_url + "_api/web/folders"
      body = {
        "__metadata": {
          "type": "SP.Folder"
        },
        "ServerRelativeUrl": "Shared%20Documents/Test_rest" # sepcify document list
      }
      response = requests.post(full_url, headers=self.auth_header, json=body, verify=False) 

    def download_file(self, folder_name, file_name):
      full_url = self.site_url + "/_api/web/GetFolderByServerRelativeUrl(" + "'" + folder_name + "'" +")/Files(" + "'" + file_name + "'" +")/$value"
      header = {
      'Authorization': "Bearer " + self.access_token,
              }
      response = requests.get(full_url, headers=header, verify=False)
      open(file_name, 'wb').write(response.content)

if __name__ == "__main__":
    main()
    print("done")
    
