#run command "pip install -r requirements.txt" to install all the python dependencies required for these scripts
import pandas as pd
import json
import base64
import logging
import requests
import csv
import os
from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

#Initialising Logs
logging.basicConfig(format='%(asctime)s - %(message)s', level=logging.INFO, filename='export.log')

personal_access_token = "<PAT TOKEN>"

#Declaring Headers to be sent with the GET request
headers = {}
headers['Content-type'] = "application/json"
headers['Authorization'] = b'Basic ' + base64.b64encode(personal_access_token.encode('utf-8'))

#Reading the WI IDs to be exported
df1= pd.read_csv('export.csv')
temp=[]
def get_comments():
    for i in df1['ID']:
        wi_filename = str(i) + '.json'
        wi_c_filename = str(i) + '_comments.json'
        try:
            #APIs for getting WI Data as well as Comments
            url="https://dev.azure.com/<ORGANIZATION>/_apis/wit/workItems/" + str(i) + "?api-version=5.1&$expand=all"
            url_c="https://dev.azure.com/<ORGANIZATION>/_apis/wit/workItems/" + str(i) +"/comments"

            #Response for WI Data
            response = (requests.get(url, headers=headers, verify=False,
                            auth=(personal_access_token, '<account>'))).text

            #Response for WI Comments
            resp = (requests.get(url_c, headers=headers, verify=False,
                                auth=(personal_access_token, '<account>'))).text
            url_data=json.loads(response)
            url_c_data=json.loads(resp)
            #Creating Directories and Sub-directories for storing WIs and its corresponding Comments json files
            dir= os.makedirs(os.path.join('WI','Comments'),exist_ok=True)
            dir_wi=open(os.path.join('WI',wi_filename), 'w', encoding='utf-8',errors='replace')
            #Writing jsons in the respective directories
            dir_wi.write(json.dumps(url_data, ensure_ascii=False, indent=4,  separators=(",", ": ")))
            dir_comment= open(os.path.join('WI','Comments',wi_c_filename), 'w', encoding='utf-8',errors='replace')
            dir_comment.write(json.dumps(url_c_data, ensure_ascii=False, indent=4,  separators=(",", ": ")))
            logging.info('Extracted ' + wi_filename +' & '+ wi_c_filename)
        except requests.ConnectionError as e:
            print('Error', e.args)

def get_attachements():
    for i in df1['ID']:
        wi_filename = str(i) + '.json'
        wi_c_filename = str(i) + '_comments.json'
    file_n= json.load(open('WI/' + wi_filename, encoding='utf-8'))
    try:
        #Getting the attachement details from the WI jsons
        for y in range(10):
            attach_details={}
            a_id=pd.DataFrame(file_n['relations'][y])
            attach_details['WI_id']=str(i)
            # attach_details['id']=a_id['attributes']['id']
            attach_details['url'] = a_id['url']['id'].rsplit('/', 1)[-1]
            attach_details['name']=a_id['attributes']['name']
            print(attach_details)
            temp.append(attach_details)
    except :
        pass
    csv_columns=['WI_id','url','name']
    with open('attach_details.csv', 'w',newline='') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=csv_columns)
        writer.writeheader()
        for data in temp:
            writer.writerow(data)
    df_attach= pd.read_csv('attach_details.csv')

    #Creating Directory for Attachments
    for x,i in df_attach.iterrows():
        dir_attach = os.makedirs(os.path.join('WI','Attachments',str(i['WI_id'])), exist_ok=True)
        #Sending GET request to Attachment API
        url_a="https://dev.azure.com/<ORGANIZATION>/<PROJECT>/_apis/wit/attachments/"+str(i['url'])+"?fileName="+ i['name']+"&download=true&api-version=5.1"

        resp_a = (requests.get(url_a, headers=headers, verify=False,
                               auth=(personal_access_token, '<account>')))
        try:
            #Saving attachnments in their respective folders based on WI ID
            with open(os.path.join('WI', 'Attachments', str(i['WI_id']),str(i['name'])), 'wb') as filename:
                filename.write(resp_a.content)
                logging.info('Extracted ' + str(i['name'])  + ' for ' + str(i['WI_id']))
        except:
            pass
get_comments()
get_attachements()