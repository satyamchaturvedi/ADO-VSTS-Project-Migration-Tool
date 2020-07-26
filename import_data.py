from vstsclient.vstsclient import VstsClient
from vstsclient.models import JsonPatchDocument, JsonPatchOperation
from vstsclient.constants import SystemFields, MicrosoftFields, LinkTypes
import csv
import time
import os
import pandas as pd
import json
import logging

#Initialising Logs
logging.basicConfig(format='%(asctime)s - %(message)s', level=logging.INFO,filename='import.log')

#Connecting to ADO using vsts-client module
client = VstsClient('dev.azure.com/<ORGANIZATION>','<PAT TOKEN>')

#Declaring Custom Fields
#get list of all fields from API - https://dev.azure.com/<ORGANIZATION>/<PROJECT>/_apis/wit/fields?api-version=5.1
ACCEPTANCE_CRITERIA = '/fields/Microsoft.VSTS.Common.AcceptanceCriteria'
REF_NO = '/fields/Custom.ReferenceNo' # example


#Reading WI IDs to be imported
df1= pd.read_csv('export.csv')

temp=[]

for i in df1['ID'][:10]:
    wi_filename = str(i) + '.json'
    wi_c_filename = str(i) + '_comments.json'
    file_n = json.load(open(os.path.join('WI', wi_filename), encoding='utf-8'))
    file_c = json.load(open(os.path.join('WI','Comments', wi_c_filename), encoding='utf-8'))
    # Creating a JsonPatchDocument and providing the values for the work item fields
    doc = JsonPatchDocument()
    doc.add(JsonPatchOperation('add', SystemFields.TITLE,file_n['fields']['System.Title']))
    doc.add(JsonPatchOperation('add', SystemFields.AREA_PATH, r'<PROJECT>\<AREA PATH>'))
    doc.add(JsonPatchOperation('add', SystemFields.ITERATION_PATH, r'<PROJECT>\<ITERATION PATH>'))
    doc.add(JsonPatchOperation('add', SystemFields.ASSIGNED_TO, file_n['fields']['Assigned To']))
    doc.add(JsonPatchOperation('add', SystemFields.DESCRIPTION,file_n['fields']['System.Description']))
    #Creating the WIs with above 5 fields
    workitem = client.create_workitem('<PROJECT>', 'WI TYPE', doc)
    logging.info("<PROJECT> WI ID: "+workitem.id + "WI ID: "+ file_n['fields']['System.Id'])
    try:
        doc.add(JsonPatchOperation('add', SystemFields.TAGS, file_n['fields']['System.Tags']))
        # client.update_workitem(workitem.id,doc)
    except:
        pass
    try:
        doc.add(JsonPatchOperation('add',REF_NO,file_n['fields']['System.Id']))
        # client.update_workitem(workitem.id, doc)
    except:
        pass
        doc.add(JsonPatchOperation('replace', SystemFields.STATE, file_n['fields']['System.State']))
        #Updating the Workitems with all the fields
        client.update_workitem(workitem.id,doc)
        logging.info("<PROJECT> WI ID: " + workitem.id + " updated Successfully!")
    try:
        #Uploading and Linking the Attachments to the Newly created WI IDs
        path = '../WI/Attachments/' + str(i) + '/'
        files = [f for f in os.listdir(path)]
        for f in files:
            with open(os.path.join(path,f), 'rb') as x:
                attachment = client.upload_attachment(f, x)
                print(attachment.url)
            client.add_attachment(workitem.id, attachment.url, 'Linking attachment: '+ str(f)+' to work item: '+ str(workitem.id))
            logging.info("Attachment : " +str(f) + " linked to " + workitem.id + " Successfully!")
    except :
        pass
    time.sleep(0.2)
    try:
        #Saving the Comments and <PROJECT> WI IDs to uplaod comments
        for x in range(10):
            c_details = {}
            a_id = file_c['comments'][x]
            c_details['id']= i
            c_details['WI_id']= workitem.id
            c_details['text'] = a_id['text']
            c_details['d_name'] = a_id['revisedBy']['displayName']
            c_details['date'] = a_id['revisedDate']
            c_details['comment'] = " - ".join([c_details['d_name'],c_details['date'],c_details['text']])
            temp.append(c_details)
    except:
        pass
csv_columns=['WI_id','id','text','d_name','date','comment']
with open('comments.csv', 'w',newline='') as csvfile:
    writer = csv.DictWriter(csvfile, fieldnames=csv_columns)
    writer.writeheader()
    for data in temp:
        writer.writerow(data)
csv_file = open('comments.csv', "r")
csv_reader = csv.DictReader(csv_file, delimiter=',')
#Updating Comments
for lines in csv_reader:
    comment = client.create_comment('<PROJECT>', lines['WI_id'], lines['comment'])
    logging.info("Comments added to " + lines['WI_id'] + " Successfully!")

