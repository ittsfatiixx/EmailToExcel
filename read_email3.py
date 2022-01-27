from __future__ import print_function
from distutils.util import execute

import email
import base64
import os.path, os
from re import A

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


import pandas as pd
from openpyxl import load_workbook

from os import remove
import re

unverified_msgs=[]


# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']


def get_message(service, u_id,msg_id):
    try:
        message=service.users().messages().get(userId=u_id, id=msg_id,format='raw').execute()
        msg_raw=base64.urlsafe_b64decode(message['raw'].encode('ASCII'))
        msg_str=email.message_from_bytes(msg_raw)
        content_types=msg_str.get_content_maintype()
        if content_types=='multipart':
            part1,part2=msg_str.get_payload()
            return part1.get_payload()
        else:
            return msg_str.get_payload()
    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f'An error occurred: {error}')


def search_message(service, uid, search_string):
    try:
        search_id=service.users().messages().list(userId=uid,q=search_string).execute()
        number_results=search_id['resultSizeEstimate']
        final_list=[]
        if number_results>0:
            message_ids=search_id['messages']
            for ids in message_ids:
                final_list.append(ids['id'])
            return final_list
        else:
            print('no new messages, so returning empty string')
            return ''            
    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f'An error occurred: {error}')




def get_service():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
        # Call the Gmail API
    service = build('gmail', 'v1', credentials=creds)
    return service

       

#msg_list=[]
#search_str1='subject:new submission on hubspot form "resume registration" newer_than:1d '
#search_str2='subject:razorpay | payment successful for the mentor newer_than:1d'
search_str1='subject:new submission on hubspot form "resume registration" newer_than:1d '
search_str2='subject:razorpay | payment successful for the mentor newer_than:1d'

service= get_service()
msg_id_list1=search_message(service,'me',search_str1) #get ids of all msgs with forms
msg_id_list2=search_message(service,'me',search_str2) #get ids of all msgs with payments

"""
df = pd.DataFrame({'Name': ['E','F','G','H'],
                   'Age': [100,70,40,60]})
"""

#m1=get_message(service,'me',msg_id_list1[0])

#m2=get_message(service,'me',msg_id_list2[0])


data_dict={'Name': [],  'Last name': [], 'Email' : [],'Phone no': [],'Current Role':[],'Product':[],'Required Details ::':[],'Payment Id':[],'Amount':[],'Verified':[]}


#Extracting Form Data {
for msg_id in msg_id_list1:
    print(msg_id)
    m1=get_message(service,'me',msg_id)
    msplit=m1.split('\r\n')
    argList=['First name:','Last name:', 'Email:','Mobile phone number:','Current Role:','Product:','Required Details ::']
    data=[]
    for arg in argList:
        if arg in msplit:
            ind=msplit.index(arg)
            data.append(msplit[ind+1])
        else:
            data.append('NAN')
#        print(data)
    print('Data: \n',data)
    if data[0]!='NAN':
        data_dict['Name'].append(data[0])
    else:
        data_dict['Name'].append(msplit[msplit.index('Last name:')-1])
    data_dict['Last name'].append(data[1])
    data_dict['Email'].append(data[2])
    data_dict['Phone no'].append('+91'+data[3])
    data_dict['Current Role'].append(data[4])
    if data[5]!='NAN':
        data_dict['Product'].append(data[5])
    else:
        data_dict['Product'].append(msplit[msplit.index('CONTACT')-1])
    data_dict['Required Details ::'].append(data[6])

print('\n\nData_dict: \n',data_dict)

    #print(data_dict)

#}


#Payments
TAG_RE = re.compile(r'<[^>]+>|\r|\n|=20')

def remove_tags(text):
    return TAG_RE.sub(',', text)


#a=remove_tags()
#asplit=a.split(',')
payment_id='Payment Id' #index('parment id')+2
amount='Amount' #index('amount')+6    '=E2=82=B9'+2
cust_det= 'Customer Details' #index('Customer Details')+3     +8
phne='&nbsp;'#+2   

PayDict={'Payment Id':[],'Amount':[],'Phone no.':[],'Email':[]}


for msg_id in msg_id_list2:
    print(msg_id)
    m1=get_message(service,'me',msg_id)
    a=remove_tags(m1)
    asplit=a.split(',')
    if 'Payment Id' in asplit:
        PayDict['Payment Id'].append(asplit[asplit.index('Payment Id')+2])
    elif 'Amount' in asplit:
        PayDict['Payment Id'].append(asplit[asplit.index('Amount')-2])
    else:
        unverified_msgs.append(msg_id)
        continue
#        PayDict['Payment Id'].append(msg_id,' Unverified message')
    if 'Amount' in asplit:
        PayDict['Amount'].append(asplit[asplit.index('=E2=82=B9')+2])
    else:
        PayDict['Amount'].append(msg_id,' Unverified message')
        PayDict['Phone no.'].append(msg_id,' Unverified message')
        PayDict['Email'].append(msg_id,' Unverified message')
        continue
    if '&nbsp;' in asplit:
        PayDict['Phone no.'].append(asplit[asplit.index('&nbsp;')+2])
    elif 'nbsp;' in asplit:   
        PayDict['Phone no.'].append(asplit[asplit.index('nbsp;')+2])
    mail1=asplit[asplit.index('Customer Details')+3]
    mail2=asplit[asplit.index('Customer Details')+5]
    mail=mail1[:len(mail1)-1]+asplit[asplit.index('Customer Details')+5]
    PayDict['Email'].append(mail)
#    print(PayDict)
print('\n\nPayDict: \n',PayDict)
#}


Prod_Dict={'Fresher I':49,'Fresher II':49,'Professional I':99,'Professional II':99,'Experienced I':129,'Experienced II':129,'One Page Resume':249,'Cover Letter':99,'Job Links':99,'Linked In Profile Building':129}


"""

>>> data_dict                                            
{
    'Name': ['Suraj', 'Rahul', 'SAURABH SANT GYANESHWAR', 'sagar', 'Aishwarya'], 
    'Last name': ['N', 'Lole', 'KHOBRAGADE', 'kolariya', 'N'], 
    'Email': ['surajganiga14913@gmail.com', 'rahullole@gmail.com', 'srbhkhbgd@gmail.com', 'sagarkolariya11@gmail.com', 'aishwarya09n@gmail.com'], 
    'Phone no': ['7259753256', '9321618237', '8349213133', '8688602182', '9148710255'], 
    'Current Role': ['NAN', 'NAN', 'NAN', 'NAN', 'NAN'], 
    'Product': ['Job Links=20', 'Fresher II', 'Linked In Profile Building=20', 'Linked In Profile Building=20', 'Professional I'], 'Required Details ::': ['JOBS RELATED TO TRAVELING INDUSTRY=20', 'NAN', 'www.linkedin.com/in/saurabh-sant-gyaneshwar-khobragade', '<a href=3D"https://www.linkedin.com/in/sagar-kolariya-51b45a22b">https://ww=', 'Education of B com is pursuing yet .']
}

PayDict
{
    'Payment Id': ['pay_IoGUX3EXGUTXEN', 'pay_InzjV3BCLY2bra', 'pay_InymiQFZMqM9jf', 'pay_InwSUuzJTDpUKH', 'pay_InvpwNYNvg8eHs'], 
    'Amount': ['99', '49', '129', '129', '99'], 
    'Phone no.': ['+917259753256', '+919321618237', '+918349213133', '+918688602182', '+919148710255'], 
    'Email': ['surajganiga14913@gmail.com', 'rahullole@gmail.com', 'srbhkhbgd@gmail.com', 'sagarkolariya11@gmail.com', 'aishwarya09n@gmail.com']}

"""
for mail in data_dict['Email']:
    print(mail)
    verified=''
    indFormMail= data_dict['Email'].index(mail)
    indFormNum=indFormMail
    prod=data_dict['Product'][indFormMail]
    phoneNo=data_dict['Phone no'][indFormNum]
    if mail in PayDict['Email'] :
        indPay=PayDict['Email'].index(mail)
    if phoneNo in PayDict['Phone no.']:
        indPay=PayDict['Phone no.'].index(phoneNo)
    else:
        verified='Unverified'
        data_dict['Payment Id'].append(' Unverified message')
        data_dict['Amount'].append(' Unverified message')
        data_dict['Verified'].append(' Unverified message')
        continue
    if prod in Prod_Dict:
        prod=prod
    elif prod[:len(prod)-3] in Prod_Dict:
        prod=prod[:len(prod)-3]
    else:
        verified='Unverified'
        data_dict['Payment Id'].append(PayDict['Payment Id'][indPay])
        data_dict['Amount'].append(PayDict['Amount'][indPay])
        data_dict['Verified'].append(verified)
        continue
    price=Prod_Dict[prod]
    paidAmt=PayDict['Amount'][indPay]
    if paidAmt==str(price):
        verified='Verified Payment'  #Put into Final dict that payment is verified
    else:
        verified='Unverified'  #Put into Final dict that payment is Unverified
    data_dict['Payment Id'].append(PayDict['Payment Id'][indPay])
    data_dict['Amount'].append(PayDict['Amount'][indPay])
    data_dict['Verified'].append(verified)

print('\n\nData_dict: \n',data_dict)

'''
FinalDict={'Name': [],  'Last name': [], 'Email' : [],'Phone no': [],'Current Role':[],'Product':[],'Required Details ::':[],'Payment Id':[],'Amount':[],'Verified':[]}
'''


df = pd.DataFrame(data_dict)

#excel file
filename='d:/demonew.xlsx'

writer = pd.ExcelWriter(filename, engine='openpyxl',mode='a')

writer.book = load_workbook(filename)
writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
reader = pd.read_excel(filename)

df.to_excel(writer,index=False,header=False,startrow=len(reader)+1)


writer.close()
