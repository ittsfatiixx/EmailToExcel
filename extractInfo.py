from os import remove
import re
a='New submission on HubSpot Form "Resume Registration"\r\n\r\nNew submission on HubSpot Form "Resume Registration"\r\nPage submitted on: <a href=3D"https://thementor-2.hubspotpagebuilder.com/re=\r\nsume-builder" target=3D"_blank" style=3D"text-decoration: none; color: #00a=\r\n4bd;">Profile Builder</a>First name:\r\nPrahalad\r\nLast name:\r\nBangali\r\nEmail:\r\nprahaladb2000@gmail.com\r\nMobile phone number:\r\n8050077480\r\nUpload a resume :*:\r\n<a href=3D"https://api-na1.hubspot.com/form-integrations/v1/uploaded-files/=\r\nsigned-url-redirect/64574311744?portalId=3D8122058&sign=3D5GfQlmoSrIYhzSN2U=\r\njuvgF3W5w8%3D&conversionId=3D6730e6f5-d510-49ba-96fa-8f0eceb19c00&filename=\r\n=3D6730e6f5-d510-49ba-96fa-8f0eceb19c00-upload_resume-Prahalad_Bangali_Resu=\r\nme.pdf">Prahalad_Bangali_Resume.pdf</a><br>\r\nProduct:\r\nExperienced I;One Page Resume=20\r\nCONTACT\r\nPrahalad Bangali\r\n'

#to extract form data from a hubspot msg
asplit=a.split('\r\n')
argList=['4bd;">Profile Builder</a>First name:','Last name:', 'Email:','Mobile phone number:']
data=[]
for arg in argList:
    if arg in asplit:
        ind=asplit.index(arg)
        data.append(asplit[ind+1])
    print(data)


import re

TAG_RE = re.compile(r'<[^>]+>|\r|\n|=20')

def remove_tags(text):
    return TAG_RE.sub(',', text)

SPACECLN=re.compile(' {2}+')

a=remove_tags()
asplit=a.split(',')
payment_id='Payment Id' #index('parment id')+2
amount='Amount' #index('amount')+6
cust_det= 'Customer Details' #index('Customer Details')+3     +8

sp='  '
for i in asplit:
    if sp in i or i=='':
        asplit.remove(i)


import re
# as per recommendation from @freylis, compile once only
CLEANR = re.compile('<.*?>') 

def cleanhtml(raw_html):
  cleantext = re.sub(CLEANR, '', raw_html)
  return cleantext

