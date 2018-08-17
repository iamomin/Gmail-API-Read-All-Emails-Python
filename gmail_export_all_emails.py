'''
Reading GMAIL using Python
    - Imran Momin
'''

'''
This script does the following:
- Go to Gmal inbox
- Find and read all messages (you can specify labels to read specific emails)
- Extract details (Date, Subject, Body) and export them to a .csv file
'''

'''
Before running this script, the user should get the authentication by following
the link: https://developers.google.com/gmail/api/quickstart/python
Also, client_secret.json should be saved in the same directory as this file
'''
from apiclient import discovery
from apiclient import errors
from httplib2 import Http
from oauth2client import file, client, tools
import base64
from bs4 import BeautifulSoup
# import dateutil.parser as parser
import csv
from time import strftime, gmtime
import sys

def ReadEmailDetails(service, user_id, msg_id):

  temp_dict = { }

  try:

      message = service.users().messages().get(userId=user_id, id=msg_id).execute() # fetch the message using API
      payLoad = message['payload'] # get payload of the message
      headr = payLoad['headers'] # get header of the payload

      for one in headr: # getting the Subject
          if one['name'] == 'Subject':
              msg_subject = one['value']
              temp_dict['Subject'] = msg_subject
          else:
              pass

      for two in headr: # getting the date
          if two['name'] == 'Date':
              msg_date = two['value']
              # date_parse = (parser.parse(msg_date))
              # m_date = (date_parse.datetime())
              temp_dict['DateTime'] = msg_date
          else:
              pass

      # Fetching message body

      part_body = None

      if 'parts' in payLoad:
        email_parts = payLoad['parts'] # fetching the message parts
        part_one  = email_parts[0] # fetching first element of the part
        part_body = part_one['body'] # fetching body of the message
      elif 'body' in payLoad:
        part_body = payLoad['body'] # fetching body of the message

      if part_body['size'] == 0:
        #print(payLoad)
        return None

      part_data = part_body['data'] # fetching data from the body
      clean_one = part_data.replace("-","+") # decoding from Base64 to UTF-8
      clean_one = clean_one.replace("_","/") # decoding from Base64 to UTF-8
      clean_two = base64.b64decode (bytes(clean_one, 'UTF-8')) # decoding from Base64 to UTF-8
      soup = BeautifulSoup(clean_two , "lxml" )
      message_body = soup.body()

      # message_body is a readible form of message body
      # depending on the end user's requirements, it can be further cleaned
      # using regex, beautiful soup, or any other method
      temp_dict['Message_body'] = message_body

  except Exception as e:
      print('Email read error: %s' % e)
      temp_dict = None
      pass

  finally:
      return temp_dict


if __name__ == "__main__":
  print('\n... start')
  sys.stdout.flush()

  # Creating a storage.JSON file with authentication details
  SCOPES = 'https://www.googleapis.com/auth/gmail.modify' # we are using modify and not readonly, as we will be marking the messages Read
  store = file.Storage('storage.json')
  creds = store.get()

  if not creds or creds.invalid:
      flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
      creds = tools.run_flow(flow, store)

  GMAIL = discovery.build('gmail', 'v1', http=creds.authorize(Http()))

  user_id =  'me'
  label_id_one = 'INBOX'
  label_id_two = 'UNREAD'

  # Exporting the values in .tsv
  rows = 0
  file = 'emails_%s.tsv' % (strftime("%Y_%m_%d_%H%M%S", gmtime()))

  print('\n... open file %s' % file)
  sys.stdout.flush()

  with open(file, 'a', encoding='utf-8', newline = '') as csvfile:

      fieldnames = ['Subject','DateTime','Message_body']
      writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter = ',')
      writer.writeheader()

      # label_ids = [label_id_one,label_id_two]  # to read unread emails from inbox
      label_ids = []

      try:
        response = GMAIL.users().messages().list(userId=user_id,
                                                   labelIds=label_ids,
                                                   maxResults=500).execute()

        messages = []
        if 'messages' in response:
          messages.extend(response['messages'])

        while 'nextPageToken' in response:
          page_token = response['nextPageToken']

          response = GMAIL.users().messages().list(userId=user_id,
                                                     labelIds=label_ids,
                                                     pageToken=page_token,
                                                     maxResults=500).execute()

          email_list = response['messages']

          for email in email_list:
            msg_id = email['id'] # get id of individual message
            email_dict = ReadEmailDetails(GMAIL,user_id,msg_id)

            if email_dict is not None:
              writer.writerow(email_dict)
              rows += 1

              print(rows,end="\r")
              sys.stdout.flush()

          print('... fetching next %d emails on [next page token: %s], %d exported so far' % (len(response['messages']), page_token, rows))
          sys.stdout.flush()

      except errors.HttpError as error:
        print('An error occurred: %s' % error)

  print('... emails exported into %s' % (file))
  print("\n... total %d message retrived" % (rows))
  sys.stdout.flush()


  print('... all Done!')

'''
pip3.7 install --upgrade pip
pip3.7 install --upgrade google-api-python-client
pip3.7 install --upgrade httplib2
pip3.7 install --upgrade oauth2client
pip3.7 install --upgrade bs4
pip3.7 install --upgrade lxml
