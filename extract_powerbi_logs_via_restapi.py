from datetime import date, timedelta
from urllib.error import HTTPError
import msal
import requests
import pandas as pd
import sys

## Set all the parameters required by the REST API

daysToBeExtracted = 30

# Set client ID and secret for the service principal
client_id = "<application-id>"
client_secret = "<secret>"
authority_url = "https://login.microsoftonline.com/<domain>"
scope = ["https://analysis.windows.net/powerbi/api/.default"]

# Set CSV path
path = "<path>"

## Use MSAL to grab token

app = msal.ConfidentialClientApplication(client_id, authority=authority_url, client_credential=client_secret)
result = app.acquire_token_for_client(scopes=scope)

# Get latest Power BI Activities
if 'access_token' in result:
    access_token = result['access_token']
    header = {'Content-Type':'application/json', 'Authorization':f'Bearer {access_token}'}
    for i in range(daysToBeExtracted):        

        # Set date to be extracted
        activityDate = date.today() - timedelta(days=i+1)
        activityDateStr = activityDate.strftime("%Y-%m-%d")

        print("Extracting Power BI Activity logs for",activityDateStr)

        # Set the Power BI REST API
        url = "https://api.powerbi.com/v1.0/myorg/admin/activityevents?startDateTime='" + activityDateStr + "T00:00:00'&endDateTime='" + activityDateStr + "T23:59:59'"

        try:
            api_call = requests.get(url=url, headers=header)

            # If unauthorized
            if api_call.status_code == 401:
                print("Error >> 401 Unauthorized - Please check the service principal permissions on Power BI")
                sys.exit()

            # If successful
            elif api_call.status_code == 200:
                #Specify empty Dataframe with all columns
                column_names = ['Id', 'RecordType', 'CreationTime', 'Operation', 'OrganizationId', 'UserType', 'UserKey', 'Workload', 'UserId', 'ClientIP', 'UserAgent', 'Activity', 'IsSuccess', 'RequestId', 'ActivityId', 'ItemName', 'WorkSpaceName', 'DatasetName', 'ReportName', 'WorkspaceId', 'ObjectId', 'DatasetId', 'ReportId', 'ReportType', 'DistributionMethod', 'ConsumptionMethod']
                df = pd.DataFrame(columns=column_names)

                #Set continuation URL
                contUrl = api_call.json()['continuationUri']

                #Get all Activities for first hour, save to dataframe (df1) and append to empty created df
                result = api_call.json()['activityEventEntities']
                df1 = pd.DataFrame(result)
                pd.concat([df, df1])

                #Call Continuation URL as long as results get one back to get all activities through the day
                while contUrl is not None:        
                    api_call_cont = requests.get(url=contUrl, headers=header)
                    contUrl = api_call_cont.json()['continuationUri']
                    result = api_call_cont.json()['activityEventEntities']
                    df2 = pd.DataFrame(result)
                    df = pd.concat([df, df2])
                
                #Set ID as Index of df
                df = df.set_index('Id')

                #Save df as CSV
                df.to_csv(path + activityDateStr + '.csv')
            
            # Others
            else:
                api_call.raise_for_status()
        except HTTPError as e:
            print("Error >> ",e)
            sys.exit()
else:
    print("Error >> Please check service principal id & secrets")