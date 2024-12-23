# samenvatting-mail-contact

Voordat je dit project kan gebruiken is het nodig om de Gmail API te activeren in je google account. 
Volg hiervoor de stappen zoals hier beschreven: https://developers.google.com/gmail/api/quickstart/python

Pick "Desktop" as app type when creating your OAuth client.

To run the app for the first time you must have a credentials.json file present. This credentials.json contains the OAuth client, which you can download from your project in the Google Workplace under APIS & Services -> Credentials (on https://console.cloud.google.com/). You will have created an OAuth client after having completed the setup steps from the quickstart, see link above. 

To deploy this app and run it via the cloud, you need the following environment variables:

```
GMAIL_ACCESS_TOKEN
GMAIL_REFRESH_TOKEN
GMAIL_CLIENT_ID
GMAIL_CLIENT_SECRET
```

The `GMAIL_CLIENT_ID` and the `GMAIL_CLIENT_SECRET` can be obtained from the Google Workspace. The `GMAIL_ACCESS_TOKEN` and the `GMAIL_REFRESH_TOKEN` can be obtained from the `token.pickle` that is generated after following the authentication flow. To perform this authentication flow, you must run the code locally and have a `credentials.json` file present.

In a notebook or from the cli, run:

```
with open('token.pickle', 'rb') as token:
    gmail_tokens = pickle.load(token)

access_token=gmail_tokens.token
resfresh_token=gmail_tokens.refresh_token
```