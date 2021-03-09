# get-data-from-gmail
When you have a lot of emails of the same or similar format in your gmail account, and you want to extract data from them into excel spreadsheet.

### Prerequisites
```shell
pip install -r requirements.txt
```

### Configuration
```
cp config.json.example config.json
```
Modify _credentials_file_, _gmail_query_, and _body_patterns_ entries in config.json

##### credentials_file:
Go to https://console.developers.google.com, create new project, then under "Credentials" create new OAuth client ID; download credentials.json.

##### gmail_query:
A query to filter emails.

##### body_patterns:
List of column names of the to-be created spreadsheet along with regexp to extract required info from email body.

