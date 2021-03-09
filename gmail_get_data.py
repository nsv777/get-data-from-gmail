import json

from classes.gmail import Gmail


def main():
    with open('config.json', 'r') as f:
        config = json.load(f)
    my_gmail = Gmail(api_scopes=config.get("api_scopes"), credentials_file=config.get("credentials_file"))
    my_gmail.process_messages(body_patterns=config.get("body_patterns"), query=config.get("gmail_query"))


if __name__ == '__main__':
    main()
