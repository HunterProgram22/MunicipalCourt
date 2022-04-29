# Download the helper library from https://www.twilio.com/docs/python/install
import os
from twilio.rest import Client


# Find your Account SID and Auth Token at twilio.com/console
# and set the environment variables. See http://twil.io/secure
account_sid = 'ACff2e4a0b9d6b051eed464409e27a4e0d'
auth_token = 'b9a079311567fc8044f014faf2c8bd9c'
client = Client(account_sid, auth_token)

# message = client.messages \
#     .create(
#     body=message_body,
#     from_='+18033531251',
#     to='+16167064582'
# )

def create_message(user_name, user_message, user_number):
    return client.messages.create(
        body = get_message_body(user_name, user_message),
        from_='+18033531251',
        to=user_number
    )

def get_message_body(user_name, user_message):
    return f'''Hello {user_name} this is a text reminder from the Delaware Municipal Court.
                {user_message}.'''

test_message_list = [
    ("Justin Kudela", "You are so cool!", "+16167064582"),
    ("Test Justin", "Buckle up!", "+16167064582"),
    ("Amanda Bunner", "We can't wait for ComFest!", "+19377256258"),
    ("Judge Hemmeter", "Hi Judge it is Greg Saylor you should bet on me! ;)", "+16147460883"),
]

for item in test_message_list:
    user_name, user_message, user_number = item
    print(user_name)
    print(user_message)
    print(user_number)
    message = create_message(user_name, user_message, user_number)
    print(message.sid)
    print(message.status)
