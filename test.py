from slacker import Slacker
slack = Slacker('xoxb-1651075977093-1651091890133-ifnzH043FvAlf4br3gQiD13J')
        # Send a message to #general channel
slack.chat.post_message('#test', ' is on signal!')