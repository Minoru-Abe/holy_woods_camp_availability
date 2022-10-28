import requests
LINECODE = "\n"

class SendNotification:
    def send_message(message):
        accesstokenfile = open("line_notify_access_token.csv", "r")
        fileheader = next(accesstokenfile)
        access_token = accesstokenfile.readline().replace(LINECODE,"")

        url = "https://notify-api.line.me/api/notify"
        headers = {"Authorization": "Bearer " + access_token}
        payload = {"message": message}
        r = requests.post(url, headers=headers, params=payload,)
