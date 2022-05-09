import json
import shutil

import requests


def makeJsonData():
    try:
        url = "https://pr.zk.court.gov.ua/new.php"

        payload = "q_court_id=0708"
        headers = {
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Referer": "https://pr.zk.court.gov.ua/sud0708/gromadyanam/csz/",
        }

        response = requests.request("POST", url, headers=headers, data=payload)

        with open("data_pr.json", "w", encoding="utf-8") as file_json:
            json.dump(response.json(), file_json, ensure_ascii=False, indent=8)
    except Exception as ex:
        print("Error")
    else:
        shutil.copy("data_pr.json", "data/data_pr.json")


makeJsonData()
