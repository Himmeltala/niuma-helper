import requests
from config.common import COOKIE, API_BASE_URL


def get_mission_info(mission_id):
    url = f'{API_BASE_URL}/{mission_id}'
    headers = {
        'Cookie': COOKIE
    }

    response = requests.get(url, headers=headers)
    return response.json()
