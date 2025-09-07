import requests
import json

def getLottoNumber(draw_no):
    api_url = f"https://www.dhlottery.co.kr/common.do?method=getLottoNumber&drwNo={draw_no}"

    try:
        res = requests.get(api_url)
        res.raise_for_status()

        data = res.json()
        return data

    except requests.exceptions.RequestException as e:
        print(f"오류가 발생했습니다: {e}")

Hisory = {
    "Id": [],
    "Date": [],
    "Draw": [],
    "Bonus": []
}
for i in range(1, 10):
    temp = getLottoNumber(i)
    Hisory["Id"].append(temp['drwNo'])
    Hisory["Date"].append(temp['drwNoDate'])
    Hisory["Draw"].append([temp[f"drwtNo{i}"] for i in range(1, 7)])
    Hisory["Bonus"].append(temp['bnusNo'])

print(Hisory)