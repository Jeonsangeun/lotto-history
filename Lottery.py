import subprocess
import json
import requests
import os

def getLottoNumber(draw_no):
    api_url = f"https://www.dhlottery.co.kr/common.do?method=getLottoNumber&drwNo={draw_no}"
    res = requests.get(api_url)
    res.raise_for_status()
    return res.json()

# 기존 기록 불러오기
if os.path.exists("lotto_history.json"):
    with open("lotto_history.json", "r", encoding="utf-8") as f:
        History = json.load(f)
else:
    History = {"Id": [], "Date": [], "Draw": [], "Bonus": []}

# 새 데이터 추가
for i in range(1, 10):
    temp = getLottoNumber(i)
    if temp['drwNo'] not in History["Id"]:
        History["Id"].append(temp['drwNo'])
        History["Date"].append(temp['drwNoDate'])
        History["Draw"].append([temp[f"drwtNo{i}"] for i in range(1, 7)])
        History["Bonus"].append(temp['bnusNo'])

# JSON 저장
with open("lotto_history.json", "w", encoding="utf-8") as f:
    json.dump(History, f, indent=2, ensure_ascii=False)

print("lotto_history.json 저장 완료!")

# Git commit & push 함수
def git_commit_and_push(message="Update lotto history"):
    subprocess.run(["git", "add", "lotto_history.json"], check=True)
    subprocess.run(["git", "commit", "-m", message], check=True)
    subprocess.run(["git", "push"], check=True)

git_commit_and_push()