# https://itmining.tistory.com/122
# 파이썬 파일과 디렉토리 경로

import json
import os


# userProfile = os.environ['USERPROFILE']
currentDirectory = os.getcwd()

file_path = currentDirectory + "\\result.json"

data = {}
data['posts'] = []
data['posts'].append({
    "title": "How to get stroage size",
    "url": "https://codechacha.com/ko/get-free-and-total-size-of-volumes-in-android/",
    "draft": "false"
})
data['posts'].append({
    "title": "Android Q, Scoped Storage",
    "url": "https://codechacha.com/ko/android-q-scoped-storage/",
    "draft": "false"
})

print(data)

with open(file_path, 'w') as outfile:
    json.dump(data, outfile, indent=4)
