from aip import AipOcr

""" 你的 APPID AK SK """
APP_ID = '15635830'
API_KEY = '4MI5BXZRkYqBEB7K9ye3F7wt'
SECRET_KEY = 'kIK1vxy8GdbpbiDran0Id7x185cmWSk9'

client = AipOcr(APP_ID, API_KEY, SECRET_KEY)

""" 读取图片 """
def get_file_content(filePath):
    with open(filePath, 'rb') as fp:
        return fp.read()

image = get_file_content('picture.png')

# print('result',client.basicGeneral(image))

result=client.basicAccurate(image)
# print(dir(result))

# print('result',result)

context=result['words_result']

for text in context:
    print(text['words'])
