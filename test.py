import json
import pandas as pd
with open('dataset.json', 'r', encoding='utf-8') as file:  # Replace 'utf-8' with the actual encoding
  # Read and parse the JSON data
  data = json.load(file)

import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
df = pd.DataFrame(columns=['facebookUrl', 'url', 'time', 'user', 'text', 'feedbackId', 'id', 'legacyId', 'attachments', 'likesCount', 'sharesCount', 'commentsCount', 'topCommentText','topCommentDate', 'facebookId'])
df2 = pd.DataFrame(columns=['facebookUrl', 'url', 'time', 'user', 'text', 'feedbackId', 'id', 'legacyId', 'attachments', 'likesCount', 'sharesCount', 'commentsCount', 'topCommentText','topCommentDate', 'facebookId'])
count = 0 
for Data in data:
    if 'topComments' in  Data.keys():
        if 'attachments' in  Data.keys():
                count = count + 1
                if 'photo_image' in Data['attachments'][0].keys():
                    new_data = {'facebookUrl':Data['facebookUrl'], 'url':Data['url'], 'time':Data['time'], 'user':Data['user'], 'text':Data['text'], 'feedbackId':Data['feedbackId'], 'id':Data['id'], 'legacyId':Data['legacyId'], 'attachments':Data['attachments'][0]['photo_image']['uri'], 'likesCount':Data['likesCount'], 'sharesCount':Data['sharesCount'], 'commentsCount':Data['commentsCount'], 'topCommentText':Data['topComments'][0]['text'],'topCommentDate':Data['topComments'][0]['date'], 'facebookId':Data['facebookId']}
                    df = df.append([new_data], ignore_index=True)
                else:
                    new_data = {'facebookUrl':Data['facebookUrl'], 'url':Data['url'], 'time':Data['time'], 'user':Data['user'], 'text':Data['text'], 'feedbackId':Data['feedbackId'], 'id':Data['id'], 'legacyId':Data['legacyId'], 'attachments':'NO', 'likesCount':Data['likesCount'], 'sharesCount':Data['sharesCount'], 'commentsCount':Data['commentsCount'], 'topCommentText':Data['topComments'][0]['text'],'topCommentDate':Data['topComments'][0]['date'], 'facebookId':Data['facebookId']}
                    df2 = df.append([new_data], ignore_index=True)
                print(count)
            
           
writer = pd.ExcelWriter('multi_data.xlsx')
df.to_excel(writer, sheet_name='Sheet1')
df2.to_excel(writer, sheet_name='Sheet2')
writer.save()