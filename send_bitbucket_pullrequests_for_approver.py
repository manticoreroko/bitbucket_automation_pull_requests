from atlassian import Bitbucket
import pandas as pd
import win32com.client
import urllib3
urllib3.disable_warnings()

bitbucket = Bitbucket(
    username='', # login BB
    password='', # password 'HTTP acces tokens' in settings BB
    url='', # https://stash.<...>
    verify_ssl=False
)

release = '01.032.00' # number of actual release

merg = list(bitbucket.get_pull_requests(project='project_name', repository_slug='repository_name', state='MERGED')) # merged pull requests of repository

cnt = 0
for i in range(len(merg)):
    if merg[i]['toRef']['displayId'][-9:] == release:
        cnt += 1  # count merged pull requests in actual release


pull_req = list(bitbucket.get_pull_requests(project='project_name', repository_slug='repository_name', state='OPEN')) # open pull requests, need approve

total = dp.DataFrame(columns=['task'])
for i in range(len(pull_req)):
    tasks = dict({})
    
    tasks['task'] = pull_req[i]['fromRef']['id'][-12:] # task numbers
    
    tasks['link'] = pull_req[i]['links']['self'][0]['href'] # task URL's
    
    for rev in range(len(pull_req[i]['reviewers'])): # approvers statuses
        tasks[pull_req[i]['reviewers'][rev]['user']['displayName']] = pull_req[i]['reviewers'][rev]['status']
        
    total = total.append(tasks, ignore_index=True) # insert approvers statuses in main table
    
try:
    total.insert(2, 'Task', '<a href="' + total['link'] + '">' + total['task'] + '</a>')
except:
    pass   # insert hyperlink

agg_rating = total[total.columns[2:]].unstack()
agg_rating = pd.DataFrame(agg_rating).reset_index()
agg_rating.columns = ['NAME', 'LVL', 'NEED_APPROVE']
del agg_rating['LVL']
agg_rating = agg_rating.query("NEED_APPROVE == 'UNAPROVED' ").groupby(['NAME']).agg({'NEED_APPROVE': 'count'}).reset_index.sort_values('NEED_APPROVED', ascending=False)
# approvers and count NEED_APPROVE status list


ol = win32com.client.Dispatch('Outlook.Application')
newmail = ol.CreateItem(0x0)
newmail.Subject = 'Please, approve it'
newmail.To = 'man@gmail.com; girl@gmail.com'
newmail.HTMLBody = f'''Please, approve this pull requests
<p><b>Merged tasks in current release: {cnt}</b></p>
<p><b>Our repository: </p></b>

<html>
    <head></head>
    <body>
    {agg_rating.to_html(index=False)}
    <body>
</html>

<html>
    <head>
    <meta charset="UTF-8">
    </head>
    <body>
    {total[total.columns[2:]].replace('APPROVED', '<p style="font-size:30px"> ✅ </p>').replace('UNAPPROVED', '').replace('NEED_WORK', '<p style="font-size:30px"> 🕗 </p>').
fillna('').to_html(render_links=True, index=False, escape=False).replace('<tr>', '<tr align="center">')}
    <body>
</html>
'''   # create html mail with table inside

newmail.Send() # send creating mail