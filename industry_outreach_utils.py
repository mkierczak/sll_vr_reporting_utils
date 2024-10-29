import secrets
import requests
from Redmine_utils import Redmine_utils

config = dict()
config['url'] = secrets.redmine_url
config['api_key'] = secrets.api_key

#print(config)

#redmine = Redmine_utils(config)
headers = {'X-Redmine-API-Key': config['api_key']}
params = {"offset" : 0, "limit" : 100, 'cf_6' : 'Industry', 'status_id' : '*'}

while True:
    response = requests.get(f"{config['url']}/issues.json", params=params, headers = headers)
    #print(response.url)
    if response.status_code == 200:
        response.raise_for_status()
        entries = response.json()
        issues = entries['issues']
        for issue in issues:
            print(issue['id'], issue['assigned_to']['name'], issue['status']['name'], issue['tracker']['name'])
            #print(issue)
        if not entries['issues']:
            break
        else:
            params['offset'] = params['offset'] + params['limit']
    
#print(redmine.get_project_structure)
#if __name__ == '__main__':
#    main()
#    