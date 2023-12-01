# -*- coding: utf-8 -*-
import requests
import yaml

class Redmine_utils:





    def __init__(self, config):

        self.url      = config['url']
        self.api_key  = config['api_key']
        self.projects = self.get_project_structure()




    def get_project_structure(self):
        """
        Build a dict with the strucutre of the Redmine projects, and a name-id translation table.
        """

        params = {
            'key': self.api_key,
            'limit': 100,
            'offset': 0
        }

        # get the project list
        redmine_projects = []
        
        response = requests.get(f"{self.url}/projects.json", params=params)
        response.raise_for_status()
        data = response.json()

        redmine_projects.extend(data['projects'])

        total_count = data['total_count']
        while params['offset'] < total_count:
            response = requests.get(f"{self.url}/projects.json", params=params)
            response.raise_for_status()
            data = response.json()


            redmine_projects.extend(data['projects'])

            params['offset'] += params['limit']

            # Calculate progress percentage
            progress = min(params['offset'], total_count) / total_count * 100
            print(f'Fetching Redmine projects: {min(params["offset"], total_count)} ({progress:.2f}%) complete           ', end='\r')

        print('Fetching Redmine projects: 100% complete                  ')

        # Initialize an empty dictionary to store the hierarchy
        projects_dict = {}

        # restructure projects at a dict
        redmine_projects = { proj['id']:proj for proj in redmine_projects }

        redmine_projects['utils']                    = {}
        redmine_projects['utils']['name2id']         = {}
        redmine_projects['utils']['id2name']         = {}
        redmine_projects['utils']['name2identifier'] = {}
        redmine_projects['utils']['identifier2name'] = {}
        redmine_projects['utils']['id2identifier']   = {}
        redmine_projects['utils']['identifier2id']   = {}
    
    
        # Function to recursively build the dictionary
        def build_project_hierarchy(project, child_ids):
            
            # add name conversions to translation tables
            redmine_projects['utils']['name2id'][project['name']]               = project['id']
            redmine_projects['utils']['id2name'][project['id']]                 = project['name']
            redmine_projects['utils']['name2identifier'][project['name']]       = project['identifier']
            redmine_projects['utils']['identifier2name'][project['identifier']] = project['name']
            redmine_projects['utils']['id2identifier'][project['id']]           = project['identifier']
            redmine_projects['utils']['identifier2id'][project['identifier']]   = project['id']

            # if there are any children to be added
            if len(child_ids) > 0:
                try:
                    redmine_projects[project['id']]['children'].update(child_ids)
                except:
                    redmine_projects[project['id']]['children'] = set()
                    redmine_projects[project['id']]['children'].update(child_ids)

            # if we have reached the top
            if 'parent' not in project:
                return

            # add this project and its children to the list of children
            child_ids.add(project['id'])
            child_ids.update(redmine_projects[project['id']].get('children', set()))

            # pass on the child list to the parent
            build_project_hierarchy(redmine_projects[project['parent']['id']], child_ids=child_ids)


        # process each project
        for key,project in redmine_projects.items():

            # skip utility key
            if key == 'utils':
                continue

            build_project_hierarchy(project, child_ids=set())



        return redmine_projects






