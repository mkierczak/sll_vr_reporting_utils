#!/uisr/bin/env python3
# -*- coding: utf-8 -*-
import pdb
from pprint import pprint

import argparse
import requests
import yaml
import re
import xlsxwriter
import sys
import logging



# create logger
logging.basicConfig(
        format='%(asctime)s %(name)-12s %(levelname)-8s %(message)s', 
        level=logging.INFO,
        )
logger = logging.getLogger(__name__)


from collections import defaultdict
def nested_dict():
        return defaultdict(nested_dict)



def uni_shortname2longname(uni):

    # define translation table
    translation = {
        'Chalmers'                          : 'Chalmers University of Technology',
        'KI'                                : 'Karolinska Institutet',
        'KTH'                               : 'KTH Royal Institute of Technology',
        'LiU'                               : 'Linköping University',
        'LU'                                : 'Lund University',
        'SU'                                : 'Stockholm University',
        'SLU'                               : 'Swedish University of Agricultural Sciences',
        'UmU'                               : 'Umeå University',
        'GU'                                : 'University of Gothenburg',
        'UU'                                : 'Uppsala University',
        'NRM'                               : 'Naturhistoriska Riksmuséet',
        'LNU'                               : 'Linnaeus University',
        'Örebro University'                 : 'Örebro University',
        'Other Swedish University'          : 'Other Swedish University',
        'Other Swedish organization'        : 'Other Swedish organization',
        'Healthcare'                        : 'Healthcare',
        'Industry'                          : 'Industry',
        'International University'          : 'International University',
        'Other international organization'  : 'Other international organization',
    }

    if uni not in translation:
        # should we look at PIs email to determin this_
        print(f"WARNING: uni not in translation list, {uni}")

    # return translation if it exists, otherwise return it untranslated
    return translation.get(uni, uni)




def uni_from_pi_email(email):
    """
    Guess the project's university based on the PIs email domain.
    """

    domain = email.split(".")[-2]
    pass




def fetch_time_entries(start_date, end_date, url, api_key):
    """
    Fetches the time entries within the specified date range and project ID.

    Args:
        start_date (str): Start date in format 'YYYY-MM-DD'.
        end_date (str): End date in format 'YYYY-MM-DD'.
        project_id (int): Project ID.
        url (str): Redmine URL.
        api_key (str): Redmine API key.

    Returns:
        set: Set of unique issue IDs.
    """
    params = {
        'key': api_key,
        'spent_on': f'><{start_date}|{end_date}',
#        'project_id': project_id,
        'limit': 100,
        'offset': 0
    }
    issue_ids = nested_dict()

    response = requests.get(f'{url}/time_entries.json', params=params)
    response.raise_for_status()
    data = response.json()

    total_count = data['total_count']

    # Fetch time entries in batches
    while params['offset'] < total_count:
        response = requests.get(f'{url}/time_entries.json', params=params)
        response.raise_for_status()
        data = response.json()

        time_entries = data['time_entries']

        for entry in time_entries:
            try:
                issue_ids[entry['issue']['id']][entry['activity']['name']] += entry['hours']

            except:
                issue_ids[entry['issue']['id']][entry['activity']['name']] = entry['hours']

        params['offset'] += params['limit']

        # Calculate progress percentage
        progress = min(params['offset'], total_count) / total_count * 100
        print(f'Fetching time entries: {progress:.2f}% complete', end='\r')

    print('Fetching time entries: 100% complete')
    
    return issue_ids

def fetch_issue_details(issue_ids, url, api_key, project_filter):
    """
    Fetches the detailed information about each issue.

    Args:
        issue_ids (set): Set of unique issue IDs.
        url (str): Redmine URL.
        api_key (str): Redmine API key.

    Returns:
        list: List of issue details.
    """
    issue_details = []

    print('Fetching issue details:')
    for i, issue_id in enumerate(issue_ids, 1):
        response = requests.get(f'{url}/issues/{issue_id}.json', params={'key': api_key})
        response.raise_for_status()
        data = response.json()

        # Calculate progress percentage
        progress = i / len(issue_ids) * 100
        print(f'Progress: {progress:.2f}%', end='\r')

        # skip issues from other projects than defined below#        if data['issue'] ['project']['name'] not in ['National Bioinformatics Support', ] and not re.search('^Round ', data['issue'] ['project']['name']):
        # filter out everything not in the requested filter list
        pdb.set_trace()
        if data['issue'] ['project']['name'] not in project_filter:
            continue
       
        data['issue']['spent_per_activity'] = dict(issue_ids[issue_id])
        issue_details.append(data['issue'])


    print('Progress: 100%      ')

    return issue_details





def get_custom_field(issue, field_name):


    # get field value
    field_value = [ field['value'] for field in issue['custom_fields'] if field['name']==field_name ]
    if len(field_value) > 0:
        field_value = field_value[0]
    else:
        field_value = ''

    return field_value




def save_issues_as_excel(issue_details, output_path):
    """
    Saves the issues as an Excel file and makes statistics as well.

    Args:
        issue_details (list): List of dictionaries containing issues.
        output_path (str): Path to save the Excel file.
    """

    # create workbook and the info sheet
    workbook  = xlsxwriter.Workbook(output_path)
    info_sheet  = workbook.add_worksheet("Report info")

    # create project list sheet
    pl_sheet  = workbook.add_worksheet("Project list")
    pl_sheet.activate()
    bold_text = workbook.add_format({'bold': True})

    # print metadata
    info_sheet.write("A1", "General info", bold_text)
    info_sheet.write("A2", "Start date:")
    info_sheet.write("B2", args.start_date)
    info_sheet.write("A3", "End date:")
    info_sheet.write("B3", args.start_date)
    info_sheet.write("A4", "Redmine projects:")
    info_sheet.write("B4", )


    # write headers
    headers = ['Project ID', 'PI first name', 'PI last name', 'email', 'Organization', 'SCB Subject Code', 'Sex', 'Tracker', 'LTS project ID', 'Publications', 'Funding', 'Spent time this period', 'Spent time total']
    for col_num, header in enumerate(headers):
        pl_sheet.write(0, col_num, header, bold_text)


    # write data rows
    for row_num, issue in enumerate(issue_details, 1):

        # get PI first and last name
        pi_name       = get_custom_field(issue, 'Principal Investigator')
        pi_name_split = pi_name.split(' ')
        pi_last_name  = pi_name_split[-1]
        pi_first_name = " ".join(pi_name_split[:-1])

        # summarize the hours spent the requested period
        time_spent_this_period = sum([ hours for hours in issue['spent_per_activity'].values() ])

        #pdb.set_trace()

        # print values
        pl_sheet.write(row_num, 0,  issue.get('id',''))
        pl_sheet.write(row_num, 1,  pi_first_name)
        pl_sheet.write(row_num, 2,  pi_last_name)
        pl_sheet.write(row_num, 3,  get_custom_field(issue, 'PI e-mail'))
        pl_sheet.write(row_num, 4,  uni_shortname2longname(get_custom_field(issue, 'Organization')))
        pl_sheet.write(row_num, 5,  get_custom_field(issue, 'SCB Subject Code'))
        pl_sheet.write(row_num, 6,  get_custom_field(issue, 'PI Gender'))
        pl_sheet.write(row_num, 7,  issue.get('tracker',{}).get('name',''))
        pl_sheet.write(row_num, 8,  get_custom_field(issue, 'WABI ID'))
        pl_sheet.write(row_num, 9,  get_custom_field(issue, 'Publication(s)'))
        pl_sheet.write(row_num, 10, get_custom_field(issue, 'Funding'))
        pl_sheet.write(row_num, 11, time_spent_this_period)
        pl_sheet.write(row_num, 12, issue.get('spent_hours',''))







    ### do other statistics

    # PIs per uni
    if 1:
        ppo_sheet = workbook.add_worksheet("Projects per org")

        # print headers
        ppo_sheet.write(f"A1", "Organization", bold_text)
        ppo_sheet.write(f"B1", "#", bold_text)

        # print the UNIQUE function to get all org names
        ppo_sheet.write(f'Y1', "Raw unsorted data for the plot, don't touch.")
        ppo_sheet.write(f"Y2", "=UNIQUE('Project list'!$E$2:'Project list'!$E$10000)") # how to get rid of the 0 0 ?

        # print the counting function
        for row_num in range(2,200):

            # for each row, print the number of occurences in the project list of the corresponding organization name, only if there is a org name on the current row
            ppo_sheet.write(f"Z{row_num}", f"=IF( ISBLANK(Y{row_num}), \"\", COUNTIF('Project list'!$E$2:'Project list'!$E$1000, Y{row_num}))")


        # create a sorted range for the pie chart
        ppo_sheet.write('A2', f"=SORT(Y2:Z10000, 2)") # how to get rid of the spill over range filled with 0?

        # create pie chart
        ppo_chart = workbook.add_chart({'type': 'pie'})

        # add data series
        #[sheetname, first_row, first_col, last_row, last_col].
        ppo_chart.add_series({
            'name'       : '# projects',
            'categories' : ['Projects per org', 1, 0, 1000, 0],
            'values'     : ['Projects per org', 1, 1, 1000, 1],
            "data_labels": {"category": True, 'position': 'outside_end'}
        })

        # tweak the chart
        ppo_chart.set_title({'name': 'Projects per organization'})
        ppo_chart.set_legend({'position': 'none'})
        ppo_chart.set_size({'x_scale': 1.5, 'y_scale': 2})
        ppo_chart.set_style(10)

        # insert the chart
        ppo_sheet.insert_chart('E2', ppo_chart)


        ## chart style gallery, devel
        #for i in range(1,49):

        #    # create pie chart
        #    ppo_chart = workbook.add_chart({'type': 'pie'})

        #    # add data series
        #    #[sheetname, first_row, first_col, last_row, last_col].
        #    ppo_chart.add_series({
        #        'name'      : '# projects',
        #        'categories': ['Projects per org', 1, 0, 1000, 0],
        #        'values'    : ['Projects per org', 1, 1, 1000, 1],
        #    })

        #    # set chart title, duh
        #    ppo_chart.set_title({'name': 'Projects per organization'})
        #    # set chart style
        #    ppo_chart.set_style(i)

        #    # insert the chart
        #    ppo_sheet.write(f"D{2+i*15}", i)
        #    ppo_sheet.insert_chart(f"E{2+i*15}", ppo_chart)




    workbook.close()
    print(f'Statistics saved as {output_path}')





def check_required_args(args):
    """
    Makes sure that we have enough info to generate a report.
    """

    # check if at least one of --sll and --vr is set
    if not (args.sll or args.vr):
        sys.exit("ERROR: At least one of --sll or --vr must be specified.")

    # check if either --long-term, --sm-term or project-[id, name] is set.
    if not (args.long_term or args.sm_term or args.project_id or args.project_name):
        sys.exit("ERROR: No project(s) selected, either --long-term, --sm-term, --project-id or --project-name must be set.")

    # check that some timeframe is set
    if not (args.year or (args.start_date and args.end_date)):
        sys.exit("ERROR: No timeframe set, either --year or --start-date and --end-date must be set.")


def resolve_args(args):
    """
    Resolves the shortcut arguments (e.g. --dm, --long-term) to the actual arguments.
    """

    # resolve --long-term
    if args.long_term:
        try:
            args.project_name += ['Long-term Support']
        except:
            args.project_name  = ['Long-term Support']
        args.recursive = True



    # resolve --sm-term
    if args.sm_term:
        try:
            args.project_name += ['National Bioinformatics Support']
        except:
            args.project_name  = ['National Bioinformatics Support']



    # resolve --dm
    if args.dm:
        try:
            args.activity_filter += ['[DM]']
        except:
            args.activity_filter  = ['[DM]']



    # resolve --year
    if args.year:
        args.start_date = f"{args.year-1}-12-01"
        args.end_date   = f"{args.year  }-11-30"

    return args




def get_redmine_project_structure(config):
    """
    Build a dict with the strucutre of the Redmine projects, and a name-id translation table.
    """

    params = {
        'key': config['api_key'],
        'limit': 100,
        'offset': 0
    }

    # get the project list
    all_projects = []
    
    response = requests.get(f"{config['url']}/projects.json", params=params)
    response.raise_for_status()
    data = response.json()

    all_projects.extend(data['projects'])

    total_count = data['total_count']
    while params['offset'] < total_count:
        response = requests.get(f"{config['url']}/projects.json", params=params)
        response.raise_for_status()
        data = response.json()


        all_projects.extend(data['projects'])

        params['offset'] += params['limit']

        # Calculate progress percentage
        progress = min(params['offset'], total_count) / total_count * 100
        print(f'Fetching Redmine project: {progress:.2f}% complete', end='\r')

    print('Fetching Redmine projects: 100% complete    ')

    # Initialize an empty dictionary to store the hierarchy
    projects_dict = {}

    # restructure projects at a dict
    all_projects = { proj['id']:proj for proj in all_projects }

    # Function to recursively build the dictionary
    def build_project_hierarchy(project, child_id=None):

        if child_id:
            try:
                all_projects[project['id']]['children'].add(child_id)
            except:
                all_projects[project['id']]['children'] = set()
                all_projects[project['id']]['children'].add(child_id)

        # if we have reached the top
        if 'parent' not in project:
            return

        build_project_hierarchy(all_projects[project['parent']['id']], child_id=project['id'])


    for project in all_projects.values():
        build_project_hierarchy(project)

    pdb.set_trace()





def main():


    parser = argparse.ArgumentParser(description='Fetch Redmine time entries between two dates')
    parser.add_argument('--project-id',       help='Redmine Project name to filer out (comma separated if multiple)', type=int, required=False,)
    parser.add_argument('--project-name',     help='Redmine Project name to filer out (comma separated if multiple)', type=str, required=False,)
    parser.add_argument('-s', '--start-date', help='Start date in YYYY-MM-DD format', type=str, required=False,)
    parser.add_argument('-e', '--end-date',   help='End date in YYYY-MM-DD format',   type=str, required=False,)
    parser.add_argument('-c', '--config',     help='Config file path', required=False,)
    parser.add_argument('-o', '--output',     help='Output file path', required=False,)
    parser.add_argument('-y', '--year',       help='Shortcut to select start and end date as $(YEAR-1)-dec to $YEAR-dec'         , type=int)
    parser.add_argument('--sll',              help='Use to include the SciLifeLab report specific statistics in the output file.',      action='store_true')
    parser.add_argument('--vr',               help='Use to include the Vetenskapsrådet report specific statistics in the output file.', action='store_true')
    parser.add_argument('--dm',               help='Use to only consider time logged in an activity with "[DM]" in its name.', action='store_true')
    parser.add_argument('--long-term',        help='Use to only include project in and under the "Long-term Support" project.', action='store_true')
    parser.add_argument('--sm-term',          help='Use to only include project in and under the "National Bioinformatics Support" project.', action='store_true')
    parser.add_argument('-r', '--recursive',  help='Use together with --project-id or --project-name to recursivly include all subprojects to the project specified.', action='store_true')
    parser.add_argument('--activity-filter',  help='Words used to filter out activity types (comma separated if multiple).', type=str, required=False,)


    args = parser.parse_args()

    # check required args
    check_required_args(args)

    # read the config file
    with open(args.config) as f:
        config = yaml.safe_load(f)

    # resolve the arguments
    args = resolve_args(args)

    # construct the project hierarchy
    project_structure = get_redmine_project_structure(config)

    issue_ids       = fetch_time_entries(args, config['url'], config['api_key'])
    issue_details   = fetch_issue_details(issue_ids, config['url'], config['api_key'], args.project_id)
    #statistics      = generate_statistics(issue_details)
    save_issues_as_excel(issue_details, args.output)

if __name__ == '__main__':
    main()
