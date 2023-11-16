import argparse
import csv
from collections import defaultdict
import yaml
import requests
import pdb
from pprint import pprint

def parse_arguments():
    """
    Parse command line arguments.
    Returns:
        The parsed command line arguments.
    """
    parser = argparse.ArgumentParser(description='Fetch spent time data from Redmine.')
    parser.add_argument('-c', '--config', help='Path to the YAML config file.')
    parser.add_argument('-g', '--group_name', help='Name of the group to fetch data for.')
    parser.add_argument('-s', '--start_date', help='Start date of the interval (YYYY-MM-DD).')
    parser.add_argument('-e', '--end_date', help='End date of the interval (YYYY-MM-DD).')
    parser.add_argument('-o', '--output', help='Path to the output file.')
    return parser.parse_args()

def load_config(path):
    """
    Load Redmine URL and API key from a YAML config file.
    Args:
        path: The path to the YAML config file.
    Returns:
        The Redmine URL and API key.
    """
    with open(path, 'r') as f:
        config = yaml.safe_load(f)
    return config['url'], config['api_key']

def get_group_id(redmine_url, api_key, group_name):
    """
    Get the ID of a group by its name.
    Args:
        redmine_url: The Redmine URL.
        api_key: The API key.
        group_name: The name of the group.
    Returns:
        The ID of the group.
    """
    response = requests.get(f"{redmine_url}/groups.json", params={"key": api_key})
    response.raise_for_status()
    groups = response.json()["groups"]
    for group in groups:
        if group["name"] == group_name:
            return group["id"]
    raise ValueError(f"No group found with name {group_name}")

def fetch_data(redmine_url, api_key, group_id, date_interval):
    """
    Fetch spent time data from Redmine for a specific group.
    Args:
        redmine_url: The Redmine URL.
        api_key: The API key.
        group_id: The ID of the group.
        date_interval: The date interval.
    Returns:
        A dictionary with the spent time data.
    """
    spent_time_data = defaultdict(lambda: defaultdict(float))

    # Fetch group members
    response = requests.get(f"{redmine_url}/groups/{group_id}.json", params={"key": api_key, "include": "users"})
    response.raise_for_status()
    group = response.json()["group"]
    user_ids = [user["id"] for user in group["users"]]

    # Fetch all time entries in the date interval
    params = {"key": api_key, "spent_on": f"><{date_interval['>=']}|{date_interval['<=']}", "limit": 100}
    offset = 0
    while True:
        params["offset"] = offset
        response = requests.get(f"{redmine_url}/time_entries.json", params=params)
        response.raise_for_status()
        entries = response.json()["time_entries"]
        if not entries:
            break
        for entry in entries:
            user_id = entry["user"]["id"]
            if user_id in user_ids:
                response = requests.get(f"{redmine_url}/users/{user_id}.json", params={"key": api_key})
                response.raise_for_status()
                user = response.json()["user"]
                spent_time_data[user_id]["full name"] = f"{user['firstname']} {user['lastname']}"
                spent_time_data[user_id]["email"] = user['mail']
                spent_time_data[user_id]["total spent time"] += entry["hours"]
                spent_time_data[user_id][entry["activity"]["name"]] += entry["hours"]
        offset += len(entries)
        print(f"Fetched {offset} time entries")

    return spent_time_data

def write_tsv(spent_time_data, path='output.tsv'):
    """
    Write spent time data to a TSV file.
    Args:
        spent_time_data: The spent time data.
        path: The path to the TSV file (default 'output.tsv').
    """
    with open(path, "w", encoding='utf8') as f:
        writer = csv.writer(f, delimiter="\t")
        
        # Write header
        header = ["full name", "email", "total spent time", "total spent time - Absence"]
        header.extend(sorted(set(activity for user_data in spent_time_data.values() for activity in user_data if activity not in header)))
        writer.writerow(header)

        # Write data
        for user_data in spent_time_data.values():
            # calculate tot-abs
            user_data['total spent time - Absence'] = user_data.get('total spent time', 0) - user_data.get('Absence (Vacation/VAB/Other)', 0)
            row = [user_data.get(col, 0) for col in header]
            writer.writerow(row)

if __name__ == "__main__":
    args = parse_arguments()
    redmine_url, api_key = load_config(args.config)
    group_id = get_group_id(redmine_url, api_key, args.group_name)
    date_interval = {"<=": args.end_date, ">=": args.start_date}
    spent_time_data = fetch_data(redmine_url, api_key, group_id, date_interval)
    write_tsv(spent_time_data, path=args.output)
