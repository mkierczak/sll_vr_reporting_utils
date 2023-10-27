# SciLifeLab & VR reporting utils

These are scripts used to gather data that is used in the SciLifeLab and VR annual reports.


## fetch_time_entries_and_make_stats_for_vr_and_ssl_reports.py

Used to fetch all time entries logged between two dates given as arguments. These time entries are used to determin which projects have been active during the selected time period. Data about these active projects are then summarized and printed to a `.xslx` file.

```bash
# example
python3 fetch_time_entries_and_make_stats_for_vr_and_ssl_reports.py -c config.yaml -s 2023-01-01 -e 2023-12-31 -o test.xlsx
```

