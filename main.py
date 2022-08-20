import os
import csv
import time

'''
Read CSV
'''


def open_csv(csv_path):
    rows = []
    with open(csv_path, 'r') as csv_file:
        csv_content = csv.reader(csv_file)
        for row in csv_content:
            rows.append(row)
    return rows


''' 
Get start and end time from the given rows
Return a list of Unix timestamp. First is start timme and the second list item is the end time
'''


def get_times(rows):
    # Convert the date time to timestamp
    start_time = time.mktime(time.strptime(f'{rows[1][1:2][0]} {rows[1][2:3][0]}', '%d.%m.%Y %H:%M:%S'))
    end_time = time.mktime(time.strptime(f'{rows[-1][1:2][0]} {rows[-1][2:3][0]}', '%d.%m.%Y %H:%M:%S'))
    return [start_time, end_time]


''' 
Main command
'''


def command():
    dirs = os.listdir()

    data_dirs = []
    cache = {}
    for dir_name in dirs:
        if dir_name[:4] == "DATA":
            data_dirs.append(dir_name)
            # Get csv files only
            files = os.listdir(dir_name)
            files.sort()
            for file in files:
                if file[-3:] == "csv":
                    print(f'{dir_name}/{file}')
                    relative_path = f'{dir_name}/{file}'
                    # open CSV
                    rows = open_csv(f'{dir_name}/{file}')

                    # Get first and last date time and store in cache
                    if rows:
                        times = get_times(rows)
                        cache[relative_path] = [times]
            print(cache)
if __name__ == '__main__':
    command()
