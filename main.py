import json
import pandas as pd
from datetime import datetime, time, timedelta
from pprint import pprint

from openpyxl.workbook import Workbook


def time_sub_time(_time_1: time, _time_2: time):
    seconds = _time_1.second - _time_2.second
    _extra_minute = 0
    _extra_hour = 0
    if seconds < 0:
        seconds = 60 + seconds
        _extra_minute += 1

    minutes = _time_1.minute - _time_2.minute - _extra_minute
    if minutes < 0:
        minutes = 60 + minutes
        _extra_hour += 1

    hours = _time_1.hour - _time_2.hour - _extra_hour

    _time = datetime.strptime(f'{hours}:{minutes}:{seconds}', '%H:%M:%S').time()

    return _time


def time_add_time(_time_1: time, _time_2: time):
    seconds = _time_1.second + _time_2.second
    _extra_minute = 0
    _extra_hour = 0
    if seconds >= 60:
        seconds %= 60
        _extra_minute += 1

    minutes = _time_1.minute + _time_2.minute + _extra_minute
    if minutes >= 60:
        minutes %= 60
        _extra_hour += 1

    hours = _time_1.hour + _time_2.hour + _extra_hour

    _time = datetime.strptime(f'{hours}:{minutes}:{seconds}', '%H:%M:%S').time()

    return _time


def get_excel_content(file_path, _format):
    result = []
    data = dict()
    data_json = dict()

    df = pd.read_excel(file_path, engine="openpyxl")
    _ = df.head().columns.values[0]
    for index, values in enumerate(df.values):
        result.append(values)

        if index > 1:
            name = values[1]
            _date_time = values[3]
            entrance_type = values[5]

            _date, _time = str(_date_time).split(' ')
            _date = datetime.strptime(_date, '%Y-%m-%d')
            _time = datetime.strptime(_time, '%H:%M:%S')

            if not data.get(name, None):
                data[name] = dict()

            if not data[name].get(_date.date(), None):
                data[name][_date.date()] = dict()

            if entrance_type == "10.0.1.148_Chiqish_Entrance Card Reader1":
                entrance_type = "entrance"
            else:
                entrance_type = "exit"

            if not data[name][_date.date()].get(entrance_type, None):
                data[name][_date.date()][entrance_type] = []

            if _time.time() < datetime.strptime('01:00:00', '%H:%M:%S').time():

                _date = _date - timedelta(days=1)
                try:
                    if not data[name][_date.date()].get(entrance_type, None):
                        data[name][_date.date()][entrance_type] = []
                except Exception as e:
                    _date = _date + timedelta(days=1)

            data[name][_date.date()][entrance_type].append(_time.time())

            # for json file
            if not data_json.get(name, None):
                data_json[name] = dict()

            if not data_json[name].get(str(_date.date()), None):
                data_json[name][str(_date.date())] = dict()

            if not data_json[name][str(_date.date())].get(entrance_type, None):
                data_json[name][str(_date.date())][entrance_type] = []

            if _time.time() < datetime.strptime('01:00:00', '%H:%M:%S').time():
                _date = _date - timedelta(days=1)
                try:
                    if not data_json[name][str(_date.date())].get(entrance_type, None):
                        data_json[name][str(_date.date())][entrance_type] = []
                except Exception as e:
                    _date = _date + timedelta(days=1)

            data_json[name][str(_date.date())][entrance_type].append(str(_time.time()))

    with open('analytics_data.json', mode='w', encoding='utf-8') as json_file:
        json.dump(data_json, json_file, indent=4, ensure_ascii=False)

    return data


def process(data: dict):
    _data = dict()
    _data_json = dict()

    for name, value_data in data.items():
        _data[name] = dict()
        _data_json[name] = dict()

        for _user_date, date_data in value_data.items():
            _data[name][_user_date] = dict()
            _data_json[name][str(_user_date)] = dict()

            min_entrance_time = entrance_length = None
            if date_data.get('entrance'):
                min_entrance_time = min(date_data['entrance'])
                entrance_length = len(date_data['entrance'])

            min_exit_time = exit_length = max_exit_time = None
            if date_data.get('exit'):
                min_exit_time = min(date_data['exit'])
                max_exit_time = max(date_data['exit'])
                exit_length = len(date_data['exit'])

            if min_entrance_time and min_exit_time:

                if min_entrance_time < min_exit_time:

                    status = "late"
                    late_time = None
                    if min_entrance_time < datetime.strptime('09:21:00', '%H:%M:%S').time():
                        status = "in time"
                    else:
                        late_time = time_sub_time(min_entrance_time, datetime.strptime('09:21:00', '%H:%M:%S').time())

                    _entrance_list = date_data['entrance']
                    _exit_list = date_data['exit']

                    _data[name][_user_date]["entrance"] = {
                        'time': min_entrance_time,
                        'status': status
                    }

                    _data[name][_user_date]["entrance"]['late_time'] = ""
                    if late_time:
                        _data[name][_user_date]["entrance"]['late_time'] = late_time

                    _data[name][_user_date]['exit'] = {
                        'time': max_exit_time
                    }

                    extra_work_time = None
                    limit_time = datetime.strptime('18:00:00', '%H:%M:%S').time()
                    if max_exit_time > limit_time:
                        extra_work_time = time_sub_time(max_exit_time, limit_time)

                    _data[name][_user_date]['extra_work'] = ""
                    if extra_work_time:
                        _data[name][_user_date]['extra_work'] = extra_work_time

                    _data[name][_user_date]['work_time'] = "Error"
                    if entrance_length == exit_length:
                        work_time = None
                        _is_work_time = True
                        for i in range(entrance_length):
                            try:
                                sub_time = time_sub_time(
                                    date_data['exit'][i],
                                    date_data['entrance'][i]
                                )
                            except Exception as e:
                                _is_work_time = False
                                break

                            if not work_time:
                                work_time = sub_time
                            else:
                                try:
                                    work_time = time_add_time(
                                        work_time,
                                        sub_time
                                    )
                                except Exception as e:
                                    _is_work_time = False
                                    break

                        if _is_work_time:
                            _data[name][_user_date]['work_time'] = work_time

                    # for json file

                    _data_json[name][str(_user_date)]["entrance"] = {
                        'time': str(min_entrance_time),
                        'status': status
                    }

                    _data_json[name][str(_user_date)]["entrance"]['late_time'] = ""
                    if late_time:
                        _data_json[name][str(_user_date)]["entrance"]['late_time'] = str(late_time)

                    _data_json[name][str(_user_date)]['exit'] = {
                        'time': str(max_exit_time)
                    }

                    _data_json[name][str(_user_date)]['extra_work'] = ""
                    if extra_work_time:
                        _data_json[name][str(_user_date)]['extra_work'] = str(extra_work_time)

                    _data_json[name][str(_user_date)]['work_time'] = "Error"
                    if entrance_length == exit_length:
                        work_time = None
                        _is_work_time = True
                        for i in range(entrance_length):
                            try:
                                sub_time = time_sub_time(
                                    date_data['exit'][i],
                                    date_data['entrance'][i]
                                )
                            except Exception as e:
                                _is_work_time = False
                                break

                            if not work_time:
                                work_time = sub_time
                            else:
                                try:
                                    work_time = time_add_time(
                                        work_time,
                                        sub_time
                                    )
                                except Exception as e:
                                    _is_work_time = False
                                    break

                        if _is_work_time:
                            _data_json[name][str(_user_date)]['work_time'] = str(work_time)

                else:
                    _data[name][_user_date] = "Error"
                    _data_json[name][str(_user_date)] = "Error"
            else:
                _data[name][_user_date] = "Error"
                _data_json[name][str(_user_date)] = "Error"

    with open('statistic.json', mode='w', encoding='utf-8') as file:
        json.dump(_data_json, file, indent=4, ensure_ascii=False)

    return _data


def write_to_excel(data: dict):
    writing_data = [
        ['Name', 'Date', 'Entrance Time', 'Status', 'Late Time', 'Exit Time', 'Extra work Time', 'Work Time']
    ]
    _data = dict()
    writing_data2 = [
        ['Name', 'Work days number', 'Late days number']
    ]

    for name, date_data in data.items():
        _data[name] = dict()
        work_days_count = 0
        late_count = 0

        for _date, value in date_data.items():
            work_days_count += 1
            array = []
            array.append(name)
            array.append(_date)

            if isinstance(value, str):
                array.append("")
                array.append("")
                array.append("")
                array.append("")
                array.append("")
                array.append("")
                continue

            if not value.get('entrance', None):
                array.append("")
                array.append("")
                array.append("")
            else:
                array.append(value['entrance']['time'])
                array.append(value['entrance']['status'])
                array.append(value['entrance']['late_time'])

                if value['entrance']['status'] == "late":
                    late_count += 1

            if not value.get('exit', None):
                array.append("")
            else:
                array.append(value['exit']['time'])
            try:
                array.append(value['extra_work'])
            except Exception as e:
                array.append("")

            try:
                array.append(value['work_time'])
            except Exception as e:
                array.append("")
            writing_data.append(array)

        _data[name]['work_days'] = work_days_count
        _data[name]['late_count'] = late_count

    for i in range(2):
        writing_data.append(['', '', '', '', '', '', '', ''])

    writing_data.append(['Name', 'Work Days Number', 'Late Days Number'])
    for key, value in _data.items():
        array = []
        array.append(key)
        array.append(value['work_days'])
        array.append(value['late_count'])
        writing_data.append(array)
        writing_data2.append(array)

    print(writing_data)

    wb = Workbook()
    ws = wb.active
    for row in writing_data:
        ws.append(row)

    wb.save('file_output.xlsx')

    wb = Workbook()
    ws = wb.active
    for row in writing_data2:
        ws.append(row)

    wb.save('file_output_statistics.xlsx')

    print("File has been saved!")


data = get_excel_content('report iyul.xlsx', 'xls')
_data = process(data)

with open('statistic.json', mode='r', encoding='utf-8') as file:
    data = json.load(file)

write_to_excel(data)
