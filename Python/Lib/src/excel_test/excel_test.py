#!/usr/bin/env python3
import sys
import os
import pandas
import openpyxl
import json


HELP = '''
Excel to Json and Json to Excel Converter

<Usage>
excel_test.py mode excel_path json_path

<Args>
mode: --to_json
      --from_json

<Example>
excel_test.py --to_json "C:\\tmp\\in.xlsx" "C:\\tmp\\out.json"
excel_test.py --from_json "C:\\tmp\\out.xlsx" "C:\\tmp\\in.json"
'''


def get_column_info(sheets, conf, row):
    clm_info = {}
    is_key = None

    for clm in conf['columns_info']:
        not_null = clm['not_null']
        value = sheets.cell(row=row, column=clm['clm']).value
        if not_null and value is None:
            return None, None
        clm_info[clm['name']] = value

        if clm['is_key']:
            is_key = str(value)

    return clm_info, is_key


def to_json(excel_path, json_path, conf):
    try:
        if not os.path.exists(excel_path):
            return 2, f'Excel file not exist! [{excel_path}]'

        book = openpyxl.load_workbook(excel_path)

        sheet_name = conf['sheet_name']
        if sheet_name not in book:
            return 2, f'Sheet name is not exist! [{sheet_name}]'

        sheets = book[sheet_name]

        start_row = conf['start_row']
        end_row = conf['end_row']

        clm_infos = {}

        for row in range(start_row, end_row + 1):
            clm_info, is_key = get_column_info(sheets, conf, row)
            if clm_info:
                clm_infos[is_key] = clm_info

        # create .json
        with open(json_path, mode='w') as f:
            write_json = json.dumps(
                clm_infos, sort_keys=False, ensure_ascii=False, indent=4)
            f.write(f'{write_json}\n')

        # create .py
        base_dir = os.path.dirname(json_path)
        new_name = os.path.basename(json_path).replace('.json', '.py')
        py_path = os.path.join(base_dir, new_name)
        with open(py_path, mode='w') as f:
            f.write(f'{sheet_name} = {clm_infos}')

    except Exception as e:
        return 3, f'Exception!! [{e}]'

    return 0, ''


def from_json(json_path, excel_path, conf):
    try:
        if not os.path.exists(json_path):
            return 2, f'Json file not exist! [{json_path}]'

        with open(json_path, mode='r') as f:
            json_data = json.load(f)

        sheet_name = conf['sheet_name']

        df = pandas.DataFrame(json_data)
        df.to_excel(excel_path, sheet_name)

    except Exception as e:
        return 3, f'Exception!! [{e}]'

    return 0, ''


def main(mode, excel_path, json_path, conf):
    if mode == Mode.TO_JSON:
        return to_json(excel_path, json_path, conf[Mode.TO_JSON])
    else:
        return from_json(json_path, excel_path, conf[Mode.FROM_JSON])


class Mode():
    TO_JSON = '--to_json'
    FROM_JSON = '--from_json'


if __name__ == '__main__':
    cd = os.path.dirname(os.path.abspath(__file__))
    args = sys.argv

    # for DEBUG >>
    # args = [
    #     'excel_test.py',
    #     Mode.TO_JSON,
    #     os.path.join(cd, r'sample1\IN.xlsx'),
    #     os.path.join(cd, r'sample1\OUT.json')
    # ]
    # args = [
    #     'excel_test.py',
    #     Mode.FROM_JSON,
    #     os.path.join(cd, r'sample1\OUT.xlsx'),
    #     os.path.join(cd, r'sample1\IN.json')
    # ]
    # for DEBUG <<

    if len(args) <= 1:
        print(HELP)
        sys.exit(0)

    if len(args) < 4:
        print(f'Args is missing.[{len(args)}]')
        sys.exit(1)

    conf = None
    conf_path = os.path.join(cd, 'excel_test.json')
    with open(conf_path, mode='r') as f:
        conf = json.load(f)

    if conf is None:
        print(f'Conf file load is failed.[{conf_path}]')
        sys.exit(1)

    mode = args[1]
    excel_path = args[2]
    json_path = args[3]

    if mode != Mode.TO_JSON and mode != Mode.FROM_JSON:
        print(f'Mode is unknown.[{mode}]')
        sys.exit(1)

    ret, err_msg = main(mode, excel_path, json_path, conf)

    if ret == 0:
        print('SUCCESS!')
    else:
        print(f'FAILED! [{err_msg}]')

    sys.exit(ret)
