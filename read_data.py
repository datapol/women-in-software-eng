#!/usr/bin/python

"""Much thanks to this blog post:
http://www.mattcutts.com/blog/write-google-spreadsheet-from-python/
"""

#### import argparse
import datetime
import re

def read_data(data_filename):
    """Get data as a list of dicts from a file"""
    with open(data_filename) as ifstrm:
        return read_list(ifstrm)

def read_list(row_dcts):
    """Get data as a list of dicts from ifstrm or a list"""
    rows_data = list(parse_ss_rows_data_from_ifstrm(row_dcts).values())
    summary_row_data = _summarize_rows_data(rows_data)
    rows_data.append(summary_row_data)
    # Sort in descending order by number of total engineers.
    rows_data.sort(key=lambda r: r['num_eng'], reverse=True)
    return rows_data

def _print_line_skip_warning(line):
    print('Warning... skipping line:\n\t%s\n' % line)


row_key_pattern = '\[(?P<row_key>\w+)\]'
row_key_prog = re.compile(row_key_pattern)


def _extract_row_key_from_data_line(line):
    row_key = None
    m = row_key_prog.match(line)
    if m and m.group('row_key'):
        row_key = m.group('row_key')
    return row_key


col_keys = ('company', 'team', 'num_female_eng', 'num_eng', 'last_updated')


def _extract_col_key_value_from_data_line(line):
    col_key, col_value = None, None

    split_line = line.split(':')
    if len(split_line) == 2:
        if split_line[0] in col_keys:
            col_key, col_value = split_line
            col_key = col_key.strip()
            col_value = col_value.strip()

    return col_key, col_value


required_col_keys = ('company', 'num_female_eng', 'num_eng')


def _clean_row_data(row_data):
    for required_col_key in required_col_keys:
        if required_col_key not in row_data.keys():
            return None

    try:
        num_female_eng = int(row_data['num_female_eng'])
        row_data['num_female_eng'] = num_female_eng
        num_eng = int(row_data['num_eng'])
        row_data['num_eng'] = num_eng
    except ValueError:
        return None

    perc = num_female_eng / num_eng if num_eng != 0 else 0.0
    row_data['percent_female_eng'] = '%.2f' % (100. * perc)

    col_keys = row_data.keys()
    if 'team' not in col_keys:
        row_data['team'] = 'N/A'
    if 'last_updated' not in col_keys:
        row_data['last_updated'] = 'Not provided'

    return row_data

def _parse_ss_rows_data_from_file(filename):
    with open(filename) as ifstrm:
        return parse_ss_rows_data_from_ifstrm(ifstrm)

def parse_ss_rows_data_from_ifstrm(ifstrm):
    rows_data = {}
    row_key = None
    row_data = {}
    for line in ifstrm:
        line = line.strip()
        if not line:
            continue
        elif line.startswith('#'):
            continue
        next_row_key = _extract_row_key_from_data_line(line)

        if row_key and next_row_key:
            # Save last row's data.
            clean_row_data = _clean_row_data(row_data)
            if clean_row_data:
                clean_row_data['key'] = row_key
                rows_data[row_key] = clean_row_data
            row_data = {}
        if next_row_key:
            # Move onto constructing data for the next row.
            row_key = next_row_key
            continue

        if not row_key:
            # No row_key!
            _print_line_skip_warning(line)
        else:
            col_key, col_value = _extract_col_key_value_from_data_line(line)
            if col_key is not None and col_value is not None:
                row_data[col_key] = col_value
            else:
                # Invalid col data.
                _print_line_skip_warning(line)
    clean_row_data = _clean_row_data(row_data)
    if clean_row_data:
        clean_row_data['key'] = row_key
        rows_data[row_key] = clean_row_data
    return rows_data


def _summarize_rows_data(rows_data):
    num_female_eng = 0
    num_eng = 0
    for row_data in rows_data:
        num_female_eng += row_data['num_female_eng']
        num_eng += row_data['num_eng']

    today = datetime.date.today()
    summary_row_data = {
        'key': 'all',
        'company': 'ALL',
        'num_female_eng': num_female_eng,
        'num_eng': num_eng,
        'last_updated': today.strftime('%m/%d/%y')
    }
    return _clean_row_data(summary_row_data)
