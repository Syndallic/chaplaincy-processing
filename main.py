import pathlib
import re
from collections import defaultdict

import pandas as pd
import xlwings as xw

MAX_ACTIVITY_LETTER = 'R'
TIME_SHEET_PATH = 'data/Time Sheets.csv'
OUTPUT_SHEET_FOLDER_PATH = 'output'

MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
          'November', 'December']


def get_time_sheet_year(time_sheet):
    """Gets the current year, and throws an exception if more than one year is present"""
    year_to_lines_dict = defaultdict(list)

    start_row = 2
    end_row = time_sheet.range('A1').end('down').row
    for i in range(start_row, end_row + 1):
        timestamp = time_sheet.range((i, 1))
        year = timestamp.value[:4]
        year_to_lines_dict[year].append(i)

    if len(year_to_lines_dict) == 1:
        return int(next(iter(year_to_lines_dict)))
    else:
        # convert the smallest lists of line numbers to a printable string
        # (as presumably the years with the fewest instances are wrong)

        # concatenates all lists except the longest one into one list
        list_of_lines = sum(sorted(year_to_lines_dict.values(), key=len)[:-1], [])
        # converts the list of ints to comma-delimited strings
        list_of_lines = ", ".join(map(str, list_of_lines))
        raise ValueError("More than one year included in time sheet. Offending line(s): {}".format(list_of_lines))


def get_time_sheet_data(time_sheet):
    data = [defaultdict(dict) for _ in range(12)]
    start_row = 2
    end_row = time_sheet.range('A1').end('down').row

    for i in range(start_row, end_row + 1):
        month = get_month_index(time_sheet.range((i, 4)).value)
        name = time_sheet.range((i, 2)).value
        activities = time_sheet.range((i, 3)).value
        notes = time_sheet.range((i, 5)).value
        story = time_sheet.range((i, 6)).value

        activities = clean_activities_string(activities)

        if invalid_syntax_in_activities(activities):
            # mark as red if problem
            time_sheet.range((i, 3)).color = (255, 0, 0)
            if 'Activities' not in data[month][name]:
                data[month][name]['Activities'] = defaultdict(int)
        else:
            try:
                add_activities(data[month][name], activities)
            except ValueError as e:
                raise ValueError("Warning: Invalid character '{}' found on row {}, column 3.".format(e.args[0], i))
        add_notes(data[month][name], notes)
        add_story(data[month][name], story)

    return data


def clean_activities_string(activities_string):
    activities_string = activities_string.replace(',', '')
    activities_string = activities_string.replace(' ', '')
    activities_string = activities_string.replace('\n', '')
    activities_string = activities_string.replace('\\', '')
    activities_string = activities_string.replace('/', '')
    activities_string = activities_string.replace(':', '')
    activities_string = activities_string.replace('.', '')
    activities_string = activities_string.upper()
    return activities_string


def invalid_syntax_in_activities(activities_string):
    valid_characters = set(map(chr, range(ord('A'), ord(MAX_ACTIVITY_LETTER) + 1)))
    valid_characters.update(['S', 'P'])
    valid_characters.update(set(map(str, range(10))))
    if not set(activities_string).issubset(valid_characters):
        return True
    # searches for a digit followed by something that isn't a digit or letter
    if re.search("\d[^A-Z0-9]", activities_string) is not None:
        return True

    return False


def add_activities(chaplain_dict, activities):
    if 'Activities' not in chaplain_dict:
        chaplain_dict['Activities'] = defaultdict(int)

    current_multiplier = 0

    i = 0
    while i < len(activities):
        current_char = activities[i]
        if current_char.isdigit():
            # this could be the beginning of a multi-digit number
            numbers_in_string = re.findall(r'\d+', activities[i:])
            current_multiplier = int(numbers_in_string[0])
            i += len(numbers_in_string[0]) - 1
        elif current_char in ['S', 'P']:
            chaplain_dict['Activities'][current_char] += current_multiplier
        elif current_char.isalpha():
            # if invalid character, just default to R
            if ord(current_char) > ord(MAX_ACTIVITY_LETTER):
                current_char = 'R'
            chaplain_dict['Activities'][current_char] += current_multiplier
        else:
            raise ValueError(current_char)

        i += 1


def add_notes(chaplain_dict, notes):
    if notes is None:
        notes = ""

    if 'Notes' not in chaplain_dict:
        chaplain_dict['Notes'] = notes
    else:
        chaplain_dict['Notes'] += "||" + notes


def add_story(chaplain_dict, story):
    if story is None:
        story = ""

    if 'Stories' not in chaplain_dict:
        chaplain_dict['Stories'] = story
    else:
        chaplain_dict['Stories'] += "||" + story


def get_month_index(month_name):
    return MONTHS.index(month_name)


def convert_to_dataframe(month_data):
    dataframe_dict = {}

    activities_list = [letter for letter in range(ord('A'), ord(MAX_ACTIVITY_LETTER) + 1) if
                       letter not in [ord('S'), ord('P')]]
    columns = list(map(chr, activities_list))
    columns += ['S', 'P']
    columns += ['Notes', 'Stories']
    for chaplain in month_data.keys():
        chaplain_data = month_data[chaplain]['Activities']
        chaplain_data['Notes'] = month_data[chaplain]['Notes']
        chaplain_data['Stories'] = month_data[chaplain]['Stories']
        # very sneaky - made sure the columns and keys match up
        # the activities dict is a defaultdict as well, which will default to 0
        dataframe_dict[chaplain] = {column: chaplain_data[column] for column in columns}

    dataframe = pd.DataFrame.from_dict(dataframe_dict, orient='index', columns=columns)
    # sort chaplains alphabetically
    dataframe.sort_index(inplace=True)

    # activities sum per chaplain
    index = ord(MAX_ACTIVITY_LETTER) - ord('A')
    values = dataframe.loc[:, 'A':MAX_ACTIVITY_LETTER].sum(axis=1)
    dataframe.insert(index, 'Activity Total', values)

    # Total sum per column:
    dataframe.loc['Total', :] = dataframe.sum(axis=0)

    return dataframe


def format_table(sheet):
    sheet.autofit()

    # the whole table
    table_range = sheet.range('A1').expand()
    table_range.api.HorizontalAlignment = xw.constants.Constants.xlCenter
    table_range.api.VerticalAlignment = xw.constants.Constants.xlCenter
    table_range.api.Borders.LineStyle = xw.constants.LineStyle.xlContinuous
    table_range.api.Borders.Color = xw.constants.RgbColor.rgbLightGray
    table_range.color = (0, 170, 240)

    column_length = table_range.shape[0]
    row_length = table_range.shape[1]

    # the header
    header = sheet.range((1, 1), (1, row_length))
    header.api.Font.Bold = True

    # the total row
    if sheet.range('A3').value is not None:
        total_row_range = sheet.range((column_length, 1), (column_length, row_length))
    else:
        # if there is no data for this month, the 'total' row will be on row 2
        total_row_range = sheet.range((2, 1), (2, row_length))
    total_row_range.color = (255, 150, 100)

    # notes and stories
    # (assumes these columns are on the end)
    notes_and_stories = sheet.range((1, row_length - 1), (column_length, row_length))
    notes_and_stories.column_width = 80
    notes_and_stories.api.WrapText = True
    notes_and_stories.api.HorizontalAlignment = xw.constants.Constants.xlLeft
    notes_and_stories.api.VerticalAlignment = xw.constants.Constants.xlCenter


def save_output_spreadsheet(book, name):
    pathlib.Path(OUTPUT_SHEET_FOLDER_PATH).mkdir(parents=True, exist_ok=True)
    file_path = OUTPUT_SHEET_FOLDER_PATH + "/{}_summary.xlsx".format(name)

    # if an old version of this spreadsheet is open, connect to and close it so it can be overwritten
    try:
        open_book = xw.Book(file_path)
        open_book.close()
    except FileNotFoundError:
        pass

    book.save(file_path)


def main():
    time_sheet_book = xw.Book(TIME_SHEET_PATH)
    time_sheet = time_sheet_book.sheets[0]
    data = get_time_sheet_data(time_sheet)
    year = get_time_sheet_year(time_sheet)

    output_book = xw.Book()

    dataframes = []

    for month in range(12):
        dataframe = convert_to_dataframe(data[month])
        dataframes.append(dataframe)

        output_sheet = output_book.sheets.add(MONTHS[month], after=output_book.sheets[month])
        output_sheet.range('A1').value = dataframe
        format_table(output_sheet)

    summary_sheet = output_book.sheets[0]
    summary_sheet.name = "Summary"

    summary_dataframe = pd.concat(dataframes)
    summary_dataframe['Notes'] = summary_dataframe['Notes'].replace(0, "")
    summary_dataframe['Stories'] = summary_dataframe['Stories'].replace(0, "")

    # have to extract and combine the string columns manually
    notes = summary_dataframe.groupby(summary_dataframe.index)['Notes'].apply(' '.join).apply(str.strip)
    stories = summary_dataframe.groupby(summary_dataframe.index)['Stories'].apply(' '.join).apply(str.strip)

    summary_dataframe = summary_dataframe.groupby(summary_dataframe.index).sum()

    summary_dataframe = summary_dataframe.assign(Notes=notes)
    summary_dataframe = summary_dataframe.assign(Stories=stories)

    summary_sheet.range('A1').value = summary_dataframe
    format_table(summary_sheet)
    summary_sheet.activate()

    save_output_spreadsheet(output_book, year)

    time_sheet_book.close()


if __name__ == "__main__":
    main()
