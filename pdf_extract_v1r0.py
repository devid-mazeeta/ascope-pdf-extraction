# -*- coding: utf-8 -*-

"""
Developer Name : Devid Mazeeta G F
Dev Start Date : 08-May-2020
Dev End Date : 12-May-2020
Dev Revised Date : 12-May-2020
Source Name : PDF Extraction
Python Version : 3.8.2
"""

import re
import json
import xlsxwriter

def regex_match(regex='', match_content=''):
    """
    Match the regex in given content and returns the result as list
    """

    try:
        match = re.findall(regex, match_content, flags=re.I|re.M)

        if match:
            return match
        else:
            return ['']
    except:
        return ['']

def data_clean(value):
    """
    Cleans unwanted text from the given content
    """

    try:
        text = re.sub('<[^>]*?>', ' ', value, flags=re.I|re.M)
        text = re.sub('&amp;', '&', text, flags=re.I|re.M)
        text = re.sub('&nbsp;', ' ', text, flags=re.I|re.M)
        text = re.sub('\s+', ' ', text, flags=re.I|re.M)
        text = text.strip()
        text = re.sub('^\\$\\s*([\\d]+)', r'\1', text, flags=re.I|re.M)
        return text
    except:
        return value

def pdf_extract(file_name=''):
    """
    Extracts data from pdf to html converted files
    """

    header = ['File Name']
    output = [file_name.replace('.html', '.pdf')]
    html_content = open(file_name, 'r', encoding='utf-8').read()
    print('\nExtracting File:', file_name)

    for datapoint_name, datapoint_regex in json_data[file_name].items():
        datapoint_values = []

        if not type(datapoint_regex) is list:
            datapoint_regex = [datapoint_regex]

        for single_regex in datapoint_regex:
            position = 0

            if '<pos>' in single_regex:
                temp = regex_match(regex='<pos>([\d]*?)</pos>(.*)$', match_content=single_regex)[0]
                position = int(temp[0])
                single_regex = temp[1]

            single_value = data_clean(regex_match(regex=single_regex, match_content=html_content)[position])

            if ('Deductible - Each Claim' in datapoint_name) or ('Deductible - Aggregate' in datapoint_name):
                single_value = re.sub('_', '', single_value.strip(), flags=re.I|re.M)

            datapoint_values.append(single_value)

        if 'City | State | Zip' in datapoint_name:
            header.extend(datapoint_name.split(' | '))
            try:
                city_state_zip = re.findall('([a-z\s]+?)\s*\,?\s+([a-z\s]+?)\s+([\d-]{5,10})', datapoint_values[0], flags=re.I|re.M)[0]
            except:
                city_state_zip = ('', '', '')
            output.extend(list(city_state_zip))
        else:
            header.append(datapoint_name)
            output.append(' | '.join(datapoint_values))

    return header, output

if __name__ == "__main__":
    """
    Extraction data from insurance and W9 pdf documents 
    """

    # Creating new excel file and adding a worksheet
    workbook = xlsxwriter.Workbook('pdf_extraction_output.xlsx')
    worksheet = workbook.add_worksheet('Output')
    worksheet.set_column(0, 20, 25)
    header_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#c8d0de'})
    output_format = workbook.add_format({'border': 1})

    # Loads configuration file and extracts file one-by-one
    json_data = json.load(open("pdf_extract.json"))
    index = 0

    for file_name in json_data.keys():
        header, output = pdf_extract(file_name)
        worksheet.write_row(index, 0, header, header_format)
        worksheet.write_row(index+1, 0, output, output_format)
        index += 3

    # Closes excel workboox
    workbook.close()
