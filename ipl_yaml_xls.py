# program to read an IPL yaml file, summarize and save it into a 
# spreadsheet using python
# get the yaml files from https://cricsheet.org/

import os
from yaml import safe_load

import pandas as pd
import datetime

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


# clear the old data from dict
def reset_dict():    
    output_dict = {}
    output_dict = {
        'index': 0,
        'date': '',
        'venue': '',
        'innings': '',
        'target': 0,
        'team': '',
        }
    return output_dict


# load output_dict with ball_by_ball information
def ball_by_ball_info(output_dict, match_info, inning):
    """

    Parameters
    ----------
    output_dict : dict
        Will update the dict with the ball by ball information.
    match_info : nested dict/list
        input from where the ball by ball information will be captured.
    inning : int
        0 or 1, first or second innings respectively.

    Returns
    -------
    output_dict

    """
    target = 0
    output_dict['date'] = match_info['info']['dates']
    output_dict['venue'] = match_info['info']['venue']
    output_dict['innings'] = list(match_info['innings'][inning].keys())[0]
    output_dict['team'] = match_info['innings'][inning][output_dict['innings']]['team']

    # load the ball-by-ball info
    
    if inning: inn_key = '2nd innings'
    else: inn_key = '1st innings'

    num_deliveries = len(match_info['innings'][inning][inn_key]['deliveries'])

    for d in range(num_deliveries):
        # set the key(delivery) and value(total run for that de3livery)
        k = list(match_info['innings'][inning][inn_key]['deliveries'][d].keys())[0]
        v = list(match_info['innings'][inning][inn_key]['deliveries'][d].values())[0]['runs']['total']
        # sum the runs and set the target

        if not inning: target += v
        output_dict.update({k: v})

    
    output_dict['target'] = target
    
    return output_dict


def ipl_df_xls(od_1, od_2, ws, first_match):
    df_1 = pd.DataFrame(od_1)
    df_2 = pd.DataFrame(od_2)
    df = pd.concat([df_1, df_2], ignore_index=True)
        
    # load the worksheet
    if first_match: 
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
    else: 
        for r in dataframe_to_rows(df, index=False, header=False):
            ws.append(r)
    
    return True


#
# main()
#
def main():
    
    OUTPUT_XLS = 'ipl_match_summary.xls'

    # create xls using openpyxl.  Create a workbook object and activate
    wb = Workbook()
    ws = wb.active
    
    # get the folder for the list of yaml to load
    ipl_data_path = os.getcwd() + '/IPLdata/'

    # get the list of ipl yaml to load
    f_list = os.listdir(ipl_data_path)

    # only yamls allowed
    f_list = [yaml_name for yaml_name in f_list if yaml_name[-4:] == 'yaml']

    f_list.sort(key=lambda item: int(item[:-5]))
    
    output_dict = reset_dict()
    
    i = -1
    # for each yaml in the dir extract match information and load 
    # only relevant info at the end create workbook
    for yaml_name in f_list:
        
        # # only yamls allowed
        # if yaml_name[-4:] != 'yaml': continue
    
        i += 1
        
        first_match = True if not i else False
        
        # following line is used for debugging purpose
        # if i == 50: break
        
        print(f'{f_list.index(yaml_name)} Open, extract, and load: {yaml_name}')
    
        # get the raw data
        with open(ipl_data_path+yaml_name, 'r') as y:
            match_info = safe_load(y)
        
        # first innings
        output_dict_1 = ball_by_ball_info(output_dict, match_info, 0)
        output_dict_1['index'] = i
            
        output_dict = reset_dict()
    
        # index
        i += 1
    
        output_dict_2 = ball_by_ball_info(output_dict, match_info, 1)
        output_dict_2['index'] = i
        
        # realign the target
        output_dict_2['target'] = output_dict_1['target']
        output_dict_1['target'] = 0
        
        ipl_df_xls(output_dict_1, output_dict_2, ws, first_match)
    
    
    # save the workbook
    wb.save(OUTPUT_XLS)


if __name__ == '__main__':
    main()
