#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jun 12 21:43:23 2023

@author: green-machine
"""

import sys

import pandas as pd


def main():

    print('Update Manager')
    file_name = sys.argv[1]
    flag = sys.argv[2].lower()

    if flag == 'pull':
        print(
            f'''About to Read {file_name} and Save the Data to `view.xlsx`'''
        )
        go_no_go = input('Continue? [y/n]: ')
        if go_no_go.lower() == 'y':
            print('Processing')
            kwargs = {
                'excel_writer': 'view.xlsx',
                'index': False
            }
            pd.read_csv(file_name).to_excel(**kwargs)
    elif flag == 'push':
        print(
            f'''About to Read `view.xlsx` and Save the Data to {file_name}'''
        )
        go_no_go = input('Continue? [y/n]: ')
        if go_no_go.lower() == 'y':
            print('Processing')
            kwargs = {
                'filepath_or_buffer': file_name,
                'index': False
            }
            pd.read_excel('view.xlsx').to_csv(**kwargs)
    else:
        print('Invalid Argument Provided')


if __name__ == '__main__':
    main()
