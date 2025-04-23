#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jun 12 21:43:23 2023

@author: green-machine
"""

import sys

import pandas as pd

from core.constants import FILE_NAME


def confirm_action(message: str) -> bool:
    print(message)
    return input('Continue? [y/n]: ').strip().lower() == 'y'


def pull_data(f_name: str) -> None:
    message = f'About to Read {f_name} and Save the Data to `{FILE_NAME}`'
    if confirm_action(message):
        print('Processing...')
        pd.read_csv(f_name).to_excel(FILE_NAME, index=False)
        print(f'Data saved to `{FILE_NAME}`')


def push_data(f_name: str) -> None:
    message = f'About to Read `{FILE_NAME}` and Save the Data to {f_name}'
    if confirm_action(message):
        print('Processing...')
        pd.read_excel(FILE_NAME).to_csv(f_name, index=False)
        print(f'Data saved to `{f_name}`')


def main():
    print('Update Manager')
    if len(sys.argv) < 3:
        print('Usage: script.py <filename> <pull|push>')
        return

    file_name = sys.argv[1]
    flag = sys.argv[2].strip().lower()

    if flag == 'pull':
        pull_data(file_name)
    elif flag == 'push':
        push_data(file_name)
    else:
        print('Invalid Argument Provided. Use "pull" or "push".')


if __name__ == '__main__':
    main()
