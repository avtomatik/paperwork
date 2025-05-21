#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jun 12 21:43:23 2023

@author: green-machine
"""

import pandas as pd

from core.classes import Work
from core.config import BASE_DIR
from core.funcs import (business_logic, create_work_from_config, load_config,
                        transform_stringify, write_to_disk)


def main(work: Work) -> None:
    kwargs = {
        'io': work.path_src.joinpath(work.data_source.file_name)
    }
    df = pd.read_excel(**kwargs).tail(work.num).pipe(business_logic)
    df_formatted = df.copy().pipe(transform_stringify)

    # =========================================================================
    # Main Loop
    # =========================================================================
    for index, row in df_formatted.iterrows():
        # =====================================================================
        # Populate Fields' Map
        # =====================================================================
        map_fields_extension = {'': ''}
        map_fields = dict(row) | map_fields_extension

        # =====================================================================
        # Write to Disk
        # =====================================================================
        write_to_disk(work, row, index, map_fields)


if __name__ == '__main__':
    config_path = BASE_DIR / 'config.yml'

    config = load_config(config_path)

    for work_config in config['works']:
        work = create_work_from_config(work_config)
        main(work)
