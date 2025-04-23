#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Apr 23 20:50:18 2025

@author: alexandermikhailov
"""


import datetime
from pathlib import Path
from zipfile import ZipFile

import pandas as pd
from mailmerge import MailMerge
from pandas import DataFrame

from core.classes import Template, Work
from core.config import (ACCOUNT0, ARCHIVE_NAME, BASE_DIR, FILE_NAME_RESERVED,
                         MERGE_HOOK, PARTNER_NAME, PREFIX, REF_PLACEHOLDER,
                         RESERVED_REF)


def business_logic(df: DataFrame) -> DataFrame:
    """Update for More Refined Business Logic"""
    return df


def generate_file_name(template: Template, df: DataFrame, row: int) -> str:
    """
    Generate File Name

    Parameters
    ----------
    template : Template
        DESCRIPTION.
    df : DataFrame
        DESCRIPTION.
    row : int
        DESCRIPTION.

    Returns
    -------
    str
        DESCRIPTION.

    """
    # =========================================================================
    # TODO: Complete This Function
    # =========================================================================
    MAP = {
        Template.ACT: f'Contract {PARTNER_NAME} Act {df.iloc[row, 0]:%Y-%m}.docx',
        Template.ADDENDUM: f'{REF_PLACEHOLDER} Addendum {df.iloc[row, 6]:04n}-{datetime.date.today()}.docx',
        Template.COVER_NOTE: f'{df.iloc[row, 23]} {df.iloc[row, 10]:%Y} Cover Note.docx',
        Template.DEBIT_NOTE: f'{REF_PLACEHOLDER} {"Debit" if df.iloc[row, 16] >= 0 else "Advice"} Note {df.iloc[row, 7]:06n}{"PRM" if df.iloc[row, 16] >= 0 else "RPM"}.docx',
        Template.ENDORSEMENT: f'{REF_PLACEHOLDER} Endorsement {df.iloc[row, 6]:04n}-{row:04n}.docx',
        Template.LETTER: f'{df.iloc[row, 5]} {df.iloc[row, 11]:%Y} Official Letter.docx',
        Template.LETTER_0x9: f'{df.iloc[row, 2]:} {df.iloc[row, 11]:%Y} 0x9.docx',
        Template.LETTER_CEM: f'{df.iloc[row, 5]} {df.iloc[row, 11]:%Y} CEM.docx',
        Template.LETTER_FIRM_ORDER: f'{df.iloc[row, 5]} {df.iloc[row, 11]:%Y} Firm Order Response.docx',
        Template.LETTER_WARRANTY: f'{df.iloc[row, 2]} {df.iloc[row, 1]:%Y} Warranty Letter.docx',
        Template.NDA: 'to_do.docx',
        Template.SCOPES: f'Contract {PARTNER_NAME} Scopes {df.iloc[row, 0]:%Y-%m}.docx',
        Template.SERVICES_ACT: f'Services Act {df.iloc[row, 0].split(";")[1]} {df.iloc[row, 8]:%Y-%m}-{df.iloc[row, 3]:04n}.docx',
        Template.SLIP: f'{df.iloc[row, 23]} {df.iloc[row, 10]:%Y} Slip {df.iloc[row, 21]}.docx',
        Template.SLIP_TREATY: f'{ACCOUNT0} Primary Treaty {RESERVED_REF} {df.iloc[row, 10]:%Y} Endorsement {df.iloc[row, 3]:04n} {df.iloc[row, 21].split(";")[-1]}.docx',
        Template.SPECIAL_ACCEPTANCE: f'{REF_PLACEHOLDER} Special Acceptance {df.iloc[row, 2]:04n}.docx',
    }
    return MAP.get(template) or 'default.docx'


def generate_string_panel(config: Work, two_columned=False) -> list:
    lines = pd.read_excel(Path(config.source).joinpath(FILE_NAME_RESERVED))

    panel = []

    for _, row in lines.iterrows():
        formatted_row = row.copy()
        formatted_row.iloc[-1] = f'{row.iloc[-1]:.7%}'
        if two_columned:
            formatted_row.iloc[-2] = f'-\xa0{row.iloc[-2]};'
        panel.append(formatted_row.iloc[-2:].to_dict())

    return panel


def get_paths(file_name: str = 'paths.txt') -> tuple[Path]:
    file_path = BASE_DIR.joinpath('core').joinpath(file_name)
    with file_path.open(encoding='utf-8') as f:
        return tuple(Path(line.strip() for line in f if line.strip()))


def transform_stringify(df: DataFrame) -> DataFrame:
    datetime_columns = df.select_dtypes(include='datetime64').columns
    float_columns = df.select_dtypes(include='float64').columns
    int_columns = df.select_dtypes(include='int64').columns

    for column in datetime_columns:
        df.loc[:, column] = df.loc[:, column].apply(
            lambda _: f'{_:%d\xa0%B\xa0%Y}'
        )

    for column in float_columns:
        # =====================================================================
        # For Monetary Values
        # =====================================================================
        df.loc[:, column] = df.loc[:, column].apply(lambda _: f'{_:,.2f}')
        # # =====================================================================
        # # For Percentage Values
        # # =====================================================================
        # df_formatted.loc[:, column] = df_formatted.loc[:, column].apply(lambda _: f'{_:.4%}')

    for column in int_columns:
        # =====================================================================
        # For Serial Numbers
        # =====================================================================
        df.loc[:, column] = df.loc[:, column].apply(lambda _: f'{_:04n}')

    for column in ['ref']:
        df.loc[:, column] = df.loc[:, column].apply(lambda _: f'{PREFIX}{_}')

    return df


def write_to_disk(
    config: Work,
    df: DataFrame,
    row: int,
    map_fields: dict[str, str]
) -> None:
    # =========================================================================
    # TODO: Remove Dependence on pandas.DataFrame
    # =========================================================================
    if ARCHIVE_NAME is None:
        template_object = Path(config.source).joinpath(
            config.template.template_name
        )
    else:
        template_object = ZipFile(
            Path(config.source).joinpath(ARCHIVE_NAME)
        ).open(config.template.template_name)
    with MailMerge(template_object) as document:
        document.merge(**map_fields)

        if 'cover_note' in config.template.template_name:
            document.merge_rows(MERGE_HOOK, generate_string_panel())
        if config.template.template_name == Template.LETTER_WARRANTY:
            document.merge_rows(MERGE_HOOK, generate_string_panel())
        if config.template.template_name == Template.LETTER_0x9:
            document.merge_rows(MERGE_HOOK, generate_string_panel(True))

        document.write(
            Path(config.destination).joinpath(
                # =========================================================
                # Generate File Name
                # =========================================================
                generate_file_name(config.template, df, row)
            )
        )
