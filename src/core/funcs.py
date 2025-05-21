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
import yaml
from mailmerge import MailMerge

from core.classes import Data, Template, Work
from core.config import (ACCOUNT, ARCHIVE_NAME, DATA_DIR, FILE_NAME_RESERVED,
                         MERGE_HOOK, PARTNER_NAME, PREFIX, REF_PLACEHOLDER,
                         REF_RESERVED)


def business_logic(df: pd.DataFrame) -> pd.DataFrame:
    """Update for More Refined Business Logic"""
    return df


def generate_file_name(template: Template, row: pd.Series, index: int) -> str:
    """
    Generate File Name

    Parameters
    ----------
    template : Template
        DESCRIPTION.
    row : pd.Series
        DESCRIPTION.
    index : int
        DESCRIPTION.

    Returns
    -------
    str
        DESCRIPTION.

    """
    # =========================================================================
    # TODO: Complete This Function
    # =========================================================================
    today = datetime.date.today()

    MAP = {
        Template.ACT: f'Contract {PARTNER_NAME} Act {row[0]:%Y-%m}.docx',
        Template.ADDENDUM:
        (
            f'{REF_PLACEHOLDER} Addendum {row[6]:04n}-{today}.docx'
        ),
        Template.COVER_NOTE: f'{row[23]} {row[10]:%Y} Cover Note.docx',
        Template.DEBIT_NOTE:
        (
            f'{REF_PLACEHOLDER} '
            f'{"Debit" if row[16] >= 0 else "Advice"} Note '
            f'{row[7]:06n}'
            f'{"PRM" if row[16] >= 0 else "RPM"}.docx'
        ),
        Template.ENDORSEMENT:
        (
            f'{REF_PLACEHOLDER} Endorsement {row[6]:04n}-'
            f'{index:04n}.docx'
        ),
        Template.LETTER: f'{row[5]} {row[11]:%Y} Official Letter.docx',
        Template.LETTER_0x9: f'{row[2]:} {row[11]:%Y} 0x9.docx',
        Template.LETTER_CEM: f'{row[5]} {row[11]:%Y} CEM.docx',
        Template.LETTER_FIRM_ORDER:
        (
            f'{row[5]} {row[11]:%Y} Firm Order Response.docx'
        ),
        Template.LETTER_WARRANTY: f'{row[2]} {row[1]:%Y} Warranty Letter.docx',
        Template.NDA: 'to_do.docx',
        Template.SCOPES: f'Contract {PARTNER_NAME} Scopes {row[0]:%Y-%m}.docx',
        Template.SERVICES_ACT:
        (
            f'Services Act {row[0].split(";")[1]} '
            f'{row[8]:%Y-%m}-{row[3]:04n}.docx'
        ),
        Template.SLIP: f'{row[23]} {row[10]:%Y} Slip {row[21]}.docx',
        Template.SLIP_TREATY:
        (
            f'{ACCOUNT} Primary Treaty {REF_RESERVED} {row[10]:%Y} '
            f'Endorsement {row[3]:04n} '
            f'{row[21].split(";")[-1]}.docx'
        ),
        Template.SPECIAL_ACCEPTANCE:
        (
            f'{REF_PLACEHOLDER} Special Acceptance {row[2]:04n}.docx'
        ),
    }

    return MAP.get(template, 'default.docx')


def generate_string_panel(two_columned=False) -> list:
    file_path = DATA_DIR / FILE_NAME_RESERVED

    lines = pd.read_excel(file_path)

    panel = []

    for _, row in lines.iterrows():
        formatted_row = row.copy()
        formatted_row.iloc[-1] = f'{row.iloc[-1]:.7%}'
        if two_columned:
            formatted_row.iloc[-2] = f'-\xa0{row.iloc[-2]};'
        panel.append(formatted_row.iloc[-2:].to_dict())

    return panel


def transform_stringify(df: pd.DataFrame) -> pd.DataFrame:
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
        # # ===================================================================
        # # For Percentage Values
        # # ===================================================================
        # df_formatted.loc[:, column] = df_formatted.loc[:, column].apply(
        #     lambda _: f'{_:.4%}'
        # )

    for column in int_columns:
        # =====================================================================
        # For Serial Numbers
        # =====================================================================
        df.loc[:, column] = df.loc[:, column].apply(lambda _: f'{_:04n}')

    for column in ['ref']:
        df.loc[:, column] = df.loc[:, column].apply(lambda _: f'{PREFIX}{_}')

    return df


def write_to_disk(
    work: Work,
    row: pd.Series,
    index: int,
    map_fields: dict[str, str]
) -> None:
    if ARCHIVE_NAME is None:
        template_path = work.path_src.joinpath(work.template.template_name)
    else:
        template_path = ZipFile(
            work.path_src.joinpath(ARCHIVE_NAME)
        ).open(work.template.template_name)
    with MailMerge(template_path) as document:
        document.merge(**map_fields)

        if 'cover_note' in work.template.template_name:
            document.merge_rows(MERGE_HOOK, generate_string_panel())
        if work.template.template_name == Template.LETTER_WARRANTY:
            document.merge_rows(MERGE_HOOK, generate_string_panel())
        if work.template.template_name == Template.LETTER_0x9:
            document.merge_rows(MERGE_HOOK, generate_string_panel(True))

        document.write(
            work.path_dst.joinpath(
                # =============================================================
                # Generate File Name
                # =============================================================
                generate_file_name(work.template, row, index)
            )
        )


def load_config(config_path: Path):
    with config_path.open() as f:
        return yaml.safe_load(f)


def create_work_from_config(work_config):
    # =========================================================================
    # TODO: Check If Making This a Method of Class
    # =========================================================================
    return Work(
        work_config['rows'],
        work_config['data_dir'],
        work_config['dst_path'],
        Data[work_config['data_source']],
        Template[work_config['template']]
    )
