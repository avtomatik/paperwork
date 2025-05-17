
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jun 12 21:43:23 2023

@author: green-machine
"""

from pathlib import Path

import pandas as pd

from core.classes import Data, Template, Work
from core.config import DATA_DIR, DST_DIR, DST_SPC_DIR
from core.funcs import business_logic, transform_stringify, write_to_disk


def main(work: Work) -> None:
    kwargs = {
        'io': Path(work.dir_src).joinpath(work.data_source.file_name)
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
    ROWS = 1
# =============================================================================
# TODO: Implement Enum for Configuration
# =============================================================================
    # =========================================================================
    # Purpose: Acts & Scopes
    # =========================================================================
    config = Work(
        ROWS,
        DATA_DIR,
        DST_SPC_DIR,
        Data.COUNTRY,
        Template.SCOPES
    )

    # =========================================================================
    # Purpose: Addendums & Endorsements
    # =========================================================================
    config = Work(
        ROWS,
        DATA_DIR,
        DST_DIR,
        Data.CONTRACT,
        Template.ADDENDUM
    )

    # =========================================================================
    # Purpose: Confirmation Letters
    # =========================================================================
    config = Work(
        ROWS,
        DATA_DIR,
        DST_DIR,
        Data.LETTER,
        Template.LETTER_CEM
    )

    # =========================================================================
    # Purpose: Debit Notes
    # =========================================================================
    config = Work(
        ROWS,
        DATA_DIR,
        DST_DIR,
        Data.DEBIT_NOTE,
        Template.DEBIT_NOTE
    )

    # =========================================================================
    # Purpose: 0x9
    # =========================================================================
    config = Work(
        ROWS,
        DATA_DIR,
        DST_DIR,
        Data.LETTER,
        Template.LETTER_0x9
    )

    # =========================================================================
    # Purpose: Slips & Cover Notes
    # =========================================================================
    config = Work(
        ROWS,
        DATA_DIR,
        DST_DIR,
        Data.CONTRACT,
        Template.SLIP
    )

    # =========================================================================
    # Purpose: Special Acceptance
    # =========================================================================
    config = Work(
        ROWS,
        DATA_DIR,
        DST_DIR,
        Data.CERTIFICATE,
        Template.SPECIAL_ACCEPTANCE
    )

    # =========================================================================
    # Purpose: Treaty Slips & Treaty Endorsements
    # =========================================================================
    config = Work(
        ROWS,
        DATA_DIR,
        DST_DIR,
        Data.CONTRACT,
        Template.SLIP_TREATY
    )

    # =========================================================================
    # Purpose: Warranty Letters
    # =========================================================================
    config = Work(
        ROWS,
        DATA_DIR,
        DST_DIR,
        Data.LETTER,
        Template.LETTER_WARRANTY
    )

    main(config)
