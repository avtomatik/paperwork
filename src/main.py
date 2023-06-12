
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jun 12 21:43:23 2023

@author: green-machine
"""

from pathlib import Path

import pandas as pd

from core.classes import Configuration, Data, Template
from core.functions import (business_logic, get_paths, transform_stringify,
                            write_to_disk)


def main(config: Configuration) -> None:
    kwargs = {
        'io': Path(config.source).joinpath(config.data.file_name)
    }
    df = pd.read_excel(**kwargs).tail(config.num).pipe(business_logic)
    df_formatted = df.copy().pipe(transform_stringify)

    # =========================================================================
    # Main Loop
    # =========================================================================
    for _ in range(df.shape[0]):
        # =====================================================================
        # Populate Fields' Map
        # =====================================================================
        map_fields_extension = {'': ''}
        map_fields = dict(df_formatted.iloc[_, :]) | map_fields_extension
        # =====================================================================
        # Write to Disk
        # =====================================================================
        write_to_disk(config, df, _, map_fields)


if __name__ == '__main__':
    ROWS = 1
    PATHS = get_paths()
# =============================================================================
# TODO: Implement Enum for Configuration
# =============================================================================
    # =========================================================================
    # Purpose: Acts & Scopes
    # =========================================================================
    config = Configuration(
        ROWS,
        PATHS[3],
        PATHS[7],
        Data.COUNTRY,
        Template.SCOPES
    )

    # =========================================================================
    # Purpose: Addendums & Endorsements
    # =========================================================================
    config = Configuration(
        ROWS,
        PATHS[3],
        PATHS[-1],
        Data.CONTRACT,
        Template.ADDENDUM
    )

    # =========================================================================
    # Purpose: Confirmation Letters
    # =========================================================================
    config = Configuration(
        ROWS,
        PATHS[3],
        PATHS[-1],
        Data.LETTER,
        Template.LETTER_CEM
    )

    # =========================================================================
    # Purpose: Debit Notes
    # =========================================================================
    config = Configuration(
        ROWS,
        PATHS[3],
        PATHS[-1],
        Data.DEBIT_NOTE,
        Template.DEBIT_NOTE
    )

    # =========================================================================
    # Purpose: 0x9
    # =========================================================================
    config = Configuration(
        ROWS,
        PATHS[3],
        PATHS[-1],
        Data.LETTER,
        Template.LETTER_0x9
    )

    # =========================================================================
    # Purpose: Slips & Cover Notes
    # =========================================================================
    config = Configuration(
        ROWS,
        PATHS[3],
        PATHS[-1],
        Data.CONTRACT,
        Template.SLIP
    )

    # =========================================================================
    # Purpose: Special Acceptance
    # =========================================================================
    config = Configuration(
        ROWS,
        PATHS[3],
        PATHS[-1],
        Data.CERTIFICATE,
        Template.SPECIAL_ACCEPTANCE
    )

    # =========================================================================
    # Purpose: Treaty Slips & Treaty Endorsements
    # =========================================================================
    config = Configuration(
        ROWS,
        PATHS[3],
        PATHS[-1],
        Data.CONTRACT,
        Template.SLIP_TREATY
    )

    # =========================================================================
    # Purpose: Warranty Letters
    # =========================================================================
    config = Configuration(
        ROWS,
        PATHS[3],
        PATHS[-1],
        Data.LETTER,
        Template.LETTER_WARRANTY
    )

    main(config)
