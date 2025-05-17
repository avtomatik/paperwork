import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent.parent.parent
DATA_DIR = BASE_DIR / 'data'

DST_DIR = os.getenv('DST_DIR')
DST_SPC_DIR = os.getenv('DST_SPC_DIR')

DST_PATH = Path(DST_DIR)
DST_SPC_PATH = Path(DST_SPC_DIR)

# =============================================================================
# Business Specific
# =============================================================================
ACCOUNT = os.getenv('ACCOUNT')
PARTNER_NAME = os.getenv('PARTNER_NAME')
PREFIX = os.getenv('PREFIX')
REF_PLACEHOLDER = os.getenv('REF_PLACEHOLDER')
REF_RESERVED = os.getenv('REF_RESERVED')

# =============================================================================
# Archive Names
# =============================================================================
ARCHIVE_NAME = os.getenv('ARCHIVE_NAME')

# =============================================================================
# Specific Excel File Names
# =============================================================================
FILE_NAME = os.getenv('FILE_NAME')
FILE_NAME_RESERVED = os.getenv('FILE_NAME_RESERVED')

# =============================================================================
# Hook for Multiline `docx-mailmerge` Field
# =============================================================================
MERGE_HOOK = os.getenv('MERGE_HOOK')
