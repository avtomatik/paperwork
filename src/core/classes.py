import datetime
from dataclasses import dataclass
from enum import Enum, auto

from core.config import ACCOUNT0


class Data(Enum):

    def _generate_next_value_(name, *args):
        return name

    CERTIFICATE = auto()
    CONTACTS = auto()
    CONTRACT = auto()
    COUNTRY = auto()
    DEBIT_NOTE = auto()
    LETTER = auto()
    SERVICES_ACT = auto()

    @property
    def file_name(self):
        return f'{self.value.lower()}.xlsx'


class Template(Enum):

    def _generate_next_value_(name, *args):
        return name

    ACT = auto()
    ADDENDUM = auto()
    COVER_NOTE = auto()
    DEBIT_NOTE = auto()
    ENDORSEMENT = auto()
    LETTER = auto()
    LETTER_0x9 = auto()
    LETTER_CEM = auto()
    LETTER_FIRM_ORDER = auto()
    LETTER_WARRANTY = auto()
    NDA = auto()
    SCOPES = auto()
    SERVICES_ACT = auto()
    SLIP = auto()
    SLIP_TREATY = auto()
    SPECIAL_ACCEPTANCE = auto()

    @property
    def template_name(self):

        DATE = datetime.date(2022, 8, 12)

        if self.value == 'SLIP_TREATY':
            return f'template_treaty_{ACCOUNT0}{DATE:%Y}_endorsement_{DATE}.docx'

        return f'{self.value.lower()}.docx'


@dataclass
class Work:
    num: int
    source: str
    destination: str
    data: Data
    template: Template
