from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
import re


def re_contains_words(hints, label: str):
    regexp = ''.join([f"(?=.*{hint})" for hint in hints])
    return re.search(regexp, label)

def parse_yes_or_no(answer) -> bool:
    if answer is None:
        return False
    return answer.lower() == 'yes'


def serialize_yes_or_no(boolean) -> str:
    return 'yes' if boolean else 'no'


# Adapted from https://stackoverflow.com/questions/13197574/openpyxl-adjust-column-width-size
def adjust_cell_sizes(worksheet):
    for column_cells in worksheet.columns:
        max_length = 0
        for cell in column_cells:
            lines = str(cell.value).split('\n')
            length = max(len(line) for line in lines)
            max_length = max(max_length, length)
        worksheet.column_dimensions[get_column_letter(column_cells[0].column)].width = max_length + 1
