from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


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


def adjust_cell_sizes_for_judge_feedback(worksheet: Worksheet):
    # The dark side of Python is a pathway to many abilities some consider to be unnatural
    column_lengths = [20] + [20] + ([5] * 6 + [20]) * 3 + [10] + [5]
    for index, column_cells in enumerate(worksheet.columns):
        worksheet.column_dimensions[get_column_letter(column_cells[0].column)].width = column_lengths[index] + 1
        column_cells[0].alignment = Alignment(wrap_text=True)  # Wrap headings of tall columns
    worksheet.row_dimensions[1].height = 80
