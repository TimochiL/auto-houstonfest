from openpyxl.utils import get_column_letter


def serialize_yes_or_no(boolean) -> str:
    return 'yes' if boolean else 'no'


# Adapted from https://stackoverflow.com/questions/13197574/openpyxl-adjust-column-width-size
def adjust_column_width(worksheet):
    for column_cells in worksheet.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        worksheet.column_dimensions[get_column_letter(column_cells[0].column)].width = length + 1
