from openpyxl import load_workbook

excel_coa = r"file_path/"
excel_default = r"file_path/"

def get_column_data(xl_file_1, xl_file_2) -> tuple:
    """
    Accesses the specified excel files and cleans and stores the data for the specified column.
    
    Args:
        xl_file_1 -> path to the first excel file.
        xl_file_2 -> path to the second excel file.
    """

    wb_xl_file_1 = load_workbook(xl_file_1)
    wb_xl_file_2 = load_workbook(xl_file_2)
    active_xl_file_1 = wb_xl_file_1["GL Account"]
    active_xl_file_2 = wb_xl_file_2["GL Account"]

    name_column = 2
    xl_file_1_name_column_data = []
    xl_file_2_name_column_data = []

    for rows in active_xl_file_1.iter_rows(min_row=3, values_only=True):
        xl_file_1_name_column_data.append(rows[name_column - 1].strip())

    for rows in active_xl_file_2.iter_rows(min_row=3, values_only=True):
        xl_file_2_name_column_data.append(rows[name_column - 1].strip())

    return xl_file_1_name_column_data, xl_file_2_name_column_data

def compare_lists(all_lists) -> None:
    """
    Compares the values of the return value of the above function for comparison and prints out the values existing
    in one but not the other.

    Args:
        all_lists -> the return type of the get_column_data() function.

    """

    unique_to_list1 = set(all_lists[0]) - set(all_lists[1])
    unique_to_list2 = set(all_lists[1]) - set(all_lists[0])

    print(unique_to_list1)
    print(unique_to_list2)


def main():
    compare_lists(all_lists= get_column_data(xl_file_1= excel_coa, xl_file_2= excel_default))

if __name__ == "__main__":
    main()