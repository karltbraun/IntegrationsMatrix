from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from typing import List, Tuple, Dict
import os

""" check_matrix - Compare values in a matrix with their transposed positions.`
"""


FILENAME_IN: str = (
    "/Users/karlbraun/Documents/DEV-L/KTB/Misc/MDR_Vendor_Comparisons.xlsx"
)
# FILENAME_IN: str = "/Users/karlbraun/Documents/DEV-L/KTB/Misc/test_file.xlsx"

TABLE_NAME: str = "B2:AL35"
# TABLE_NAME: str = "A1:F6"


# ################### Excel Table to Matrix ###################


def excel_table_to_matrix(table: Worksheet) -> List[List[str]]:
    matrix = []
    for row in table:
        row_values = [
            cell.value.strip().upper() if cell.value is not None else None
            for cell in row
        ]
        matrix.append(row_values)
    return matrix


# ################### Transform Data ###################


def transform_data(
    data: List[List[str]],
) -> Tuple[Dict[str, Dict[str, str]], Dict[str, Dict[str, str]]]:

    row_dict = {}
    col_dict = {}

    row_headers = data[0][1:]
    col_headers = [row[0] for row in data[1:]]

    for i, row in enumerate(data[1:], start=1):
        row_header = row[0]
        row_dict[row_header] = {
            col_header: row[j] for j, col_header in enumerate(row_headers, start=1)
        }

    for j, col_header in enumerate(row_headers, start=1):
        col_dict[col_header] = {row[0]: row[j] for row in data[1:]}

    return row_dict, col_dict


# ################### Check for None ###################


def check_for_none(value: str) -> str:
    if value is None:
        return "<None>"
    else:
        return value


# ################### Compare Values 2 ###################


def compare_values2(
    row_dict: Dict[str, Dict[str, str]],
    col_dict: Dict[str, Dict[str, str]],
) -> dict:
    """
    Compare the values of each (row, column) combination with the corresponding
    transposed (column, row) combination.
    Return a list of tuples containing the comparison results.
    """

    ROW_COL_KEY = "val_row_col"
    COL_ROW_KEY = "val_col_row"
    MATCH_KEY = "match"

    results = {}
    cell_results = {}
    row_hdrs = list(row_dict.keys())
    col_hdrs = list(col_dict.keys())
    all_hdrs = sorted(set(row_hdrs + col_hdrs))

    print("=============================== Comparing ===============================")
    print(f"Adding to results: {cell_results}")

    print(
        f"{'Vendor 1':<15} {'Vendor 2':<15} {'row,col':<15} {'col,row':<15} {'Match':<5} {'Value':<5}"
    )

    for row_hdr in all_hdrs:

        for col_hdr in all_hdrs:
            # initialize results dictionary

            for col_hdr in all_hdrs:
                cell_results[row_hdr] = {}

                if row_hdr == col_hdr:
                    # skip the diagonal
                    continue

                if (  # ----------------- 0, 0, 0, 0 ----------------
                    (row_hdr not in row_hdrs)
                    and (col_hdr not in col_hdrs)
                    and (col_hdr not in row_hdrs)
                    and (row_hdr not in col_hdrs)
                ):
                    # we shouldn't get here
                    raise ValueError("Invalid header combination")

                #####

                cell_results[row_hdr]["ROW"] = row_hdr
                cell_results[row_hdr]["COL"] = col_hdr
                cell_results[row_hdr]["VALUE"] = ""

                if (  # ----------------- 0, 0, 0, 1 ----------------
                    (row_hdr not in row_hdrs)
                    and (col_hdr not in col_hdrs)
                    and (col_hdr not in row_hdrs)
                    and (row_hdr in col_hdrs)
                ):
                    # row,col value is valid because no valid row header
                    # col,row value is not valid because no valid transposed row header
                    cell_results[row_hdr][ROW_COL_KEY] = f"X {row_hdr}"
                    cell_results[row_hdr][COL_ROW_KEY] = f"X {col_hdr}"
                    cell_results[row_hdr][
                        MATCH_KEY
                    ] = False  # False because we want to manually check

                elif (  # ----------------- 0, 0, 1, 0 ----------------
                    (row_hdr not in row_hdrs)
                    and (col_hdr not in col_hdrs)
                    and (col_hdr in row_hdrs)
                    and (row_hdr not in col_hdrs)
                ):
                    # row,col value is not valid because no valid column header
                    # col,row value is valid because there is a row with the col_hdr value
                    cell_results[row_hdr][ROW_COL_KEY] = f"X {row_hdr}"
                    cell_results[row_hdr][COL_ROW_KEY] = value = check_for_none(
                        col_dict[col_hdr][row_hdr]
                    )
                    cell_results[row_hdr][
                        MATCH_KEY
                    ] = True  # we just use the one good value we have
                    cell_results[row_hdr]["VALUE"] = value

                elif (  # ----------------- 0, 0, 1, 1 ----------------
                    (row_hdr not in row_hdrs)
                    and (col_hdr not in col_hdrs)
                    and (col_hdr in row_hdrs)
                    and (row_hdr in col_hdrs)
                ):
                    # normal row,col value not valid because no valid column header in the normal position
                    # yes, no valid column header either, but we 'fail' on the first item
                    cell_results[row_hdr][ROW_COL_KEY] = f"X {row_hdr}"
                    cell_results[row_hdr][COL_ROW_KEY] = value = check_for_none(
                        col_dict[row_hdr][col_hdr]
                    )
                    cell_results[row_hdr][
                        MATCH_KEY
                    ] = True  # we just use the one good value we have
                    cell_results[row_hdr]["VALUE"] = value

                elif (  # ----------------- 0, 1, 0, 0 ----------------
                    (row_hdr not in row_hdrs)
                    and (col_hdr in col_hdrs)
                    and (col_hdr not in row_hdrs)
                    and (row_hdr not in col_hdrs)
                ):
                    cell_results[row_hdr][ROW_COL_KEY] = f"X {row_hdr}"
                    cell_results[row_hdr][COL_ROW_KEY] = f"X {col_hdr}"
                    cell_results[row_hdr][
                        MATCH_KEY
                    ] = False  # False because we want to manually check

                elif (  # ----------------- 0, 1, 0, 1 ----------------
                    (row_hdr not in row_hdrs)
                    and (col_hdr in col_hdrs)
                    and (col_hdr not in row_hdrs)
                    and (row_hdr in col_hdrs)
                ):
                    cell_results[row_hdr][ROW_COL_KEY] = f"X {row_hdr}"
                    cell_results[row_hdr][COL_ROW_KEY] = f"X {col_hdr}"
                    cell_results[row_hdr][
                        MATCH_KEY
                    ] = False  # False because we want to manually check

                elif (  # ----------------- 0, 1, 1, 0 ----------------
                    (row_hdr not in row_hdrs)
                    and (col_hdr in col_hdrs)
                    and (col_hdr in row_hdrs)
                    and (row_hdr not in col_hdrs)
                ):
                    cell_results[row_hdr][ROW_COL_KEY] = f"X {row_hdr}"
                    cell_results[col_hdr][COL_ROW_KEY] = f"X {col_hdr}"
                    cell_results[row_hdr][
                        MATCH_KEY
                    ] = False  # False because we want to manually check

                elif (  # ----------------- 0, 1, 1, 1 ----------------
                    (row_hdr not in row_hdrs)
                    and (col_hdr in col_hdrs)
                    and (col_hdr in row_hdrs)
                    and (row_hdr in col_hdrs)
                ):
                    cell_results[row_hdr][ROW_COL_KEY] = f"X {row_hdr}"
                    cell_results[row_hdr][COL_ROW_KEY] = value = check_for_none(
                        col_dict[row_hdr][col_hdr]
                    )
                    cell_results[row_hdr][
                        MATCH_KEY
                    ] = True  # we just use the one good value we have
                    cell_results[row_hdr]["VALUE"] = value

                elif (  # ----------------- 1, 0, 0, 0 ----------------
                    (row_hdr in row_hdrs)
                    and (col_hdr not in col_hdrs)
                    and (col_hdr not in row_hdrs)
                    and (row_hdr not in col_hdrs)
                ):
                    cell_results[row_hdr][ROW_COL_KEY] = f"X {col_hdr}"
                    cell_results[row_hdr][COL_ROW_KEY] = f"X {col_hdr}"
                    cell_results[row_hdr][
                        MATCH_KEY
                    ] = False  # False because we want to manually check

                elif (  # ----------------- 1, 0, 0, 1 ----------------
                    (row_hdr in row_hdrs)
                    and (col_hdr not in col_hdrs)
                    and (col_hdr not in row_hdrs)
                    and (row_hdr in col_hdrs)
                ):
                    cell_results[row_hdr][ROW_COL_KEY] = f"X {col_hdr}"
                    cell_results[row_hdr][COL_ROW_KEY] = f"X {col_hdr}"
                    cell_results[row_hdr][
                        MATCH_KEY
                    ] = False  # False because we want to manually check

                elif (  # ----------------- 1, 0, 1, 0 ----------------
                    (row_hdr in row_hdrs)
                    and (col_hdr not in col_hdrs)
                    and (col_hdr in row_hdrs)
                    and (row_hdr not in col_hdrs)
                ):
                    cell_results[row_hdr][ROW_COL_KEY] = f"X {col_hdr}"
                    cell_results[row_hdr][COL_ROW_KEY] = f"X {row_hdr}"
                    cell_results[row_hdr][
                        MATCH_KEY
                    ] = False  # False because we want to manually check

                elif (  # ----------------- 1, 0, 1, 1 ---------------
                    (row_hdr in row_hdrs)
                    and (col_hdr not in col_hdrs)
                    and (col_hdr in row_hdrs)
                    and (row_hdr in col_hdrs)
                ):
                    cell_results[row_hdr][ROW_COL_KEY] = f"X {col_hdr}"
                    cell_results[row_hdr][COL_ROW_KEY] = value = check_for_none(
                        col_dict[row_hdr][col_hdr]
                    )
                    cell_results[row_hdr][
                        MATCH_KEY
                    ] = True  # we just use the one good value we have
                    cell_results[row_hdr]["VALUE"] = value

                elif (  # ----------------- 1, 1, 0, 0 ----------------
                    (row_hdr in row_hdrs)
                    and (col_hdr in col_hdrs)
                    and (col_hdr not in row_hdrs)
                    and (row_hdr not in col_hdrs)
                ):
                    cell_results[row_hdr][ROW_COL_KEY] = value = check_for_none(
                        row_dict[row_hdr][col_hdr]
                    )
                    cell_results[row_hdr][COL_ROW_KEY] = f"X {col_hdr}"
                    cell_results[row_hdr][
                        MATCH_KEY
                    ] = True  # we just use the one good value we have
                    cell_results[row_hdr]["VALUE"] = value

                elif (  # ----------------- 1, 1, 0, 1 ----------------
                    (row_hdr in row_hdrs)
                    and (col_hdr in col_hdrs)
                    and (col_hdr not in row_hdrs)
                    and (row_hdr in col_hdrs)
                ):
                    cell_results[row_hdr][ROW_COL_KEY] = value = check_for_none(
                        row_dict[row_hdr][col_hdr]
                    )
                    cell_results[row_hdr][COL_ROW_KEY] = f"X {col_hdr}"
                    cell_results[row_hdr][
                        MATCH_KEY
                    ] = True  # we just use the one good value we have
                    cell_results[row_hdr]["VALUE"] = value

                elif (  # ----------------- 1, 1, 1, 0 ----------------
                    (row_hdr in row_hdrs)
                    and (col_hdr in col_hdrs)
                    and (col_hdr in row_hdrs)
                    and (row_hdr not in col_hdrs)
                ):
                    cell_results[row_hdr][ROW_COL_KEY] = value = check_for_none(
                        row_dict[row_hdr][col_hdr]
                    )
                    cell_results[row_hdr][COL_ROW_KEY] = f"X {row_hdr}"
                    cell_results[row_hdr][
                        MATCH_KEY
                    ] = True  # we just use the one good value we have
                    cell_results[row_hdr]["VALUE"] = value

                elif (  # ---------------- 1, 1, 1, 1 ---------------
                    (row_hdr in row_hdrs)
                    and (col_hdr in col_hdrs)
                    and (col_hdr in row_hdrs)
                    and (row_hdr in col_hdrs)
                ):
                    # both headers are valid in the default iand transpositional positions
                    cell_results[row_hdr][ROW_COL_KEY] = check_for_none(
                        row_dict[row_hdr][col_hdr]
                    )
                    cell_results[row_hdr][COL_ROW_KEY] = check_for_none(
                        col_dict[row_hdr][col_hdr]
                    )
                    cell_results[row_hdr][MATCH_KEY] = (
                        cell_results[row_hdr][ROW_COL_KEY]
                        == cell_results[row_hdr][COL_ROW_KEY]
                    )
                    if cell_results[row_hdr][MATCH_KEY] == True:
                        cell_results[row_hdr]["VALUE"] = cell_results[row_hdr][
                            ROW_COL_KEY
                        ]

                # at this point, one or both of the headers is not valid in either the default or
                #   transpositional positions

                # print(f"{'Vendor 1':<15} {'Vendor 2':<15} {'row,col'}:<15 {'col,row'}:<15 {'Match':<5} {'Value':<5}")
                try:
                    print(
                        f"{row_hdr[0:15]:<15} {col_hdr[0:15]:<15} {cell_results[row_hdr][ROW_COL_KEY][0:15]:<15} {cell_results[row_hdr][COL_ROW_KEY][0:15]:<15} {cell_results[row_hdr][MATCH_KEY]:<5} {cell_results[row_hdr]['VALUE']:<5}"
                    )
                except TypeError:
                    print()
                    print("Type Error:")
                    print(
                        f"row_hdr: {row_hdr} col_hdr: {col_hdr} cell_results: {cell_results}"
                    )
                    print()

                results[row_hdr] = cell_results

    print("=============================== Done ===============================")
    return results


# ################### Print Report ###################


def print_report(
    results: List[Tuple[Tuple[str, str], Tuple[str, str], str, str, str]]
) -> None:
    """
    Print the comparison results in a formatted report.
    """
    return

    for result in results:
        print(f"Comparison: {result[0]} vs {result[1]}")
        print(
            f"Is Equal: {result[2]:<20}Value: {result[3]:<20}Transposed Value: {result[4]}"
        )
        print()


# ################### Load Matrix ###################


def load_matrix(filename: str, table_name: str) -> List[List[str]]:

    matrix: List[List[str]] = []

    if not os.path.exists(filename):
        print(f"File '{filename}' does not exist.")
    elif not os.access(filename, os.R_OK):
        print(f"Cannot read file '{filename}'.")
    else:
        wb = load_workbook(filename)
        ws = wb.worksheets[0]
        table = ws[table_name]
        matrix = excel_table_to_matrix(table)

    return matrix


# ################## print dictionary ##################


def print_dict(dictionary: Dict[str, Dict[str, str]], title: str) -> None:
    print(f"########## {title} ##########")
    for key, value in dictionary.items():
        print(f"{key}:")
        for k, v in value.items():
            print(f"  {k}: {v}")
        print()
    print()


# ################### MAIN ###################


def main():
    # read in the workbook and get the first worksheet
    if (input_data := load_matrix(FILENAME_IN, TABLE_NAME)) is None:
        return 1

    # Transform data into dictionaries
    row_dict, col_dict = transform_data(input_data)

    # Print the two dictionaries
    print_dict(row_dict, "Row Dictionary")
    print_dict(col_dict, "Column Dictionary")

    # Compare values and generate results
    comparison_results = compare_values2(row_dict, col_dict)
    print("##################### Comparison Results #####################")
    print(f"Items successfully compared: {len(comparison_results)}")
    for item in comparison_results:
        print(item)
    print()

    # Print the report
    print_report(comparison_results)

    return 0


if __name__ == "__main__":
    exit(main())
