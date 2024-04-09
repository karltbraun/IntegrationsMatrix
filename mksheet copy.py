import csv

FILENAME_OUTPUT = "test.csv"

# create a list of siteIDs and column headers
siteIDs = [
    "BAJA",
    "BPSN",
    "CASA",
    "CASH",
    "CASS",
    "CAT2",
    "CAYG",
    "CNTO",
    "COCN",
    "EBSJ",
    "EBYA",
    "FLO2",
    "FLOA",
    "FLOC",
    "ILCA",
    "ILCW",
    "KYCL",
    "MDAJ",
    "MXDM",
    "MXEC",
    "MXFB",
    "NENC",
    "NJSH 406",
    "NJSH 407",
    "NWK2",
    "PATC",
    "PATM",
    "RTGM",
    "RTYA",
    "SCCE",
    "SCOL",
    "SWAZ",
    "TNSS",
    "TNWW",
    "TXDB",
    "TXDC",
    "VIGP",
]

column_headers = [
    "SiteID",
    "Vendor/Use",
    "dumper",
    "cutter",
    "slicer",
    "r-optical-srt",
    "r-v-inspect",
    "smartwash",
    "washline-2",
    "dryers",
    "dryer logistics",
    "po-dump",
    "po-elev",
    "bag-scales",
    "bag-mp-ins",
    "bag-bagger",
    "bag-labeler",
    "bag-label-ver",
    "bag-CW",
    "bag-MD",
    "bag-v-inspect",
    "bag-packing",
    "tray-scales",
    "tray-mp-ins",
    "tray-pack",
    "tray-tamp",
    "tray-labeler",
    "tray-label-v",
    "tray-seal",
    "tray-CW",
    "tray-MD",
    "tray-v-insp",
    "case-form",
    "case-label",
    "case-label-v",
    "case-packing",
    "palletizer",
    "pallet-label",
    "pallet-label-v",
]


def do_thing(lst_sites: list[str], num_data_cols: int) -> list[str]:
    # given a list of siteID (4-5 character identifiers), we are going go
    # create a list of comma separated values to output to a csv file.
    # The calling routine will supply the list of siteIDs, and the number of
    # data columns which this routine will add to the lines as blank fields.
    # The calling routine will take care of writing the csv file,
    # including the header row.

    lst_lines = []
    # blank_data = ", " * (num_data_cols - 1)  # create a string of commas for blank data
    # create a list n items that all contain a single blank character, where 'n' == (num_data_cols - 1)
    blank_data = [" " for i in range(num_data_cols - 1)]
    # copy blank_data to another list, with the string "  Vendor" as the first item
    blank_data1 = blank_data.copy()
    blank_data1.insert(0, "  Vendor")
    # copy blank_data to another list, with the string "  Use" as the first item
    blank_data2 = blank_data.copy()
    blank_data2.insert(0, "  Use")

    for site in lst_sites:
        lst_lines.append(line)

    return lst_lines


def main():
    num_data_cols = len(column_headers)  # number of data columns to add to each line

    # get the list of lines to write to the csv file
    lst_lines = do_thing(siteIDs, num_data_cols)

    print(
        "*******************************************************************************"
    )

    line = ""
    for hdr in column_headers:
        line += f"{hdr},"
    line = line[:-1]  # remove the last comma
    print(line)

    for line in lst_lines:
        print(line)
    print(
        "-------------------------------------------------------------------------------"
    )

    # write the csv file
    lst_data = []
    for line in lst_lines:
        data = line.split(",")
        lst_data.append(data)

    with open(FILENAME_OUTPUT, "w", newline="") as csvfile:
        writer = csv.writer(csvfile)
        print(column_headers)
        writer.writerow(column_headers)
        for data in lst_data:
            print(data)
            writer.writerow([data])
        print(
            "*******************************************************************************"
        )


if __name__ == "__main__":
    exit(main())
