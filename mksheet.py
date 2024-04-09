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
    # given a list of siteID (4-5 character identifiers), we are going to
    # create a list of data_items, where each data_item is a list of strings
    # where the first item is the siteID, the second item is either "Vendor" or "Use",
    # and the remaining items are blank characters, one for each of the remaining data column headers
    # (so, len(column_headers) - 2).
    # Each site will be added to the list of data_items twice.  Once with "Vendor" as the second item,
    # and once with "Use" as the second item.
    # The list of data_items will be returned.
    # This can be used by the caller to create a csv file.

    lst_data_items = []
    for siteID in lst_sites:
        lst_data_items.append([siteID, "Vendor"])
        lst_data_items.append([siteID, "Use"])

    return lst_data_items


def main():
    num_data_cols = len(column_headers)  # number of data columns to add to each line

    # get the list of lines to write to the csv file
    lst_data_items = do_thing(siteIDs, num_data_cols)

    print(
        "*******************************************************************************"
    )

    line = ""
    for hdr in column_headers:
        line += f"{hdr},"
    print(line)

    for data_item in lst_data_items:
        line = ""
        for data in data_item:
            line += f"{data},"
        print(line)

    print(
        "-------------------------------------------------------------------------------"
    )

    # write the csv file

    with open(FILENAME_OUTPUT, "w", newline="") as csvfile:
        writer = csv.writer(csvfile)
        print(column_headers)
        writer.writerow(column_headers)
        for data in lst_data_items:
            print(data)
            writer.writerow(data)
        print(
            "*******************************************************************************"
        )


if __name__ == "__main__":
    exit(main())
