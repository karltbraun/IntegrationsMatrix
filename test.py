class Vendor:
    def __init__(self, name):
        self.name = name
        self.integrations = set()
        self.integrations_count = 0
        self.limited = set()
        self.limited_count = 0

    def add_integration(self, integration):
        self.integrations.add(integration)
        self.integrations_count += 1

    def add_limited(self, integration):
        self.limited.add(integration)
        self.limited_count += 1


def foo(matrix_in: list[list[str]]):

    vendors = []
    row_hdrs = matrix_in[0][1:]
    col_hdrs = [row[0] for row in matrix_in]
    col_hdrs = col_hdrs[1:]

    for v1 in row_hdrs:
        vendor = Vendor(v1)
        for v2 in col_hdrs:
            print(f"checking {v1} v {v2}")
            if v1 == v2:
                continue
            # v1 is a valid row_hdr, v2 is a valid col_hdr
            value = matrix_in[row_hdrs.index(v1)][col_hdrs.index(v2)]
            if value == "YES":
                vendor.add_integration(v2)
            elif value == "LIMITED":
                vendor.add_limited(v2)
        vendors.append(vendor)

    return vendors


def main():
    print("Running main()")
    matrix = [
        ["Vendors", "Vendor_A", "Vendor_B", "Vendor_C", "Vendor_E"],
        ["Vendor_A", "YES", "NO", "LIMITED"],
        ["Vendor_B", "NO", "YES", "NO"],
        ["Vendor_F", "LIMITED", "NO", "YES"],
    ]

    vendors = foo(matrix)

    print("Input matrix:")
    for row in matrix:
        print(row)

    print("\nOutput:")
    for vendor in vendors:
        print(f"Vendor: {vendor.name}")
        print(f"  Number of integrations: {vendor.integrations_count}")
        print(f"    Integrations: {vendor.integrations}")
        print(f"  Number of limited integrations: {vendor.limited_count}")
        print(f"    Limited: {vendor.limited}")
        print()


# ########################### MAIN #########################


if __name__ == "__main__":
    exit(main())
