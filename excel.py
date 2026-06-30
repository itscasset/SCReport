import csv

with open(r"C:\Users\rattanachote\Downloads\VRptExpension.csv", encoding="utf-8") as f:
    reader = csv.reader(f)

    expected = None

    for i, row in enumerate(reader, start=1):
        if expected is None:
            expected = len(row)
            print(expected)

        if len(row) != expected:
            print(i, len(row), row)
            print('some problem')