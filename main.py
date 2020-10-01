import math
import os
import random
import xlrd
import xlwt


def writeCols(cols):
    for j in range(len(cols)):
        sheetW.write(0, j, cols[j])
    writebook.save("Distances.xls")


# populates excel file with random values for data points
def randomPoints():
    for i in range(1, pointCount):
        for j in range(len(columns)):
            sheetW.write(i, j, random.randint(1, 100))
    writebook.save("Distances.xls")


# calculate and write Euclidean distances for coordinates
def eucDist():

    columns.append("Euclidean Distance")

    for i in range(1, pointCount):
        x1 = sheetR.cell_value(i, columns.index("X1"))
        y1 = sheetR.cell_value(i, columns.index("Y1"))
        x2 = sheetR.cell_value(i, columns.index("X2"))
        y2 = sheetR.cell_value(i, columns.index("Y2"))

        euc = math.sqrt(math.pow((x2 - x1), 2) + math.pow((y2 - y1), 2))
        sheetW.write(i, columns.index("Euclidean Distance"), euc)
        sheetW.col(columns.index("Euclidean Distance")).width = 4500

    writebook.save("Distances.xls")


# calculate and write Manhattan distances for coordinates
def manhattanDist():
    columns.append("Manhattan Distance")

    for i in range(1, pointCount):
        x1 = sheetR.cell_value(i, columns.index("X1"))
        y1 = sheetR.cell_value(i, columns.index("Y1"))
        x2 = sheetR.cell_value(i, columns.index("X2"))
        y2 = sheetR.cell_value(i, columns.index("Y2"))

        man = abs(x2 - x1) + abs(y2 - y1)
        sheetW.write(i, columns.index("Manhattan Distance"), man)
        sheetW.col(columns.index("Manhattan Distance")).width = 4500

    writebook.save("Distances.xls")


# initialize writing to workbook with xlwt
writebook = xlwt.Workbook()
sheetW = writebook.add_sheet("Sheet1", True)

# global variables
columns = ["X1", "Y1", "X2", "Y2"]      # list for column names

pointCols = 4                           # number of columns that the coordinate data points' information exists in
pointCount = 20                            # variable for number of data points

# write current column names
writeCols(columns)
randomPoints()

# initialize reading from workbook with xlrd
#   (done separately from xlwt initialization so that workbook is written to and saved with random points first,
#   then opened to be read from
path = os.path.expanduser("Distances.xls")
readbook = xlrd.open_workbook(path)
sheetR = readbook.sheet_by_index(0)

eucDist()
manhattanDist()

# rewrite current column names (after the edition of the euclidean distance column
writeCols(columns)

