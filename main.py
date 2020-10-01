import math
import os
import random
import xlrd
import xlwt

# This program creates or modifies an excel file to be populated with a variable number of randomly generated
# cartesian coordinate pairs. Afterwards, The Euclidean and Manhattan distances are calculated for each coordinate pair.
# Coordinate pairs are defined by their X and Y values, so that each pair is in the form: (x1, y1), (x2, y2).
# *NOTE* the program will not execute if the excel file is open. You must exit the excel file before running.


# global variable for file to be worked with
file = "Distances.xls"


# function to write the column names to the excel file
def writeCols(cols):
    for j in range(len(cols)):
        # write only on the top row, so i = 0
        sheetW.write(0, j, cols[j])
    writebook.save(file)


# populates excel file with random values for data points
def randomPoints():
    # iterate through pointCount size for rows excluding labels
    for i in range(1, pointCount):
        # iterate through columns for columns (only length 4 when this is called)
        for j in range(len(columns)):
            # write random int between 1 and 100
            sheetW.write(i, j, random.randint(1, 100))
    # output message for clarity
    print("\nSuccessfully populated '" + file + "' with " + str(pointCount) + " randomly generated cartesian coordinate pairs.")
    writebook.save(file)


# calculate and write Euclidean distances for coordinates
def eucDist():
    # add euclidean column to be written
    columns.append("Euclidean Distance")

    # iterate through data excluding labels
    for i in range(1, pointCount):
        # assign variables for calculation according to corresponding column names
        x1 = sheetR.cell_value(i, columns.index("X1"))
        y1 = sheetR.cell_value(i, columns.index("Y1"))
        x2 = sheetR.cell_value(i, columns.index("X2"))
        y2 = sheetR.cell_value(i, columns.index("Y2"))
        # calculate euclidean distance based on above variables
        euc = math.sqrt(math.pow((x2 - x1), 2) + math.pow((y2 - y1), 2))
        # write calculated euclidean distances to corresponding column given the current row i
        sheetW.write(i, columns.index("Euclidean Distance"), euc)
        # change column width to accommodate for longer label
        sheetW.col(columns.index("Euclidean Distance")).width = 4500

    # output message for clarity
    print("Successfully calculated the Euclidean distances for the " + str(pointCount) + " coordinate pairs.")
    writebook.save(file)


# calculate and write Manhattan distances for coordinates
def manhattanDist():
    # add manhattan column to be written
    columns.append("Manhattan Distance")
    # iterate through data excluding labels
    for i in range(1, pointCount):
        # assign variables for calculation according to corresponding column names
        x1 = sheetR.cell_value(i, columns.index("X1"))
        y1 = sheetR.cell_value(i, columns.index("Y1"))
        x2 = sheetR.cell_value(i, columns.index("X2"))
        y2 = sheetR.cell_value(i, columns.index("Y2"))
        # calculate manhattan distance based on above variables
        man = abs(x2 - x1) + abs(y2 - y1)
        # write calculated manhattan distances to corresponding column given the current row i
        sheetW.write(i, columns.index("Manhattan Distance"), man)
        # change column width to accommodate for longer label
        sheetW.col(columns.index("Manhattan Distance")).width = 4500

    # output message for clarity
    print("Successfully calculated the Manhattan distances for the " + str(pointCount) + " coordinate pairs.")
    writebook.save(file)


# initialize writing to workbook with xlwt
writebook = xlwt.Workbook()
# overwrite = True so that points can be rerolled
sheetW = writebook.add_sheet("Sheet1", True)

# global variables
columns = ["X1", "Y1", "X2", "Y2"]      # list for column names

pointCols = 4                           # number of columns that the coordinate data points' information exists in
pointCount = 20                         # variable for number of data points

# write current column names
writeCols(columns)
# populate with random data points
randomPoints()

# initialize reading from workbook with xlrd
#   (done separately from xlwt initialization so that workbook is written to and saved with random points first,
#   then opened to be read from afterwards)
path = os.path.expanduser(file)
readbook = xlrd.open_workbook(path)
sheetR = readbook.sheet_by_index(0)

# call calculation functions
eucDist()
manhattanDist()

# rewrite current column names (after the edition of the euclidean distance and manhattan distance columns)
writeCols(columns)

