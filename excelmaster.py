# Data manipulation of excel or csv files by Rophi
import pandas as pd

values = []
index = []
color = ["red", "cyan", "purple", "grey", "orange", "pink", "blue", "green", "brown", "yellow"]

# Update code to check for file extension
# If file is a csv it will be converted to an excel because the syntax for excel in pandas is a lot much easier
df = pd.read_excel('ref.xlsx')


# This functino checks whether the current row is in the duplicates list ie the 'values' list
def checker(indx):
    if len(values) > 0:
        for i in range(0, len(values)):
            if (df.values[indx] == values[i]).all():
                return True
        return False
    return False


# This for loop checks for duplicates
# Stores the duplicate value in the values list and the index of the duplicates in the index list
for row1 in range(0, len(df.index)):
    dup_index = []

    if not checker(row1):
        print(row1)
        dup_index.append(row1)
        for row2 in range(row1+1, len(df.index)):
            if (df.values[row1] == df.values[row2]).all():
                dup_index.append(row2)

    if len(dup_index) > 1:
        values.append(df.values[row1])
        index.append(dup_index)


# Highlights duplicate and creates a new file 'highlight.xlsx'
def highlight():
    df.style.apply(colorscheme).to_excel('highlight.xlsx', index=False)


# Returns an array which is used to determine the colors of the duplicate rows
def colorscheme(cell):
    ls = []

    for i in range(0, len(df.index)):
        for j in range(0, len(index)):
            if i in index[j]:
                ls.append(f"background-color:{color[j]}")
                break
            elif j == len(index) - 1:
                ls.append("")

    return ls


# Removes duplicate rows and creates a new file 'remove.xlsx'
def remove():
    ls = []
    for arr in index:
        for i in range(1, len(arr)):
            ls.append(arr[i])

    df.drop(ls).to_excel("remove.xlsx", index=False)


highlight()
