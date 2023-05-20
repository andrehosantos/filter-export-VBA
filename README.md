## VBA Code Description

The following VBA code is a macro named `FilterAndExport` that I have developed for performing filtering and exporting tasks based on specified criteria. This code is part of my portfolio and showcases my VBA programming skills. Here is a breakdown of what the code does:

1. The code declares variables and an array to store values.
2. It reads values from "Sheet2" in the range "A2:A100" and assigns them to the array.
3. It loops through each value in the array and applies it as a filter to "Sheet1" if the value is not blank.
4. If the value is numeric, it converts it to a text value; otherwise, it uses the value as-is.
5. The code filters "Sheet1" based on the criteria and copies the visible cells.
6. If a file with the same name already exists in the specified directory, it is deleted.
7. A new workbook is created to export the filtered data as a CSV file.
8. The filtered data is pasted in the new workbook, and specific formatting is applied.
9. The new workbook is saved with the criteria as the filename in the specified directory.
10. The new workbook is closed.
11. The filter on "Sheet1" is cleared.
12. Finally, the filters on "Sheet2" are cleared.

This code demonstrates my ability to automate data filtering and exporting tasks using VBA. It showcases my proficiency in working with arrays, loops, conditional statements, and manipulating Excel workbooks and ranges.
