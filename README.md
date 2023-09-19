# pyetest

Tests
+ Comparision to fixed value
+ Comparision to measured value
+ Comparision to calculated value

# Objects
+ test description input table
	+ should include measurements
+ measured values table
+ processed values table
+ output table

# Process
+ Create measurement table from test description table
+ Expand depending on test description
	+ Expand strings for expressions? -> this can be aweful, abstract code execution
	+ Use excel format commands?

# openpyxl
+ Can input formulas but relies on spreadsheet program to evaluate them.
+ Call libreoffice with "libreoffice --calc --convert-to xlsx /tmp/out.xlsx" to update them.


# Steps
+ Load spreadsheet template with relationships defined
+ Fill the spreadsheet values, store as a new export. Have all intermediate values stored
+ Reload spreadsheet to check thresholds. This way the requirements only knows about the high/low targets
+ Steps are read, create spreadsheet


# Example Process
+ Pull all status from a device
+ There may be processes like changing the commanded levels, these are each an input into the measured values table
+ Call the pyetest script to expand the template and run the pass/fail tests.
