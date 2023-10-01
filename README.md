# pyetest

The following introduces a python library targeted at filling in spreadsheets to simplify writing pass/fail tests for electronics testing. This method generates easy to maintain test description and acceptance test software with easily defined measurements and test criteria. The method relies on using spreadsheets allowing users who don't program to define tests and acceptance criteria. The test data reports are generated as intermediate files before the pass/fail tests simplifying
reporting and diagnosis.


## pyetest Operation

### Steps
+ Load spreadsheet template with relationships defined
+ Fill the spreadsheet values, store as a new export. Have all intermediate values stored
+ Reload spreadsheet to check thresholds. This way the requirements only knows about the high/low targets
+ Steps are read, create spreadsheet


### Inputs

#### Data template
This is an xlsx spreadsheet which contains the test names, data source, and field type. The field type specifies if the data is measured, fixed, or calculated. Measured values will be filled in during the test. Fixed values are used in the spreadsheet for calculations of other values such as the number of bits in an ADC. Calculated values have a formula in the value field. 

#### Test template
A spreadsheet which defines the pass/fail criteria. The structure is open for extension but includes at a minimum the field name, a value field that will be filled in, and a passes column that contains an expression (usually something like `=E1=B1` or `=E1>=C1 & E1<=D1`).

### Outputs

#### Measurement Results
This is the data template file with the measured values filled in. All the calculation expressions are retained in the file.

#### Test Results
This is the test template file with the values filled in. The boolean passes values can then be read out to determine which tests passed.


## Procedure
1. Take all the measurements needed and generate a dataframe of data. This can be made all sorts of ways. I typically add a source column to the data template to parse the tag names to define the tests. How to setup these measurements is a separate section of a test suite and is essentially tied to the system you're using.
2. Load the template files and call pyetest_run
3. Save the data workbook and test results for future retrieval and analysis
4. Turn the tests into a data frame and return a pass/fail analysis

```python
import pandas
import openpyxl
import pyetest

data_template_file = "data_template.xlsx"
measurement_result_file = "measurement_result.xlsx"
measured_data_file = "dummy-values.xlsx"
test_template_file = "test_template.xlsx"
test_result_file = "test-results.xlsx"

data_template_workbook = openpyxl.open(data_template_file)
test_template_workbook = openpyxl.open(test_template_file)
dummy_values_df = pandas.read_excel(measured_data_file)

data_workbook, test_workbook = pyetest.pyetest_run(data_template_workbook, test_template_workbook, dummy_values_df)
test_results_df = pyetest.worksheet_to_df(test_workbook.active)
print(len(test_results_df[test_results_df["passes"] == False]) == 0)
print(test_results_df[test_results_df["passes"] == False])

test_workbook.save(test_result_file)
data_workbook.save(measurement_result_file)
```

```
False
         name expected value  unit  min  max value passes
         15   +5 Fault              0  None  NaN  NaN     1  False
         18  DAQ Fault              0  None  NaN  NaN     1  False
```

