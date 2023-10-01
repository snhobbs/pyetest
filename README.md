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

|    | name                | value                                      | unit   | type       |
|---:|:--------------------|:-------------------------------------------|:-------|:-----------|
|  0 | version             | 0.8.X                                      | nan    | measured   |
|  1 | firmware_version    | 0.1.8                                      | nan    | measured   |
|  2 | compile_time        | 09:09:00                                   | nan    | measured   |
|  3 | compile_date        | September 9th, 2023                        | nan    | measured   |
|  4 | slave_address       | 246                                        | nan    | measured   |
|  5 | p3_3_micro_volts    | 3260000                                    | uV     | measured   |
|  6 | p3_3_reading        | 2024                                       | ADC    | measured   |
|  7 | +3.3GAIN            | =1/2                                       | nan    | fixed      |
|  8 | +3.3OFFSET          | 0                                          | V      | fixed      |
|  9 | +3.3V               | =B8/4096*3.3/B9                            | V      | calculated |
| 10 | +3.3ERROR           | =ABS(B11/24-1)*100                         | %      | calculated |
| 11 | +3.3COMPARISONERROR | =ABS(B11/(B7/1000000)-1)*100               | %      | calculated |
| 12 | p23_micro_volts     | 23000000                                   | uV     | measured   |
| 13 | p23_reading         | 1200                                       | ADC    | measured   |
| 14 | +23GAIN             | =1/25                                      | nan    | fixed      |
| 15 | +23OFFSET           | 0                                          | V      | fixed      |
| 16 | +23V                | =B15/4096*3.3/B16                          | V      | calculated |
| 17 | +23ERROR            | =ABS(B18/23.3-1)*100                       | %      | calculated |
| 18 | +23COMPARISONERROR  | =ABS(B18/(B14/1000000)-1)*100              | %      | calculated |
| 19 | p5_micro_volts      | 5000000                                    | uV     | measured   |
| 20 | p5_reading          | 2500                                       | ADC    | measured   |
| 21 | +5OFFSET            | 0                                          | V      | fixed      |
| 22 | +5GAIN              | =5/2                                       | nan    | fixed      |
| 23 | +5                  | =B22/4096*B24*3.3                          | V      | calculated |
| 24 | +5ERROR             | =ABS((B25/5-1)*100)                        | %      | calculated |
| 25 | +5COMPARISONERROR   | =ABS(B25/(B21/1000000)-1)*100              | %      | calculated |
| 26 | fault_status        | 9                                          | nan    | measured   |
| 27 | +5 Fault            | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,0)),2)  | nan    | calculated |
| 28 | +24 Fault           | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,1)),2)  | nan    | calculated |
| 29 | Error Code 2        | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,2)),2)  | nan    | calculated |
| 30 | DAQ Fault           | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,3)),2)  | nan    | calculated |
| 31 | Error Code 4        | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,4)),2)  | nan    | calculated |
| 32 | MW Fault            | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,5)),2)  | nan    | calculated |
| 33 | Error Code 6        | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,6)),2)  | nan    | calculated |
| 34 | Error Code 7        | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,7)),2)  | nan    | calculated |
| 35 | Visible Fault       | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,8)),2)  | nan    | calculated |
| 36 | Temperature Fault   | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,9)),2)  | nan    | calculated |
| 37 | Moisture Fault      | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,10)),2) | nan    | calculated |
| 38 | Watchdog Fault      | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,11)),2) | nan    | calculated |
| 39 | Hardware Fault      | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,12)),2) | nan    | calculated |
| 40 | I2C Fault           | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,13)),2) | nan    | calculated |
| 41 | Error Code 14       | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,14)),2) | nan    | calculated |
| 42 | Modbus Fault        | =MOD(_xlfn.FLOOR.MATH(B$28/POWER(2,15)),2) | nan    | calculated |



#### Test template
A spreadsheet which defines the pass/fail criteria. The structure is open for extension but includes at a minimum the field name, a value field that will be filled in, and a passes column that contains an expression (usually something like `=E1=B1` or `=E1>=C1 & E1<=D1`).


|    | name              | expected value                             | unit   |       min |       max | value               | passes               |
|---:|:------------------|:-------------------------------------------|:-------|----------:|----------:|:--------------------|:---------------------|
|  0 | version           | 0.8.X                                      | nan    | nan       | nan       | 0.8.X               | =B2=F2               |
|  1 | firmware_version  | 0.1.8                                      | nan    | nan       | nan       | 0.1.8               | =B3=F3               |
|  2 | compile_time      | 09:09:00                                   | nan    | nan       | nan       | 09:09:00            | =B4=F4               |
|  3 | compile_date      | September 9th, 2023                        | nan    | nan       | nan       | September 9th, 2023 | =B5=F5               |
|  4 | slave_address     | 246                                        | nan    | nan       | nan       | 246                 | =B6=F6               |
|  5 | p3_3_micro_volts  | 3300000                                    | uV     |   3.2e+06 |   3.4e+06 | 3260000             | =F7>=D7 & F7<=E7     |
|  6 | p3_3_reading      | 2100                                       | ADC    | nan       | nan       | 2024                | nan                  |
|  7 | +3.3ERROR         | 0                                          | %      |   0       |   3       | 86.4111328125       | =F9>=D9 & F9<=E9     |
|  8 | p23_micro_volts   | 23000000                                   | uV     |   1.8e+07 |   3e+07   | 23000000            | =F10>=D10 & F10<=E10 |
|  9 | p23_reading       | 1200                                       | ADC    | nan       | nan       | 1200                | nan                  |
| 10 | +23ERROR          | 0                                          | %      |   0       |   3       | 3.73357027896994    | =F12>=D12 & F12<=E12 |
| 11 | p5_micro_volts    | 5000000                                    | uV     |   4.9e+06 |   5.1e+06 | 5000000             | =F13>=D13 & F13<=E13 |
| 12 | p5_reading        | 2500                                       | ADC    | nan       | nan       | 2500                | nan                  |
| 13 | +5ERROR           | 0                                          | %      |   0       |   3       | 0.7080078125        | =F15>=D15 & F15<=E15 |
| 14 | fault_status      | 0                                          | nan    | nan       | nan       | 9                   | nan                  |
| 15 | +5 Fault          | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,0)),2)  | nan    | nan       | nan       | 1                   | =B17=F17             |
| 16 | +24 Fault         | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,1)),2)  | nan    | nan       | nan       | 0                   | =B18=F18             |
| 17 | Error Code 2      | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,2)),2)  | nan    | nan       | nan       | 0                   | =B19=F19             |
| 18 | DAQ Fault         | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,3)),2)  | nan    | nan       | nan       | 1                   | =B20=F20             |
| 19 | Error Code 4      | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,4)),2)  | nan    | nan       | nan       | 0                   | =B21=F21             |
| 20 | MW Fault          | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,5)),2)  | nan    | nan       | nan       | 0                   | =B22=F22             |
| 21 | Error Code 6      | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,6)),2)  | nan    | nan       | nan       | 0                   | =B23=F23             |
| 22 | Error Code 7      | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,7)),2)  | nan    | nan       | nan       | 0                   | =B24=F24             |
| 23 | Visible Fault     | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,8)),2)  | nan    | nan       | nan       | 0                   | =B25=F25             |
| 24 | Temperature Fault | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,9)),2)  | nan    | nan       | nan       | 0                   | =B27=F27             |
| 25 | Moisture Fault    | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,10)),2) | nan    | nan       | nan       | 0                   | nan                  |
| 26 | Watchdog Fault    | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,11)),2) | nan    | nan       | nan       | 0                   | =B28=F28             |
| 27 | Hardware Fault    | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,12)),2) | nan    | nan       | nan       | 0                   | =B29=F29             |
| 28 | I2C Fault         | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,13)),2) | nan    | nan       | nan       | 0                   | =B30=F30             |
| 29 | Error Code 14     | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,14)),2) | nan    | nan       | nan       | 0                   | =B31=F31             |
| 30 | Modbus Fault      | =MOD(_xlfn.FLOOR.MATH(B$16/POWER(2,15)),2) | nan    | nan       | nan       | 0                   | =B32=F32             |



#### Measurements Data Frame
Data taken from all our sources.

|    | name             | value               | unit   | type     |
|---:|:-----------------|:--------------------|:-------|:---------|
|  0 | version          | 0.8.X               | nan    | measured |
|  1 | firmware_version | 0.1.8               | nan    | measured |
|  2 | compile_time     | 09:09:00            | nan    | measured |
|  3 | compile_date     | September 9th, 2023 | nan    | measured |
|  4 | slave_address    | 246                 | nan    | measured |
|  5 | p3_3_micro_volts | 3260000             | uV     | measured |
|  6 | p3_3_reading     | 2024                | ADC    | measured |
|  7 | p23_micro_volts  | 23000000            | uV     | measured |
|  8 | p23_reading      | 1200                | ADC    | measured |
|  9 | p5_micro_volts   | 5000000             | uV     | measured |
| 10 | p5_reading       | 2500                | ADC    | measured |
| 11 | fault_status     | 9                   | nan    | measured |


### Outputs

#### Measurement Results
This is the data template file with the measured values filled in. All the calculation expressions are retained in the file.

|    | name                | value               | unit   | type       |
|---:|:--------------------|:--------------------|:-------|:-----------|
|  0 | version             | 0.8.X               | nan    | measured   |
|  1 | firmware_version    | 0.1.8               | nan    | measured   |
|  2 | compile_time        | 09:09:00            | nan    | measured   |
|  3 | compile_date        | September 9th, 2023 | nan    | measured   |
|  4 | slave_address       | 246                 | nan    | measured   |
|  5 | p3_3_micro_volts    | 3260000             | uV     | measured   |
|  6 | p3_3_reading        | 2024                | ADC    | measured   |
|  7 | +3.3GAIN            | 0.5                 | nan    | fixed      |
|  8 | +3.3OFFSET          | 0                   | V      | fixed      |
|  9 | +3.3V               | 3.261328125         | V      | calculated |
| 10 | +3.3ERROR           | 86.4111328125       | %      | calculated |
| 11 | +3.3COMPARISONERROR | 0.0407400306748462  | %      | calculated |
| 12 | p23_micro_volts     | 23000000            | uV     | measured   |
| 13 | p23_reading         | 1200                | ADC    | measured   |
| 14 | +23GAIN             | 0.04                | nan    | fixed      |
| 15 | +23OFFSET           | 0                   | V      | fixed      |
| 16 | +23V                | 24.169921875        | V      | calculated |
| 17 | +23ERROR            | 3.73357027896994    | %      | calculated |
| 18 | +23COMPARISONERROR  | 5.0866168478261     | %      | calculated |
| 19 | p5_micro_volts      | 5000000             | uV     | measured   |
| 20 | p5_reading          | 2500                | ADC    | measured   |
| 21 | +5OFFSET            | 0                   | V      | fixed      |
| 22 | +5GAIN              | 2.5                 | nan    | fixed      |
| 23 | +5                  | 5.035400390625      | V      | calculated |
| 24 | +5ERROR             | 0.7080078125        | %      | calculated |
| 25 | +5COMPARISONERROR   | 0.7080078125        | %      | calculated |
| 26 | fault_status        | 9                   | nan    | measured   |
| 27 | +5 Fault            | 1                   | nan    | calculated |
| 28 | +24 Fault           | 0                   | nan    | calculated |
| 29 | Error Code 2        | 0                   | nan    | calculated |
| 30 | DAQ Fault           | 1                   | nan    | calculated |
| 31 | Error Code 4        | 0                   | nan    | calculated |
| 32 | MW Fault            | 0                   | nan    | calculated |
| 33 | Error Code 6        | 0                   | nan    | calculated |
| 34 | Error Code 7        | 0                   | nan    | calculated |
| 35 | Visible Fault       | 0                   | nan    | calculated |
| 36 | Temperature Fault   | 0                   | nan    | calculated |
| 37 | Moisture Fault      | 0                   | nan    | calculated |
| 38 | Watchdog Fault      | 0                   | nan    | calculated |
| 39 | Hardware Fault      | 0                   | nan    | calculated |
| 40 | I2C Fault           | 0                   | nan    | calculated |
| 41 | Error Code 14       | 0                   | nan    | calculated |
| 42 | Modbus Fault        | 0                   | nan    | calculated |




#### Test Results
This is the test template file with the values filled in. The boolean passes values can then be read out to determine which tests passed.

|    | name              | expected value      | unit   |       min |       max | value               |   passes |
|---:|:------------------|:--------------------|:-------|----------:|----------:|:--------------------|---------:|
|  0 | version           | 0.8.X               | nan    | nan       | nan       | 0.8.X               |        1 |
|  1 | firmware_version  | 0.1.8               | nan    | nan       | nan       | 0.1.8               |        1 |
|  2 | compile_time      | 09:09:00            | nan    | nan       | nan       | 09:09:00            |        1 |
|  3 | compile_date      | September 9th, 2023 | nan    | nan       | nan       | September 9th, 2023 |        1 |
|  4 | slave_address     | 246                 | nan    | nan       | nan       | 246                 |        1 |
|  5 | p3_3_micro_volts  | 3300000             | uV     |   3.2e+06 |   3.4e+06 | 3260000             |        1 |
|  6 | p3_3_reading      | 2100                | ADC    | nan       | nan       | 2024                |      nan |
|  7 | +3.3ERROR         | 0                   | %      |   0       |   3       | 86.4111328125       |        1 |
|  8 | p23_micro_volts   | 23000000            | uV     |   1.8e+07 |   3e+07   | 23000000            |        1 |
|  9 | p23_reading       | 1200                | ADC    | nan       | nan       | 1200                |      nan |
| 10 | +23ERROR          | 0                   | %      |   0       |   3       | 3.73357027896994    |        1 |
| 11 | p5_micro_volts    | 5000000             | uV     |   4.9e+06 |   5.1e+06 | 5000000             |        1 |
| 12 | p5_reading        | 2500                | ADC    | nan       | nan       | 2500                |      nan |
| 13 | +5ERROR           | 0                   | %      |   0       |   3       | 0.7080078125        |        1 |
| 14 | fault_status      | 0                   | nan    | nan       | nan       | 9                   |      nan |
| 15 | +5 Fault          | 0                   | nan    | nan       | nan       | 1                   |        0 |
| 16 | +24 Fault         | 0                   | nan    | nan       | nan       | 0                   |        1 |
| 17 | Error Code 2      | 0                   | nan    | nan       | nan       | 0                   |        1 |
| 18 | DAQ Fault         | 0                   | nan    | nan       | nan       | 1                   |        0 |
| 19 | Error Code 4      | 0                   | nan    | nan       | nan       | 0                   |        1 |
| 20 | MW Fault          | 0                   | nan    | nan       | nan       | 0                   |        1 |
| 21 | Error Code 6      | 0                   | nan    | nan       | nan       | 0                   |        1 |
| 22 | Error Code 7      | 0                   | nan    | nan       | nan       | 0                   |        1 |
| 23 | Visible Fault     | 0                   | nan    | nan       | nan       | 0                   |        1 |
| 24 | Temperature Fault | 0                   | nan    | nan       | nan       | 0                   |        1 |
| 25 | Moisture Fault    | 0                   | nan    | nan       | nan       | 0                   |      nan |
| 26 | Watchdog Fault    | 0                   | nan    | nan       | nan       | 0                   |        1 |
| 27 | Hardware Fault    | 0                   | nan    | nan       | nan       | 0                   |        1 |
| 28 | I2C Fault         | 0                   | nan    | nan       | nan       | 0                   |        1 |
| 29 | Error Code 14     | 0                   | nan    | nan       | nan       | 0                   |        1 |
| 30 | Modbus Fault      | 0                   | nan    | nan       | nan       | 0                   |        1 |


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

