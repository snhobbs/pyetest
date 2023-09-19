import subprocess
import tempfile
import shutil
import logging
import os
import openpyxl


def reload_spreadsheet(workbook):
    '''
    Uses libreoffice to fill in the cached values.
    Will not update the file if the exported filename is the same.
    Filename can be any type of spreadsheet
    '''
    dir_a = tempfile.mkdtemp()
    dir_b = tempfile.mkdtemp()
    fname_intermediate = os.path.join(dir_a, "pyetest.xlsx")
    fname_final = os.path.join(dir_b, "pyetest.xlsx")

    workbook.save(fname_intermediate)
    subprocess.call(["libreoffice", "--convert-to", "xlsx", fname_intermediate, "--outdir", dir_b])
    #logging.getLogger().info("Generated %s", outname)
    return openpyxl.open(fname_final, data_only=True)


def get_operator_dict():
    '''
    Returns a lookup table to use symbols that correspond to a comparison function
    '''
    def in_range(value, line) -> bool:
        return value >= line["min"] and value <= line["max"]

    def equals(value, line) -> bool:
        return value == line["expected value"]

    def ge(value, line) -> bool:
        return value >= line["min"]

    def le(value, line) -> bool:
        return value <= line["min"]

    def not_zero(value) -> bool:
        return value !=0 and value is not None

    operations = {"<=>": in_range, "==": equals, ">=": ge, "<=": le, "!=0": not_zero}
    return operations


def get_column_number_dict(workbook):
    '''
    Generates a lookup table from column name to column number
    :param Workbook workbook: template object
    '''
    column_names = {}
    for column, in workbook.active.iter_cols(min_row=1, max_row=1, values_only=False):
        column_names[column.value] = column.column - 1#_letter
    return column_names


def compile_template(workbook, data):
    '''
    Fills a template with values from a seperate dataframe
    :param Workbook workbook: Template to be filled, field names need to be correct
    :param [dict, DataFrame] data: Table of names and values matching the names in the template
    '''
    starting_row = 1
    column_names = get_column_number_dict(workbook)
    for row in workbook.active.iter_rows(min_row=starting_row+1, values_only=False):
        if row[column_names["type"]].value == "measured":
            column = column_names["value"]+1 # counts from 1
            coordinate = workbook.active.cell(row[0].row, column=column).coordinate
            name = row[column_names["name"]].value
            value = data[data["name"] == name].iloc[0]["value"]
            workbook.active[coordinate] = value
    return workbook


def run_text_matrix(data, test_matrix):
    '''
    Runs through a dataframe and runs all tests. Returns a dataframe expanded with
    the test results. All names in the test_matrix that have an operation defined need
    to exist in data.
    :param DataFrame data: name, value, unit [optional], type
    :param DataFrame test_matrix: name, expected value, unit [optional], min, max, operation
    :return DataFrame: test_matrix with the added value and pass fields
    '''
    values = []
    passes = []

    operations = get_operator_dict()
    for i, test_line in test_matrix.iterrows():
        line = data[data["name"] == test_line["name"]].iloc[0]
        value = line["value"]
        values.append(value)
        passes.append(operations[test_line["operation"].strip("\"'")](value, test_line))

    test_matrix["value"] = values
    test_matrix["passes"] = passes
    return test_matrix

