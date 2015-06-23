"""Usage: eqcell.py FILE [SHEET] [--close-after]

Fill SHEET in Excel FILE with formulas in START-END cells range.

Arguments:
  FILE          Excel input file name (e.g. cells.xls) or path (e.g. c:\\work\\cells.xls).
  SHEET         Optional Excel sheet name or sheet number (default value = 1).

Flags:
  close-after   Close Excel file open after script execution.
"""

from collections import deque
from xlwings import Workbook, Range, Sheet
from config import TIME_INDEX_VARIABLES, START_RANGE_MARKER, END_RANGE_MARKER, FIRST_SHEET
from sympy import var
import os
from docopt import docopt
import itertools

global current_sheet
current_sheet = FIRST_SHEET

def set_sheet(sh):
    """
    Set the current_sheet number or current_sheet name

    input
    -----
    sh: current_sheet number (int) or name (string)
    """
    global current_sheet
    current_sheet = sh

def get_sheet():
    """
    Get the current_sheet number or current_sheet name
    """
    global current_sheet
    return current_sheet

def solver(workfile, sheet=FIRST_SHEET, savefile=None, close_after=False, is_contiguous=False):
    """
    Processes a range on Excel sheet - applies formulas defined as strings to corresponding Excel formulas in cells.
    Saves the workbook with formulas applied (overwrites existing file by default)

    input
    -----
    workfile:  input excel filepath to be processed
    sheet:     sheet number or name to be processed (default 1)
    savefile:  output excel filepath to be saved.
    """

    wb = Workbook(workfile)
    # wb = Workbook(get_source_file_path(arguments))

    # work on given sheet
    set_sheet(sheet)

    start_cell, end_cell = find_start_end_indices(wb, is_contiguous)
    # Parse the column under START label
    variables, formulas, comments = parse_variable_values(wb, start_cell, end_cell)
    # checks if 'is_forecast' variable name is present on sheet
    require_variable(variables, 'is_forecast')
    # parse formulas and obtain Sympy expressions
    parsed_formulas = parse_formulas(formulas, variables)

    apply_formulas_on_sheet(wb, variables, parsed_formulas, start_cell)
    if savefile is None:
        savefile = workfile
    save_workbook(wb, savefile)

    # close file
    Sheet(FIRST_SHEET).activate  # are this missing a funcion call? Should always activate the first sheet?
    if close_after is True:
        close_workbook(wb)

def move():
    """
    This function is used to read the excel cells in a sheet in zigzag manner. The zig zag pattern in which the
    indices are returned is similar to encoding of JPEG.

    This generator is used to find the 'START' cell in the sheet.

    Yields:
    (x, y): x,y coordinate, where x is the row, y is column equivalent in excel
    """
    x = 1
    y = 1
    directions = deque([(1, -1), (1, 0), (-1, 1), (0, 1)])
    while True:    # Warning: possibility of a infinite loop if - 1) END is not present in sheet or 2) END appears before START
        yield x, y
        if x == 1 or y == 1:
            directions.rotate(-1)

        x += directions[0][0]
        y += directions[0][1]

def find_start_end_indices_contiguous(workbook):
    """
    Given the workbook reference, find the indices of START, END cells on the sheet

    Returns:
    start_cell: A dictionary of cell coordinates containing value 'START'
    end_cell:   A dictionary of cell coordinates containing value 'END'
    """
    workbook.set_current()    # Sets the workbook as the current working workbook
    zigzag_gen = move()       # Used to read cells in zigzag pattern, faster than going row-wise or column-wise
    start_index = next(coords for coords in zigzag_gen if Range(get_sheet(), coords).value == START_RANGE_MARKER)    # (x, y) coord of 'START'
    end_x = start_index[0] + len(Range(get_sheet(), start_index).vertical.value)
    end_y = start_index[1]
    # risks infinite loops

    while True:    # Please make sure that the cell in column of START and row of END is empty
        if Range(get_sheet(), (end_x, end_y)).value == END_RANGE_MARKER:
            break
        else:
            end_y += 1

    end_index = (end_x, end_y)    # (x, y) coords of 'END'
    keys = ('row', 'col')
    start_cell = dict(zip(keys, start_index))
    end_cell = dict(zip(keys, end_index))
    return start_cell, end_cell

def find_start_end_indices_not_contiguous(workbook):
    """
    Given the workbook reference, find the indices of START, END cells on the sheet

    Returns:
    start_cell: A tuple pair of excel sheet coordinates of cell containing value 'START'
    end_cell: A tuple pair of excel sheet coordinates of cell containing value 'END'
    """
    workbook.set_current()    # sets the workbook as the current working workbook
    zigzag_gen = move()       # used to read cells in zigzag pattern, faster than going row-wise or column-wise
    start_index = next(coords for coords in zigzag_gen if Range(get_sheet(), coords).value == 'START')    # (x, y) coord of 'START'
    end_index = next(coords for coords in zigzag_gen if Range(get_sheet(), coords).value == 'END')        # (x, y) coord of 'END''
    keys = ('row', 'col')

    start_cell = dict(zip(keys, start_index))
    end_cell = dict(zip(keys, end_index))
    return start_cell, end_cell

def find_start_end_indices(workbook, is_contiguous=True):
    # 2015-05-12 12:41 PM
    # new behaviour - find START, then try contiguous first to find END,
    #                 if not successful, then try not_contiguousto find END
    if is_contiguous is True:
        return find_start_end_indices_contiguous(workbook)
    else:
        return find_start_end_indices_not_contiguous(workbook)

def parse_variable_values(workbook, start_cell, end_cell):
    """
    Given the workbook reference and START-END index pair, this function parses the values in the variable row
    and saves it as a list of the same name.

    input
    -----
    workbook:   Workbook xlwings object
    start_cell: Start cell dictionary
    end_cell:   End cell dictionary

    returns:    lists of variables, formulas, comments
    """
    workbook.set_current()    # sets the workbook as the current working workbook
    variables = dict()
    formulas = dict()
    comments = dict()
    start = (start_cell['row'], start_cell['col'])
    end = (end_cell['row'],   start_cell['col'])

    start_column = Range(get_sheet(), start, end).value
    # [1:] excludes 'START' element
    start_column = start_column[1:]

    for relative_index, element in enumerate(start_column):
        current_index = start_cell['row'] + relative_index + 1
        if element:    # if non-empty
            if not isinstance(element, str):
                raise ValueError("The column below START can contain only strings")

            # print(element)
            element = element.strip()

            if '=' in element:
                formulas[element] = current_index
            elif '#' == element[0]:
                comments[element] = current_index
            else:
                variables[element] = current_index

    return variables, formulas, comments

def require_variable(variables, var='is_forecast'):
    """
    Checks if variable string (default: `is_forcast`) is in the sheet variables dict, else raises error

    input
    -----
    variables: A dict of variables from excel sheet
    var:       A variable name string, to be checked if exists in variables.
    """
    if var not in variables.keys():
        raise ValueError('is_forecast is a mandatory value under START cell in excel sheet')

def evaluate_variable(x):
        try:
            x = eval(x)     # converting the formula into sympy expressions
        except NameError:
            raise NameError('Undefined variables in formulas, check excel sheet')
        return x

def parse_formulas(formulas, variables):
    """
    Takes formulas as a dict of strings and returns a dict
    where dependent (left-hand side) variable and (right-hand side) formula
    are separated and converted to sympy expressions.

    input variable example:
    formulas = {'a(t)=a(t-1)*a_rate(t)': 6, 'b(t)=b_share(t)*a(t)': 11}

    output example:
    formulas_dict = {5: {'dependent_var': a(t), 'formula': a(t-1)*a_rate(t)},
                     9: {'dependent_var': b(t), 'formula': b(t-1)+2}}

    5, 6, 9, 11 are the row indices in the sheet. Row indices in formulas_dict changed to rows with variables.
    These rows contain data and are used to fill in formulas in forecast period.
    a(t), b(t-1)+2, ... are sympy expressions.
    """
    varirable_list = list(variables.keys()) + TIME_INDEX_VARIABLES

    # declares sympy variables
    var(' '.join(varirable_list))
    parsed_formulas = dict()

    for formula_string in formulas.keys():
        # removing white spaces
        formula_string = formula_string.strip()
        dependent_var, formula = formula_string.split('=')
        dependent_var = evaluate_variable(dependent_var)
        formula = evaluate_variable(formula)
        # finding the row where the formula will be applied - may be a function
        row_index = variables[str(dependent_var.func)]
        parsed_formulas[row_index] = {'dependent_var': dependent_var, 'formula': formula}

    return parsed_formulas

def simplify_expression(expression, time_period, variables, depth=0):
    # get_variable_to_cell_segments
    """
    A recursive function which breaks a Sympy expression into segments, where each segment points to one cell on the
    excel sheet upon substitution of time index variable (t). Returns a dictionary of such segments and the computed
    cells.

    input
    -----
    expression:       Sympy expression, e.g: a(t - 1)*a_rate(t)
    time_period:      A value to be time_periodtituted for the time index, t.
    variables:        A list of all variables extracted from excel sheet.
    depth:            Depth of recursion, used internally

    returns:          A dict with a segment as key and computed excel cell index as value, e.g: {a(t - 1): (5, 4), a_rate(t): (4, 5)}
    """
    result = {}
    variable = expression.func        # get the function from sympy expression, e.g for expression = f(t), `f` is the function
    if variable.is_Function:
        # for simple expressions like f(t), variable=f and variable.is_Function = True,
        # for more complex expressions, variable would be another expression, hence would have to be broken down recursively.
        cell_row = variables[str(variable)]            # get the row index from variable name
        x = list(expression.args[0].free_symbols)[0]   # get the independent var, mostly `t` from the argument in expression
        cell_col = int(expression.args[0].subs(x, time_period))
        result[expression] = (cell_row, cell_col)
    else:
        if depth > 5:
            raise ValueError("Expression is too complicated: " + expression)

        depth += 1
        for segment in expression.args:
            result.update(simplify_expression(segment, time_period, variables, depth))

    return result

def get_excel_formula_as_string(right_side_expression, time_period, variables):
    """
    Using the right-hand side of a math expression (e.g. a(t)=a(t-1)*a_rate(t)), converted to sympy
    expression, and substituting the time index variable (t) in it, the function finds the Excel formula
    corresponding to the right-hand side expression.

    input
    -----
    right_side_expression:         sympy expression, e.g. a(t-1)*a_rate(t)
    time_period:        value of time index variable (t) for time_periodtitution
    output:
    formula_string:     a string of excel formula, e.g. '=A20*B21'
    """
    right_dict = simplify_expression(right_side_expression, time_period, variables)
    for right_key, right_coords in right_dict.items():
        excel_index = str(Range(get_sheet(), tuple(right_coords)).get_address(False, False))
        right_side_expression = right_side_expression.subs(right_key, excel_index)
    formula_str = '=' + str(right_side_expression)
    return formula_str

def _get_formula(parsed_formulas, row, col):
    """
    Returns the formula for a given row and column.
    """
    try:
        formula_dict = parsed_formulas[row]
    except KeyError:
        formula_dict = dict()
        if Range(get_sheet(), (row, col)).value is None:    # if cell is empty and formula for it not found
            print("Warning: Formula for empty cell not found, incomplete sheet, cell: " +
                  Range(get_sheet(), (row, col)).get_address(False, False))

    return formula_dict

def apply_formulas_on_sheet(workbook, variables, parsed_formulas, start_cell):
    """
    Takes each cell in the sheet inside the rectangle formed by Start_cell and End_cell
    checks 1) if the cell is in a row with a variable as first element
           2) if the cell is in a column with `is_forecast=1`

    If all above conditions are met, then apply a fitting formula as obtained from find_formulas()

    Apply's the solution on the workbook cells. Raises error if any problem arises.

    input
    -----
    workbook:   Workbook xlwings object
    variables: A dict of variables from excel sheet
    parsed_formulas: A dict of formulas with key as row_index and value as dict of left-side and right-side sympy expressions
    start_cell: Start cell dictionary

    """
    workbook.set_current()    # sets the workbook as the current working workbook
    forecast_row = Range(get_sheet(), (variables['is_forecast'], start_cell['col'] + 1)).horizontal.value
    col_indices = [start_cell['col'] + 1 + index for index, el in enumerate(forecast_row) if el == 1]    # checks if is_forecast value in this col is = 1 and notes down col index
    row_indices = list(variables.values())
    row_indices.remove(variables['is_forecast'])

    for col, row in itertools.product(col_indices, row_indices):

        formula_dict = _get_formula(parsed_formulas, row, col)

        if formula_dict:
            dependent_variable_with_time_index = formula_dict['dependent_var']     # get expression for dependent variable, e.g. a(t)
            # dependent_variable_locations - values like {b(t): (8, 5)}
            dependent_variable_locations = simplify_expression(dependent_variable_with_time_index, col, variables)
            dv_key, dv_coords = dependent_variable_locations.popitem()

            # 2015-05-12 03:09 PM
            # --- Need to make this check elsewhere
            if dependent_variable_locations:
                raise ValueError('cannot have more than one dependent variable on left side of equation')
            # --- end

            # find excel type formula string
            right_side_expression = formula_dict['formula']
            formula_str = get_excel_formula_as_string(right_side_expression, col, variables)
            Range(get_sheet(), dv_coords).formula = formula_str                # Apply formula on excel cell

def save_workbook(workbook, savepath=None):
    """
    Saves the workbook in given path or overwrites existing file.
    """
    if savepath is None:
        workbook.save()                          # save over the same workbook (overwrite)
    else:
        savepath = os.path.normcase(savepath)    # makes '/' into '\' for windows compatibility
        workbook.save(savepath)                  # SaveAs with given path

def close_workbook(workbook):
    """
    Closes the workbook in given path
    """
    workbook.close()

def check_file_path(path):
    """
    _____
    """
    if not os.path.isfile(path):
        raise ValueError(path + 'does not exist')
    else:
        # print(path)
        pass

def get_source_file_path(arguments):
    """
    Returns absolute path to Excel file
    """
    if os.path.isabs(arguments['FILE']):
        source_file_path = arguments['FILE']
    else:
        project_directory = os.path.dirname(os.path.abspath(__file__))
        source_file_path = os.path.join(project_directory, arguments['FILE'])

    source_file_path = os.path.normcase(source_file_path)
    check_file_path(source_file_path)
    return source_file_path

def get_output_file_path(arguments):
    """
    DEPRECIATED
    """
    if arguments['--output'] is False:
        output_file_path = get_source_file_path(arguments)
    else:
        output_file_path = arguments['OUTFILE']
    check_file_path(output_file_path)
    return output_file_path

def get_sheet_argument(arguments):
    # cannot process string names like '1', '2' - will convert to intergers
    # need to add - if arguments['SHEET'] starts and ends with ' or " consider it a string name.
    # this will allow calls like: eqcell.py test.xls '1'
    if arguments['SHEET'] is not None:
        try:
            sheet = int(arguments['SHEET'])
        except ValueError:
            sheet = arguments['SHEET']
    else:
        sheet = 1
    return sheet

if __name__ == '__main__':
    arguments = docopt(__doc__)
    source_file_path = get_source_file_path(arguments)
    output_file_path = source_file_path
    sheet_choice = get_sheet_argument(arguments)
    solver(source_file_path, sheet=sheet_choice, savefile=output_file_path,
           close_after=arguments['--close-after'], is_contiguous=True)

# TEST:
# python eqcell.py test/test.xls Test1
# python eqcell.py test/test.xls Test3

# TODO:
# try is_contigious followed by not_contigious
# faster search for END on page + timing of this search
# SHEET = '1' or SHEET = "1"
# refactroing mentioned in code above
# py2exe - make executable
# different formulas for is_forecast=0 and is_forecast=1
# ...
