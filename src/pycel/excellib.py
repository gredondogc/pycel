"""Python equivalents of various excel functions."""
from __future__ import division
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from logging import getLogger
import math
from numpy import linalg, vander, zeros
from pahws.pycel.excelutil import (date_from_int, find_corresponding_index, flatten,
                             integer_types, is_leap_year, is_number, list_types,
                             normalize_year, number_types, string_types)

__all__ = ('average', 'count', 'countif', 'countifs', 'date', 'index', 'isNa', 'linest',
           'ln', 'lookup', 'match', 'mid', 'mod', 'npv', 'right', 'roundup', 'sumif',
           'value', 'vlookup', 'yearfrac',
           'xl_log', 'xl_round',
           'xl_eq', 'xl_neq', 'xl_gt', 'xl_gte', 'xl_lt', 'xl_lte',
           'xl_max', 'xl_min', 'xl_sum', 'concatenate', 'full_concatenate', 'text', 'substitute',
           'FUNCTION_MAP')


# A dictionary that maps excel function names onto python equivalents. You should
# only add an entry to this map if the python name is different to the excel name
# (which it may need to be to  prevent conflicts with existing python functions
# with that name, e.g., max).

# So if excel defines a function foobar(), all you have to do is add a function
# called foobar to this module.  You only need to add it to the function map,
# if you want to use a different name in the python code.

# Note: some functions (if, pi, atan2, and, or, array, ...) are already taken care of
# in the FunctionNode code, so adding them here will have no effect.

FUNCTION_MAP = {
    'log': 'xl_log',
    "min": "xl_min",
    "max": "xl_max",
    "sum": "xl_sum",
    "round": "xl_round"
}

logger = getLogger(__name__)

# TODO: add tests for the functions in this module


def value(text):
    """
    Converts a text string that represents a number to a number.

    :param text: string input to be converted to a numerical value
    :return: a numerical value

    """
    try:
        return int(text)
    except ValueError:
        return float(text)


def ln(number):
    """
    Returns the natural logarithm of a number.
    Natural logarithms are based on the constant e (2.71828182845904).

    :param number: Required. The positive real number for which you want the natural logarithm.
    :return: the natural logarithm of the number.

    """

    if isinstance(number, list_types):
        return [math.log(x) for x in flatten(number)]
    else:
        return math.log(number)


def xl_log(number, base=10):
    """
    Returns the logarithm of a number to the base you specify.

    :param number: Required. The positive real number for which you want the logarithm.
    :param base: Optional. The base of the logarithm. If base is omitted, it is assumed to be 10.
    :return: the logarithm of the number.

    """

    if isinstance(number, list_types):
        return [math.log(item, base) for item in flatten(number)]
    else:
        return math.log(number, base)


def xl_max(*args):
    """
    Returns the largest value in a set of values.

    :param args: items to find the largest number in.
    :return: largest number.

    .. remarks::

        * Arguments can either be numbers or names, arrays, or references that contain numbers.
        * Logical values and text representations of numbers that you type directly into the list
        of arguments are counted.
        * If an argument is an array or reference, only numbers in that array or reference are used.
        Empty cells, logical values, or text in the array or reference are ignored.
        * If the arguments contain no numbers, MAX returns 0 (zero).
        * Arguments that are error values or text that cannot be translated into numbers cause errors.
        * If you want to include logical values and text representations of numbers in a reference as
        part of the calculation, use the MAXA function.

    """

    # ignore non-numeric cells
    data = [x for x in flatten(args) if is_number(x)]
    # data = [x for x in flatten(args) if is_number(x, number_types)]

    # however, if no non numeric cells, return zero (is what excel does)
    if len(data) < 1:
        return 0
    else:
        return max(data)


def xl_min(*args):
    """
    Returns the smallest number in a set of values.

    :param args: items to find the smallest number in.
    :return: smallest number.

    .. remarks::
        * Arguments can either be numbers or names, arrays, or references that contain numbers.
        * Logical values and text representations of numbers that you type directly into the list
        of arguments are counted.
        * If an argument is an array or reference, only numbers in that array or reference are used.
        Empty cells, logical values, or text in the array or reference are ignored.
        * If the arguments contain no numbers, MIN returns 0.
        * Arguments that are error values or text that cannot be translated into numbers cause errors.
        * If you want to include logical values and text representations of numbers in a reference as
        part of the calculation, use the MINA function.

    """

    # ignore non numeric cells
    data = [x for x in flatten(args) if isinstance(x, number_types)]

    # however, if no non numeric cells, return zero (is what excel does)
    if len(data) < 1:
        return 0
    else:
        return min(data)


def xl_sum(*args):
    # ignore non numeric cells
    data = [x for x in flatten(args) if isinstance(x, number_types)]

    # however, if no non numeric cells, return zero (is what excel does)
    if len(data) < 1:
        return 0
    else:
        return sum(data)


def sumif(rng, criteria, sum_range=None):
    """
    You use the SUMIF function to sum the values in a range that meet criteria that you specify.
    For example, suppose that in a column that contains numbers, you want to sum only the values
    that are larger than 5. You can use the following formula: =SUMIF(B2:B25,">5")

    :param rng: Required. The range of cells that you want evaluated by criteria. Cells in each
        range must be numbers or names, arrays, or references that contain numbers. Blank and
        text values are ignored. The selected range may contain dates in standard Excel format.
    :param criteria: Required. The criteria in the form of a number, expression, a cell reference,
        text, or a function that defines which cells will be added. For example, criteria can be
        expressed as 32, ">32", B5, "32", "apples", or TODAY().
    :param sum_range:  Optional. The actual cells to add, if you want to add cells other than those
        specified in the range argument. If the sum_range argument is omitted, Excel adds the cells
        that are specified in the range argument (the same cells to which the criteria is applied).
    :return: the total sum of the elements that meet the criteria.

    .. reference::
        https://support.office.com/en-us/article/SUMIF-function-169b8c99-c05c-4483-a712-1697a653039b
    """

    # WARNING:
    # - wildcards not supported
    # - doesn't really follow 2nd remark about sum_range length

    sum_range = sum_range or rng

    if type(rng) != list:
        raise TypeError('%s must be a list' % str(rng))

    if type(sum_range) != list:
        raise TypeError('%s must be a list' % str(sum_range))

    if isinstance(criteria, list_types) and not isinstance(criteria, (string_types, bool)):
        return 0

    indexes = find_corresponding_index(rng, criteria)

    def f(x):
        return sum_range[x] if x < len(sum_range) else 0

    if len(sum_range) == 0:
        return sum(map(lambda x: rng[x], indexes))
    else:
        return sum(map(f, indexes))


def average(*args):
    l = list(flatten(*args))
    return sum(l) / len(l)


def right(text, n):
    # NOTE: hack to deal with NACA section numbers
    if isinstance(text, string_types):
        return text[-n:]
    else:
        # TODO: get rid of the decimal
        return str(int(text))[-n:]


def index(*args):
    array = args[0]
    row = args[1]

    if len(args) == 3:
        col = args[2]
    else:
        col = 1

    if isinstance(array[0], list_types):
        # rectangular array
        # TODO: figure out what is going on here
        array[row-1][col-1]
    elif row == 1 or col == 1:
        return array[row-1] if col == 1 else array[col-1]
    else:
        raise IndexError("index (%s,%s) out of range for %s" %(row,col,array))


def lookup(lookup_value, lookup_vector, result_vector=None):
    # TODO: add non-numeric lookup functionality
    if not isinstance(lookup_value, number_types):
        raise IndexError("Non-numeric lookups (%s) are not supported" % lookup_value)

    # TODO: note, may return the last equal lookup_value

    # index of the last numeric lookup_value
    lastnum = -1
    i = 0
    for i, v in enumerate(lookup_vector):
        if isinstance(v, number_types):
            if v > lookup_value:
                break
            else:
                lastnum = i

    result_vector = result_vector or lookup_vector

    if lastnum < 0:
        raise IndexError("No numeric data found in the lookup range")
    else:
        if i == 0:
            raise IndexError("All values in the lookup range are bigger than %s" % lookup_value)
        else:
            if i >= len(lookup_vector)-1:
                # return the biggest number smaller than lookup_value
                return result_vector[lastnum]
            else:
                return result_vector[i-1]


def vlookup(lookup_value, table_array, col_index_num, range_lookup=True):
    """
    Use VLOOKUP, one of the lookup and reference functions, when you need to find
    things in a table or a range by row. For example, look up a price of an
    automotive part by the part number.

    In its simplest form, the VLOOKUP function says:

    =VLOOKUP(Value you want to look up, range where you want to lookup the value,
        the column number in the range containing the return value,
        Exact Match or Approximate Match – indicated as 0/FALSE or 1/TRUE).

    We'll use the first element to check the array dimension

    :param lookup_value:
    :param table_array:
    :param col_index_num:
    :param range_lookup:
    :return:

    """

    single_column = type(table_array[0]) not in (tuple, list)

    if range_lookup:
        if not isinstance(lookup_value, number_types):
            raise ValueError("Can only do approximate VLOOKUPS with numbers")
        idx = -1
        for idx in range(len(table_array) - 1):
            if single_column:
                if all((table_array[idx] <= lookup_value,
                        table_array[idx + 1] > lookup_value)):
                    return table_array[idx]
            else:
                if all((table_array[idx][0] <= lookup_value,
                        table_array[idx + 1][0] > lookup_value)):
                    return table_array[idx][col_index_num - 1]
        if single_column:
            return table_array[idx+1]
        else:
            return table_array[idx+1][col_index_num - 1]
    else:
        if single_column:
            values = [row for row in table_array]
            return table_array[values.index(lookup_value)]
        else:
            values = [row[0] for row in table_array]
            if lookup_value in values:
                return table_array[values.index(lookup_value)][col_index_num - 1]

    return None


def linest(*args, **kwargs):
    y = args[0]
    x = args[1]

    if len(args) == 3:
        const = args[2]
        if isinstance(const, string_types):
            const = (const.lower() == "true")
    else:
        const = True

    degree = kwargs.get('degree', 1)

    # build the Vandermonde matrix
    a = vander(x, degree + 1)

    if not const:
        # force the intercept to zero
        a[:, -1] = zeros((1, len(x)))

    # perform the fit
    coeffs, residuals, rank, sing_vals = linalg.lstsq(a, y)

    return coeffs


def npv(*args):
    discount_rate = args[0]
    cashflow = args[1]
    return sum([float(x)*(1+discount_rate)**-(i+1) for (i, x) in enumerate(cashflow)])


def match(lookup_value, lookup_array, match_type=1):
    """
    The MATCH function searches for a specified item in a range of cells,
    and then returns the relative position of that item in the range.
    For example, if the range A1:A3 contains the values 5, 25, and 38,
    then the formula =MATCH(25,A1:A3,0) returns the number 2, because 25
    is the second item in the range.

    :param lookup_value: Required. The value that you want to match in lookup_array.
        For example, when you look up someone's number in a telephone book, you are
        using the person's name as the lookup value, but the telephone number is
        the value you want.

        The lookup_value argument can be a value (number, text, or logical value)
        or a cell reference to a number, text, or logical value.

    :param lookup_array: Required. The range of cells being searched.
    :param match_type: Optional. The number -1, 0, or 1. The match_type argument
        specifies how Excel matches lookup_value with values in lookup_array.
        The default value for this argument is 1.

    :return: The index of the first instance of lookup_value in lookup_array.

    ..remarks::
        * MATCH returns the position of the matched value within lookup_array, not
            the value itself. For example, MATCH("b",{"a","b","c"},0) returns 2,
            which is the relative position of "b" within the array {"a","b","c"}.
        * MATCH does not distinguish between uppercase and lowercase letters when matching text values.
        * If MATCH is unsuccessful in finding a match, it returns the #N/A error value.
        * If match_type is 0 and lookup_value is a text string, you can use the
            wildcard characters — the question mark (?) and asterisk (*) — in the
            lookup_value argument. A question mark matches any single character;
            an asterisk matches any sequence of characters. If you want to find an
            actual question mark or asterisk, type a tilde (~) before the character.

    """

    def type_convert(val):
        if isinstance(val, string_types):
            val = val.lower()
        elif isinstance(val, number_types):
            val = float(val)
        return val

    lookup_value = type_convert(lookup_value)

    if match_type == 1:
        # Verify ascending sort
        pos_max = -1
        for i in range((len(lookup_array))):
            current = type_convert(lookup_array[i])
            if i is not len(lookup_array)-1 and current > type_convert(lookup_array[i+1]):
                raise ValueError('for match_type 0, lookup_array must be sorted ascending')
            if current <= lookup_value:
                pos_max = i
        if pos_max == -1:
            raise ValueError('No result in lookup_array for match_type 0')
        # Excel starts at 1
        return pos_max + 1

    elif match_type == 0:
        # No string wildcard
        return [type_convert(x) for x in lookup_array].index(lookup_value) + 1

    elif match_type == -1:
        # Verify descending sort
        pos_min = -1
        for i in range((len(lookup_array))):
            current = type_convert(lookup_array[i])
            if i is not len(lookup_array)-1 and current < type_convert(lookup_array[i+1]):
                raise ValueError('For match_type 0, lookup_array must be sorted descending')
            if current >= lookup_value:
                pos_min = i
        if pos_min == -1:
            raise Exception('no result in lookup_array for match_type 0')
        # Excel starts at 1
        return pos_min + 1


def mod(number, divisor):
    """
    Returns the remainder after number is divided by divisor. The result has the same sign as divisor.

    :param number: Required. The number for which you want to find the remainder.
    :param divisor: Required. The number by which you want to divide number.
    :return: the remainder of dividing the `number` by the `divisor`.

    .. reference::
        https://support.office.com/en-us/article/MOD-function-9b6cd169-b6ee-406a-a97b-edf2a9dc24f3

    .. remarks::
        * If divisor is 0, MOD returns the #DIV/0! error value.
        * The MOD function can be expressed in terms of the INT function:
            * MOD(n, d) = n - d*INT(n/d)

    """

    if not isinstance(number, integer_types):
        raise TypeError("'{}' is not an integer".format(number))
    elif not isinstance(divisor, integer_types):
        raise TypeError("'{}' is not an integer".format(divisor))
    elif divisor == 0:
        raise ZeroDivisionError("Can't return remainder of '{}'".format(number))
    else:
        return number % divisor


def count(*args):
    """

    :param args:
    :return:

    .. reference::
        https://support.office.com/en-us/article/COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c
    """

    l = list(args)

    total = 0

    for arg in l:
        if isinstance(arg, list_types):
            # count inside a list
            total += len(list(filter(lambda x: is_number(x) and type(x) is not bool, arg)))
        # int() is used for text representation of numbers
        elif is_number(arg):
            total += 1

    return total


def concatenate(*args):
    try:
        largs = list(args)

        total_str = ""

        for arg in largs:
            if type(arg) == list:
                for arg2 in arg:
                    total_str = "{0}{1}".format(total_str, arg2 or "")
            else: # int() is used for text representation of numbers
                total_str = "{0}{1}".format(total_str, arg or "")

        return total_str
    except Exception as e:
        return "Error!!!"


def full_concatenate(a1, a2):
    return "{0}{1}".format(a1, a2)


def substitute(input_str, a, b):
    return input_str.replace(a, b)


def text(input_value, format_str):
    try:
        if input_value is None:
            return ""

        if is_number(input_value):
            if format_str == "0,0000":
                if type(input_value) is str:
                    return "{0:0.4f}".format(float(input_value))

                return "{0:0.4f}".format(input_value)
            else:
                return str(input_value)

        return str(input_value)
    except Exception as e:
        return "#value"


def countif(rng, criteria):
    """
    Use COUNTIF, one of the statistical functions, to count the number of cells
    that meet a criterion; for example, to count the number of times a
    particular city appears in a customer list.

    In its simplest form, COUNTIF says:

    =COUNTIF(Where do you want to look?, What do you want to look for?)

    :param rng:
    :param criteria:
    :return:

    .. reference::
        https://support.office.com/en-us/article/COUNTIF-function-e0de10c6-f885-4e71-abb4-1f464816df34

    .. warning:
        wildcards are not supported.

    """

    valid = find_corresponding_index(rng, criteria)

    return len(valid)


def countifs(*args):
    """
    The COUNTIFS function applies criteria to cells across multiple ranges and
    counts the number of times all criteria are met.

    :param args:
    :return:

    .. reference::
        https://support.office.com/en-us/article/COUNTIFS-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842
    """

    arg_list = list(args)
    l = len(arg_list)

    if l % 2 != 0:
        raise Exception('excellib.countifs() must have a pair number of arguments, here %d' % l)

    if l >= 2:
        # find indexes that match first layer of countif
        indexes = find_corresponding_index(args[0], args[1])

        # get only ranges
        remaining_ranges = [elem for i, elem in enumerate(arg_list[2:]) if i % 2 == 0]
        # get only criteria
        remaining_criteria = [elem for i, elem in enumerate(arg_list[2:]) if i % 2 == 1]

        filtered_remaining_ranges = []

        # filter items in remaining_ranges that match valid indexes from first countif layer
        for rng in remaining_ranges:
            filtered_remaining_range = []

            for idx, item in enumerate(rng):
                if idx in indexes:
                    filtered_remaining_range.append(item)

            filtered_remaining_ranges.append(filtered_remaining_range)

        new_tuple = ()

        # rebuild the tuple that will be the argument of next layer
        for idx, rng in enumerate(filtered_remaining_ranges):
            new_tuple += (rng, remaining_criteria[idx])

        # only consider the minimum number across all layer responses
        return min(countifs(*new_tuple), len(indexes))

    else:
        return float('inf')


def roundup(number, num_digits=0):
    """
    Rounds a number up, away from 0 (zero).

    :param number: Required. Any real number that you want rounded up.
    :param num_digits: Required. The number of digits to which you want to round number.
    :return: the rounded number.

    """
    new = round(number, num_digits)
    new += 10 ** -num_digits if number > new else 0
    return round(new, num_digits)


def xl_round(number, num_digits=0):
    """
    The ROUND function rounds a number to a specified number of digits. For example,
    if cell A1 contains 23.7825, and you want to round that value to two decimal places,
    you can use the following formula:

    =ROUND(A1, 2)

    The result of this function is 23.78.

    :param number: Required. The number that you want to round.
    :param num_digits: Required. The number of digits to which you want to round the number argument.
    :return: the rounded number

    .. remarks::

        * If num_digits is greater than 0 (zero), then number is rounded to the specified number of decimal places.
        * If num_digits is 0, the number is rounded to the nearest integer.
        * If num_digits is less than 0, the number is rounded to the left of the decimal point.
        * To always round up (away from zero), use the ROUNDUP function.
        * To always round down (toward zero), use the ROUNDDOWN function.
        * To round a number to a specific multiple (for example, to round to the nearest 0.5), use the MROUND function.

    .. reference::
        https://support.office.com/en-us/article/ROUND-function-c018c5d8-40fb-4053-90b1-b3e7f61a213c

    """

    if not is_number(number):
        raise TypeError("%s is not a number" % str(number))
    if not is_number(num_digits):
        raise TypeError("%s is not a number" % str(num_digits))

    # round to the right side of the point
    if num_digits >= 0:
        return float(Decimal(repr(number)).quantize(Decimal(repr(pow(10, -num_digits))), rounding=ROUND_HALF_UP))
        # see https://docs.python.org/2/library/functions.html#round
        # and https://gist.github.com/ejamesc/cedc886c5f36e2d075c5
    else:
        return float(round(number, num_digits))


def mid(text, start_num, num_chars):
    """

    :param text:
    :param start_num:
    :param num_chars:
    :return:

    .. reference::
        https://support.office.com/en-us/article/MID-MIDB-functions-d5f9e25c-d7d6-472e-b568-4ecb12433028
    """

    text = str(text)

    if type(start_num) != int:
        raise TypeError("%s is not an integer" % str(start_num))
    if type(num_chars) != int:
        raise TypeError("%s is not an integer" % str(num_chars))

    if start_num < 1:
        raise ValueError("%s is < 1" % str(start_num))
    if num_chars < 0:
        raise ValueError("%s is < 0" % str(num_chars))

    return text[start_num:num_chars]


def date(year, month, day):
    """
    The DATE function returns the sequential serial number that represents a particular date.

    :param year: Required. The value of the year argument can include one to four digits.
        Excel interprets the year argument according to the date system your computer is
        using. By default, Microsoft Excel for Windows uses the 1900 date system, which
        means the first date is January 1, 1900.
    :param month: Required. A positive or negative integer representing the month of the
        year from 1 to 12 (January to December).
    :param day: Required. A positive or negative integer representing the day of the month from 1 to 31.
    :return: serial number that represents the particular date.

    .. reference::
        https://support.office.com/en-us/article/DATE-function-e36c0c8c-4104-49da-ab83-82328b832349

    """

    if type(year) != int:
        raise TypeError("%s is not an integer" % str(year))

    if type(month) != int:
        raise TypeError("%s is not an integer" % str(month))

    if type(day) != int:
        raise TypeError("%s is not an integer" % str(day))

    if year < 0 or year > 9999:
        raise ValueError("Year must be between 1 and 9999, instead %s" % str(year))

    if year < 1900:
        year += 1900

    # taking into account negative month and day values
    year, month, day = normalize_year(year, month, day)

    date_0 = datetime(1900, 1, 1)
    date = datetime(year, month, day)

    result = (date - date_0).days + 2

    if result <= 0:
        raise ArithmeticError("Date result is negative")
    else:
        return result


def yearfrac(start_date, end_date, basis=0):
    """
    Calculates the fraction of the year represented by the number of whole days between
    two dates (the start_date and the end_date). Use the YEARFRAC worksheet function to
    identify the proportion of a whole year's benefits or obligations to assign to a
    specific term.

    :param start_date: Required. A date that represents the start date.
    :param end_date: Required. A date that represents the end date.
    :param basis: Optional. The type of day count basis to use.
    :return: Fraction of the year between the two dates.

    .. reference::
        https://support.office.com/en-us/article/YEARFRAC-function-3844141e-c76d-4143-82b6-208454ddc6a8

    .. remarks::
        * Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations.
            By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448
            because it is 39,448 days after January 1, 1900.
        * All arguments are truncated to integers.
        * If start_date or end_date are not valid dates, YEARFRAC returns the #VALUE! error value.
        * If basis < 0 or if basis > 4, YEARFRAC returns the #NUM! error value.


    """

    def actual_nb_days_ISDA(start, end): # needed to separate days_in_leap_year from days_not_leap_year
        y1, m1, d1 = start
        y2, m2, d2 = end

        days_in_leap_year = 0
        days_not_in_leap_year = 0

        year_range = range(y1, y2 + 1)

        for y in year_range:

            if y == y1 and y == y2:
                nb_days = date(y2, m2, d2) - date(y1, m1, d1)
            elif y == y1:
                nb_days = date(y1 + 1, 1, 1) - date(y1, m1, d1)
            elif y == y2:
                nb_days = date(y2, m2, d2) - date(y2, 1, 1)
            else:
                nb_days = 366 if is_leap_year(y) else 365

            if is_leap_year(y):
                days_in_leap_year += nb_days
            else:
                days_not_in_leap_year += nb_days

        return days_not_in_leap_year, days_in_leap_year

    def actual_nb_days_AFB_alter(start, end):
        """
        .. note::
            Converted from: http://svn.finmath.net/finmath%20lib/trunk/src/main/java/net/finmath/time/daycount/DayCountConvention_ACT_ACT_YEARFRAC.java
        """
        y1, m1, d1 = start
        y2, m2, d2 = end

        delta = date(*end) - date(*start)

        if delta <= 365:
            if is_leap_year(y1) and is_leap_year(y2):
                denom = 366
            elif is_leap_year(y1) and date(y1, m1, d1) <= date(y1, 2, 29):
                denom = 366
            elif is_leap_year(y2) and date(y2, m2, d2) >= date(y2, 2, 29):
                denom = 366
            else:
                denom = 365
        else:
            year_range = range(y1, y2 + 1)
            nb = 0

            for y in year_range:
                nb += 366 if is_leap_year(y) else 365

            denom = nb / len(year_range)

        return delta / denom

    if not is_number(start_date):
        raise TypeError("start_date %s must be a number" % str(start_date))
    if not is_number(end_date):
        raise TypeError("end_date %s must be number" % str(end_date))
    if start_date < 0:
        raise ValueError("start_date %s must be positive" % str(start_date))
    if end_date < 0:
        raise ValueError("end_date %s must be positive" % str(end_date))

    # switch dates if start_date > end_date
    if start_date > end_date:
        temp = end_date
        end_date = start_date
        start_date = temp

    y1, m1, d1 = date_from_int(start_date)
    y2, m2, d2 = date_from_int(end_date)

    # US 30/360
    if basis == 0:
        d2 = 30 if d2 == 31 and (d1 == 31 or d1 == 30) else min(d2, 31)
        d1 = 30 if d1 == 31 else d1

        count = 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)
        result = count / 360

    # Actual/actual
    elif basis == 1:
        result = actual_nb_days_AFB_alter((y1, m1, d1), (y2, m2, d2))

    # Actual/360
    elif basis == 2:
        result = (end_date - start_date) / 360

    # Actual/365
    elif basis == 3:
        result = (end_date - start_date) / 365

    # Eurobond 30/360
    elif basis == 4:
        d2 = 30 if d2 == 31 else d2
        d1 = 30 if d1 == 31 else d1

        result = (360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1)) / 360

    else:
        raise ValueError("%d must be 0, 1, 2, 3 or 4" % basis)

    return result


def alpha_value(text):
    return sum(ord(char) for char in text)


def xl_eq(value1, value2):
    if all(isinstance(value, number_types) for value in (value1, value2)):
        return float(value1) == float(value2)
    # in Excel capitalization does not matter
    elif all(isinstance(val, string_types) for val in (value1, value2)):
        return value1.lower() == value2.lower()
    # in Excel "" equals empty cell (most of the times)
    elif None in (value1, value2) and '' in (value1, value2):
        return True
    else:
        try:
            return value1 == value2
        except TypeError:
            logger.debug("Could not compare '{}' and '{}'".format(value1, value2))
            return False


def xl_neq(value1, value2):
    return not xl_eq(value1, value2)


def xl_gt(value1, value2):
    if all(isinstance(value, number_types) for value in (value1, value2)):
        return float(value1) > float(value2)
    # in Excel capitalization does not matter
    elif all(isinstance(val, string_types) for val in (value1, value2)):
        ordered = sorted(value1.lower(), value2.lower())
        return (not(value1.lower() == value2.lower()) and
                ordered[-1] == value1.lower())
    elif (isinstance(value1, string_types) and
          isinstance(value2, number_types)):
        return True
    elif (isinstance(value1, number_types) and
          isinstance(value2, string_types)):
        return False
    else:
        try:
            return value1 > value2
        except TypeError:
            logger.debug("Could not compare '{}' and '{}'".format(value1, value2))
            return False


def xl_gte(value1, value2):
    return xl_eq(value1, value2) or xl_gt(value1, value2)


def xl_lt(value1, value2):
    return not xl_gte(value1, value2)


def xl_lte(value1, value2):
    return not xl_gt(value1, value2)


def isNa(val):
    # This function might need more solid testing
    # TODO: maybe use proper float('nan') and test against that
    try:
        eval(val)
        return False
    except (ValueError, TypeError):
        return True
    except Exception as exc:
        logger.debug("'{}' evaluated as #NA but not ValueError nor TypeError, {}".format(val, exc))
        return True
