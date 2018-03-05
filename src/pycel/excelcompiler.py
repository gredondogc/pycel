from __future__ import print_function, division, unicode_literals
from collections import defaultdict, OrderedDict, deque
from functools import reduce
from logging import getLogger
from math import (acos, acosh, asin, asinh, atan, atan2, atanh, cos, cosh, exp, sin, sinh, sqrt, tan, tanh)
from networkx import (DiGraph, descendants,
                      draw_networkx_nodes, draw_networkx_edges, draw_networkx_labels, spring_layout)
from networkx.drawing.nx_pydot import write_dot
from networkx.readwrite.gexf import write_gexf
import pickle
from six import itervalues

try:
    import matplotlib.pyplot as plt
except ImportError:
    plt = None

# TODO: make these imports more specific
from .excellib import *
from .excelutil import *
from .excelwrapper import wrapper_implementation
from .tokenizer import ExcelParser, f_token, shunting_yard

__version__ = filter(str.isdigit, "$Revision: 2524 $")
__date__ = filter(str.isdigit, "$Date: 2011-09-06 17:05:00 +0100 (Tue, 06 Sep 2011) $")
__author__ = filter(str.isdigit, "$Author: dg2d09 $")

logger = getLogger(__name__)


class Spreadsheet(object):
    def __init__(self, graph, cellmap, filename=''):
        super(Spreadsheet, self).__init__()
        self.graph = graph
        self.cellmap = cellmap
        self.filename = filename
        self.params = None

        self.dependent = tuple(set(edge[1].address() for edge in self.graph.edges()))
        self.independent = tuple([name for name, cell in self.cellmap.items()
                                  if name not in self.dependent and cell.value is not None])

        data = defaultdict(set)
        for edge in self.graph.edges():
            data[edge[1].address()].add(edge[0].address())
        self.sorted_cells = tuple(self.topological_sort(data))

    @staticmethod
    def load_from_file(filename):
        with open(filename, 'rb') as f:
            obj = pickle.load(f)
        return obj

    @staticmethod
    def topological_sort(data):
        r"""
        Sort elements based on their dependencies.

        Dependencies are expressed as a dictionary whose keys are items
        and whose values are a set of dependent items. Output is a list of
        sets in topological order. The first set consists of items with no
        dependants, each subsequent set consists of items that depend upon
        items in the preceeding sets.

        >>> print('\\n'.join(repr(sorted(x)) for x in topological_sort({
        ...     2: {11},
        ...     9: {11, 8},
        ...     10: {11, 3},
        ...     11: {7, 5},
        ...     8: {7, 3},
        ...     })))
        [3, 5, 7]
        [8, 11]
        [2, 9, 10]

        @author: Pady McCarthy
        @date: 2010, Sep. 28
        @modified by: Sam Denton
        @date: 2012, Oct. 29
        @modified by: Santiago Balestrini
        @date: 2013, Jul. 5

        """
        # Ignore self dependencies.
        for k, v in data.items():
            v.discard(k)
        # Find all items that don't depend on anything
        extra_items_in_deps = reduce(set.union, data.values()) - set(data.keys())
        # Add empty dependencies where needed
        data.update({item: set() for item in extra_items_in_deps})
        while True:
            ordered = set(item for item, dep in data.items() if not dep)
            if not ordered:
                break
            yield ordered
            data = {item: (dep - ordered) for item, dep in data.items()
                    if item not in ordered}

        msg = "Cyclic dependencies exist among these items:\n\t{}"
        try:
            assert not data, msg.format('\n\t'.join(repr(x) for x in data.items()))
        except AssertionError:
            logger.debug(msg.format('\n\t'.join(repr(x) for x in data.items())))

    def find_divergence(self, wb):
        invalid_cells = []
        level = 0
        for level, cell_names in enumerate(self.sorted_cells):
            for cell_name in cell_names:
                xl_val = wb.excel.get_range(cell_name).value2
                if not almost_equal(xl_val, self.cellmap[cell_name].value):
                    invalid_cells.append((cell_name, xl_val, self.cellmap[cell_name].value))
            if invalid_cells:
                break
        if invalid_cells:
            msg = "Divergence in {:d} calculation steps after constant declaration".format(level)
        else:
            msg = "No divergence between Excel and Python models"
        return invalid_cells, msg

    def save_to_file(self, filename, protocol=2):
        with open(filename, 'wb') as fp:
            pickle.dump(self, fp, protocol=protocol)

    def export_to_dot(self, filename):
        write_dot(self.graph, filename)

    def export_to_gexf(self, filename):
        write_gexf(self.graph, filename)

    def plot_graph(self, iterations=2000, show_arrows=True):
        if plt:
            pos = spring_layout(self.graph, iterations=iterations)
            draw_networkx_nodes(self.graph, pos)
            draw_networkx_edges(self.graph, pos, arrows=show_arrows)
            draw_networkx_labels(self.graph, pos)
            plt.show()
        else:
            logger.error("Could not draw graph because matlplotlib could not be found!")

    def set_value(self, cell, val, is_addr=True):
        if is_addr:
            cell = self.cellmap[cell]

        if cell.value != val:
            # reset the node + its dependencies
            self.reset(cell)
            # set the value
            cell.value = val

    def reset(self, cell):
        if cell.value is None:
            return
        cell.value = None
        for descendant in descendants(self.graph, cell):
            if isinstance(descendant, CellRange) or descendant.formula:
                descendant.value = None

    def print_value_tree(self, addr, indent=2):
        cell = self.cellmap[addr]
        print("%s %s = %s" % (" "*indent, addr, cell.value))
        for c in self.graph.predecessors_iter(cell):
            self.print_value_tree(c.address(), indent+1)

    def get_formula_tree(self, addr, cell_list):
        cell = self.cellmap.get(addr, None)
        if cell is None:
            print("Address not Found {0}".format(addr))
            return

        cell_list.append(cell)
        for c in self.graph.predecessors_iter(cell):
            self.get_formula_tree(c.address(), cell_list)

    def recalculate(self):
        for cell in reversed(self.cellmap.values()):
            if isinstance(cell, CellRange):
                self.evaluate_range(cell, is_addr=False)
            else:
                self.evaluate(cell, is_addr=False)

    def evaluate_range(self, rng, is_addr=True):
        if is_addr:
            rng = self.cellmap[rng]

        # it's important that [] gets treated as false here
        if rng.value:
            return rng.value

        cells, nrows, ncols = rng.celladdrs, rng.nrows, rng.ncols

        if nrows == 1 or ncols == 1:
            data = [self.evaluate(cell) for cell in cells]
        else:
            data = [[self.evaluate(cell) for cell in cells[j]] for j in range(len(cells))]

        rng.value = data
        return data

    def evaluate(self, cell, is_addr=True):
        if is_addr:
            cell = self.cellmap[cell]

        # no formula, fixed value
        if not cell.formula or cell.value is not None:
            # Returning constant or cached value for cell
            return cell.value

        def eval_cell(address):
            """
            Recalculate the values based on the formula.

            :param address:
            :return:

            .. note::
                The compiled expression calls this function.
            """

            result = self.evaluate(address)
            # if result is None:  # TODO isto não faz sentido. Se fizeres xl_min com isto estragava tudo
            #     return 0.
            return result

        def eval_range(rng):
            """
            Recalculate the values of the range based on the formulas of the cells.

            :param rng:
            :return:

            .. note::
                The compiled expression calls this function.
            """
            # TODO: add a similar check to replace Nones with 0.0
            return self.evaluate_range(rng)

        try:
            logger.debug("Evaluating: %s, %s" % (cell.address(), cell.python_expression))
            cell.value = eval(cell.compiled_expression)
        except TypeError as exc:
            if 'if ' in cell.python_expression:
                logger.error("IF statement TypeError failure!")
            cell.value = 0.
            logger.error("TypeError for '{}' ({})".format(cell.address(), cell.python_expression))
            logger.error("{}".format(exc))
        except Exception as exc:
            if exc.args[0].startswith("Problem evaluating"):
                raise exc
            else:
                raise Exception("Problem evaluating: %s for %s, %s" % (exc, cell.address(),
                                                                       cell.python_expression))

        # Cells in Excel that are empty have an effective numerical value of zero
        if cell.value is None and cell.formula:
            cell.value = 0.

        if cell.value is None:
            return 0.

        return cell.value


class ASTNode(object):
    """A generic node in the AST"""

    def __init__(self,token):
        super(ASTNode, self).__init__()
        self.token = token

    def __str__(self):
        return self.token.tvalue

    def __getattr__(self,name):
        return getattr(self.token, name)

    def children(self,ast):
        args = ast.predecessors(self)
        args = sorted(args, key=lambda x: ast.node[x]['pos'])
        return args

    def parent(self,ast):
        args = ast.successors(self)
        return args[0] if args else None

    def emit(self,ast, context=None):
        """Emit code"""
        return self.token.tvalue


class OperatorNode(ASTNode):
    def __init__(self, *args):
        super(OperatorNode,self).__init__(*args)

        # convert the operator to python equivalents
        self.opmap = {
            "^": "**",
            "&": "full_concatenate",
            "": "+",  # union
            "=": "xl_eq",
            "<>": "xl_neq",
            ">": "xl_gt",
            "<": "xl_lt",
            ">=": "xl_gte",
            "<=": "xl_lte"
        }

    def emit(self, ast, context=None):
        xop = self.tvalue

        # Get the arguments
        args = self.children(ast)

        op = self.opmap.get(xop, xop)

        if self.ttype == "operator-prefix":
            return "-" + args[0].emit(ast, context=context)

        parent = self.parent(ast)
        # do not render the ^{1,2,..} part in a linest formula
        # TODO: bit of a hack
        if op == "**":
            if parent and parent.tvalue.lower() == "linest":
                return args[0].emit(ast, context=context)

        # work around because None < 0 is True (happens on blank cells)
        # TODO: make this work for string to number and string to string comparisons like Excel
        if op in ('xl_lt', 'xl_gt', 'xl_lte', 'xl_gte', 'xl_eq', 'xl_neq', 'full_concatenate'):
            ss = "{}({}, {})".format(op,
                                     args[0].emit(ast, context=context),
                                     args[1].emit(ast, context=context))
        else:
            ss = "{0} {1} {2}".format(args[0].emit(ast, context=context), op, args[1].emit(ast, context=context))

        # avoid needless parentheses
        if parent and not isinstance(parent, FunctionNode):
            ss = "(" + ss + ")"

        return ss


class OperandNode(ASTNode):
    def __init__(self,*args):
        super(OperandNode,self).__init__(*args)

    def emit(self,ast, context=None):
        t = self.tsubtype

        if t == "logical":
            return str(self.tvalue.lower() == "true")
        elif t == "text" or t == "error":
            # if the string contains quotes, escape them
            val = self.tvalue.replace('"', '\\"')
            return '"' + val + '"'
        else:
            return str(self.tvalue)


class RangeNode(OperandNode):
    """Represents a spreadsheet cell or range, e.g., A5 or B3:C20"""
    def __init__(self, *args):
        super(RangeNode, self).__init__(*args)

    def get_cells(self):
        return resolve_range(self.tvalue)[0]

    def emit(self,ast, context=None):
        # resolve the range into cells
        rng = self.tvalue.replace('$', '')
        sheet = context.curcell.sheet + "!" if context else ""
        if is_range(rng):
            sh, start, end = split_range(rng)
            if sh:
                result = 'eval_range("' + rng + '")'
            else:
                result = 'eval_range("' + sheet + rng + '")'
        else:
            sh, col, row = split_address(rng)
            if sh:
                result = 'eval_cell("' + rng + '")'
            else:
                result = 'eval_cell("' + sheet + rng + '")'

        return result


class FunctionNode(ASTNode):
    """AST node representing a function call"""
    def __init__(self,*args):
        super(FunctionNode, self).__init__(*args)
        self.numargs = 0

        # map  excel functions onto their python equivalents
        self.funmap = FUNCTION_MAP

    def emit(self, ast, context=None):
        fun = self.tvalue.lower()
        result = ''

        # Get the arguments
        args = self.children(ast)

        if fun == "atan2":
            # swap arguments
            result = "atan2(%s, %s)" % (args[1].emit(ast, context=context), args[0].emit(ast, context=context))
        elif fun == "pi":
            # constant, no parens
            result = "pi"
        elif fun == "if":
            # inline the if
            if len(args) == 2:
                result = "%s if %s else False" % (args[1].emit(ast, context=context),
                                                  args[0].emit(ast, context=context))
            elif len(args) == 3:
                result = "(%s if %s else %s)" % (args[1].emit(ast, context=context),
                                                 args[0].emit(ast, context=context),
                                                 args[2].emit(ast, context=context))
            else:
                raise Exception("if with %s arguments not supported" % len(args))

        elif fun == "array":
            result += '['
            if len(args) == 1:
                # only one row
                result += args[0].emit(ast, context=context)
            else:
                # multiple rows
                result += ", ".join(['[' + n.emit(ast, context=context) + ']' for n in args])
            result += ']'
        elif fun == "arrayrow":
            # simply create a list
            result += ", ".join([n.emit(ast, context=context) for n in args])
        # TODO: figure out what is "linestmario"
        elif fun == "linest" or fun == "linestmario":
            result = fun + "(" + ", ".join([n.emit(ast, context=context) for n in args])
            if not context:
                degree, coef = -1, -1
            else:
                # linests are often used as part of an array formula spanning multiple cells,
                # one cell for each coefficient.  We have to figure out where we currently are
                # in that range
                degree, coef = get_linest_degree(context.excel, context.curcell)

            # if we are the only linest (degree is one) and linest is nested -> return vector
            # else return the coef.
            if degree == 1 and self.parent(ast):
                if fun == "linest":
                    result += ", degree=%s)" % degree
                else:
                    result += ")"
            else:
                if fun == "linest":
                    result += ", degree=%s)[%s]" % (degree, coef-1)
                else:
                    result += ")[%s]" % (coef-1)

        elif fun == "and":
            result = "all([" + ", ".join([n.emit(ast, context=context) for n in args]) + "])"
        elif fun == "or":
            result = "any([" + ", ".join([n.emit(ast, context=context) for n in args]) + "])"
        else:
            # map to the correct name
            f = self.funmap.get(fun,fun)
            result = f + "(" + ", ".join([n.emit(ast, context=context) for n in args]) + ")"

        return result


def create_node(t):
    """Simple factory function"""
    if t.ttype == "operand":
        if t.tsubtype == "range":
            return RangeNode(t)
        else:
            return OperandNode(t)
    elif t.ttype == "function":
        return FunctionNode(t)
    elif t.ttype.startswith("operator"):
        return OperatorNode(t)
    else:
        return ASTNode(t)


class Operator(object):
    """Small wrapper class to manage operators during shunting yard"""
    def __init__(self, val, precedence, associativity):
        super(Operator, self).__init__()
        self.value = val
        self.precedence = precedence
        self.associativity = associativity


def shunting_yard(expression):
    """
    Tokenize an excel formula expression into reverse polish notation

    .. note::
        Core algorithm taken from wikipedia with varargs extensions from
        http://www.kallisti.net.nz/blog/2008/02/extension-to-the-shunting-yard-algorithm-to-allow-variable-numbers-of-arguments-to-functions/

    """
    # remove leading '='
    if expression.startswith('='):
        expression = expression[1:]

    p = ExcelParser()
    p.parse(expression)

    # insert tokens for '(' and ')', to make things clearer below
    tokens = []
    for t in p.tokens.items:
        if t.ttype == "function" and t.tsubtype == "start":
            t.tsubtype = ""
            tokens.append(t)
            tokens.append(f_token('(', 'arglist', 'start'))
        elif t.ttype == "function" and t.tsubtype == "stop":
            tokens.append(f_token(')', 'arglist', 'stop'))
        elif t.ttype == "subexpression" and t.tsubtype == "start":
            t.tvalue = '('
            tokens.append(t)
        elif t.ttype == "subexpression" and t.tsubtype == "stop":
            t.tvalue = ')'
            tokens.append(t)
        else:
            tokens.append(t)

    # All the MS Excel operators, as described in:
    # http://office.microsoft.com/en-us/excel-help/calculation-operators-and-precedence-HP010078886.aspx
    operators = {
        ':': Operator(':', 8, 'left'),
        '': Operator(' ', 8, 'left'),
        ',': Operator(',', 8, 'left'),
        'u-': Operator('u-', 7, 'left'),  # unary negation
        '%': Operator('%', 6, 'left'),
        '^': Operator('^', 5, 'left'),
        '*': Operator('*', 4, 'left'),
        '/': Operator('/', 4, 'left'),
        '+': Operator('+', 3, 'left'),
        '-': Operator('-', 3, 'left'),
        '&': Operator('&', 2, 'left'),
        '=': Operator('=', 1, 'left'),
        '<': Operator('<', 1, 'left'),
        '>': Operator('>', 1, 'left'),
        '<=': Operator('<=', 1, 'left'),
        '>=': Operator('>=', 1, 'left'),
        '<>': Operator('<>', 1, 'left'),
    }

    output = deque()
    stack = []
    were_values = []
    arg_count = []

    for t in tokens:
        if t.ttype == "operand":

            output.append(create_node(t))
            if were_values:
                were_values.pop()
                were_values.append(True)

        elif t.ttype == "function":

            stack.append(t)
            arg_count.append(0)
            if were_values:
                were_values.pop()
                were_values.append(True)
            were_values.append(False)

        elif t.ttype == "argument":
            while stack and (stack[-1].tsubtype != "start"):
                output.append(create_node(stack.pop()))

            if were_values.pop(): arg_count[-1] += 1
            were_values.append(False)

            if not len(stack):
                raise Exception("Mismatched or misplaced parentheses")

        elif t.ttype.startswith('operator'):
            if t.ttype.endswith('-prefix') and t.tvalue == "-":
                o1 = operators['u-']
            else:
                o1 = operators[t.tvalue]

            while stack and stack[-1].ttype.startswith('operator'):
                if stack[-1].ttype.endswith('-prefix') and stack[-1].tvalue == "-":
                    o2 = operators['u-']
                else:
                    o2 = operators[stack[-1].tvalue]

                if any((o1.associativity == "left" and o1.precedence <= o2.precedence,
                        o1.associativity == "right" and o1.precedence < o2.precedence)):
                    output.append(create_node(stack.pop()))
                else:
                    break
            stack.append(t)

        elif t.tsubtype == "start":
            stack.append(t)

        elif t.tsubtype == "stop":
            while stack and stack[-1].tsubtype != "start":
                output.append(create_node(stack.pop()))

            if not stack:
                raise Exception("Mismatched or misplaced parentheses")

            stack.pop()

            if stack and stack[-1].ttype == "function":
                f = create_node(stack.pop())
                a = arg_count.pop()
                w = were_values.pop()
                if w: a += 1
                f.num_args = a
                output.append(f)

    while stack:
        if stack[-1].tsubtype == "start" or stack[-1].tsubtype == "stop":
            raise Exception("Mismatched or misplaced parentheses")

        output.append(create_node(stack.pop()))

    # convert to list
    result = [x for x in output]
    return result


def build_ast(expression):
    """Build an AST from an Excel formula expression in reverse polish notation."""

    # use a directed graph to store the tree
    graph = DiGraph()

    stack = []

    for n in expression:
        # Since the graph does not maintain the order of adding nodes/edges
        # add an extra attribute 'pos' so we can always sort to the correct order
        if isinstance(n,OperatorNode):
            if n.ttype == "operator-infix":
                arg2 = stack.pop()
                arg1 = stack.pop()
                graph.add_node(arg1, {'pos': 1})
                graph.add_node(arg2, {'pos': 2})
                graph.add_edge(arg1, n)
                graph.add_edge(arg2, n)
            else:
                arg1 = stack.pop()
                graph.add_node(arg1, {'pos': 1})
                graph.add_edge(arg1, n)

        elif isinstance(n, FunctionNode):
            args = [stack.pop() for _ in range(n.num_args)]
            args.reverse()
            for i, a in enumerate(args):
                graph.add_node(a, {'pos': i})
                graph.add_edge(a, n)
        else:
            graph.add_node(n, {'pos': 0})

        stack.append(n)

    return graph, stack.pop()


class Context(object):
    """A small context object that nodes in the AST can use to emit code"""
    def __init__(self, curcell, excel):
        # the current cell for which we are generating code
        self.curcell = curcell
        # a handle to an excel instance
        self.excel = excel


class ExcelCompiler(object):
    """Class responsible for taking an Excel spreadsheet and compiling it to a Spreadsheet instance
       that can be serialized to disk, and executed independently of excel.
    """

    def __init__(self, filename=None, excel=None, visible=False):

        super(ExcelCompiler, self).__init__()
        self.filename = filename

        if excel:
            # if we are running as an excel addin, this gets passed to us
            self.excel = excel
        else:
            # TODO: use a proper interface so we can (eventually) support loading from file
            self.excel = wrapper_implementation(filename=filename, visible=visible)
            self.excel.connect()

    def cell2code(self,cell):
        """Generate python code for the given cell"""
        if cell.formula:
            e = shunting_yard(cell.formula or str(cell.value))
            ast, root = build_ast(e)
            code = root.emit(ast, context=Context(cell, self.excel))
        else:
            ast = None
            code = cell.value
        return code, ast

    def add_node_to_graph(self, graph, n):
        graph.add_node(n)
        graph.node[n]['sheet'] = n.sheet

        if isinstance(n, Cell):
            graph.node[n]['label'] = n.col + str(n.row)
        else:
            # strip the sheet
            graph.node[n]['label'] = n.address()[n.address().find('!')+1:]

    def make_python_model(self, seeds):
        pass

    def gen_graph(self, seed, sheet=None, formulas_replace=None):
        """
        Given a starting point (e.g., A6, or A3:B7) on a particular sheet,
        generate a Spreadsheet instance that captures the logic and control
        flow of the equations.

        """

        # starting points
        cursheet = sheet if sheet is not None else self.excel.get_active_sheet()
        self.excel.set_sheet(cursheet)

        # no need to output nr and nc here, since seed can be a list of unlinked cells
        seeds, nr, nc = Cell.make_cells(self.excel, seed, sheet=cursheet)
        seeds = list(flatten(seeds))

        logger.debug("Seed %s expanded into %s cells" % (seed, len(seeds)))

        # only keep seeds with formulas or numbers
        seeds = [s for s in seeds if s.formula or isinstance(s.value, number_types)]

        logger.debug("%s filtered seeds" % len(seeds))

        # cells to analyze: only formulas
        todo = [s for s in seeds if s.formula]

        logger.debug("%s cells on the todo list" % len(todo))

        # map of all cells
        cellmap = OrderedDict([(x.address(), x) for x in seeds])

        # directed graph
        graph = DiGraph()

        # match the info in cellmap
        for cell in itervalues(cellmap):
            self.add_node_to_graph(graph, cell)

        while todo:
            c1 = todo.pop()

            logger.debug("Handling {}".format(c1.address()))

            # set the current sheet so relative addresses resolve properly
            if c1.sheet != cursheet:
                cursheet = c1.sheet
                self.excel.set_sheet(cursheet)

            # parse the formula into code
            # HACK??? TODO
            if formulas_replace:
                if c1.formula:
                    for fr in formulas_replace:
                        new_formula = c1.formula.replace(fr[0], fr[1])
                        if new_formula != c1.formula:
                            c1.set_formula(new_formula)

            pystr, ast = self.cell2code(c1)

            # set the code & compile it (will flag problems sooner rather than later)
            c1.python_expression = pystr
            c1.compile()

            # get all the cells/ranges this formula refers to
            deps = [x.tvalue.replace('$', '') for x in ast.nodes() if isinstance(x, RangeNode)]

            # remove dupes
            deps = uniqueify(deps)

            for dep in deps:
                # if the dependency is a multi-cell range, create a range object
                if is_range(dep):
                    # this will make sure we always have an absolute address
                    rng = CellRange(dep, sheet=cursheet)

                    if rng.address() in cellmap:
                        # already dealt with this range
                        # add an edge from the range to the parent
                        graph.add_edge(cellmap[rng.address()], cellmap[c1.address()])
                        continue
                    else:
                        # turn into cell objects
                        cells, nrows, ncols = Cell.make_cells(self.excel, dep, sheet=cursheet)

                        # get the values so we can set the range value
                        if nrows == 1 or ncols == 1:
                            rng.value = [c.value for c in cells]
                        else:
                            rng.value = [[c.value for c in cells[j]] for j in range(len(cells))]

                        # save the range
                        cellmap[rng.address()] = rng
                        # add an edge from the range to the parent
                        self.add_node_to_graph(graph, rng)
                        graph.add_edge(rng, cellmap[c1.address()])
                        # cells in the range should point to the range as their parent
                        target = rng
                else:
                    # not a range, create the cell object
                    cells = [Cell.resolve_cell(self.excel, dep, sheet=cursheet)]
                    target = cellmap[c1.address()]

                # process each cell
                for c2 in flatten(cells):
                    # if we haven't treated this cell already
                    if c2.address() not in cellmap:
                        if c2.formula:
                            # cell with a formula, needs to be added to the `todo` list
                            todo.append(c2)
                        else:
                            # constant cell, no need for further processing, just remember to set the code
                            pystr, ast = self.cell2code(c2)
                            c2.python_expression = pystr
                            if not isinstance(pystr, (int, float)):
                                c2.compile()

                        # save in the cellmap
                        cellmap[c2.address()] = c2
                        # add to the graph
                        self.add_node_to_graph(graph, c2)

                    # add an edge from the cell to the parent (range or cell)
                    graph.add_edge(cellmap[c2.address()], target)

        logger.info("Graph construction done, %s nodes, %s edges, %s cellmap entries" % (len(graph.nodes()),
                                                                                         len(graph.edges()),
                                                                                         len(cellmap)))

        sp = Spreadsheet(graph=graph, cellmap=cellmap, filename=self.filename)

        return sp


if __name__ == '__main__':
    # some test formulas
    inputs = [
              '=A1 & A2 & A3',
              '=concatenate(A1,A2,A3)',
              '=I158*_get_sum_bk4_bk109958(D154,D155)^-3',
              '=SUM((A:A 1:1))',
              '=A1',
              '=atan2(A1,B1)',
              '=5*log(sin()+2)',
              '=5*log(sin(3,7,9)+2)',
              '=3 + 4 * 2 / ( 1 - 5 ) ^ 2 ^ 3',
              '=1+3+5',
              '=3 * 4 + 5',
              '=50',
              '=1+1',
              '=$A1',
              '=$B$2',
              '=SUM(B5:B15)',
              '=SUM(B5:B15,D5:D15)',
              '=SUM(B5:B15 A7:D7)',
              '=SUM(sheet1!$A$1:$B$2)',
              '=[data.xls]sheet1!$A$1',
              '=SUM((A:A,1:1))',
              '=SUM((A:A A1:B1))',
              '=SUM(D9:D11,E9:E11,F9:F11)',
              '=SUM((D9:D11,(E9:E11,F9:F11)))',
              '=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))',
              '={SUM(B2:D2*B3:D3)}',
              '=SUM(123 + SUM(456) + (45<6))+456+789',
              '=AVG(((((123 + 4 + AVG(A1:A2))))))',

              # E. W. Bachtal's test formulae
              '=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &   "  more ""test"" text"',
              # '=+ AName- (-+-+-2^6) = {"A","B"} + @SUM(R1C1) + (@ERROR.TYPE(#VALUE!) = 2)',
              '=IF(R13C3>DATE(2002,1,6),0,IF(ISERROR(R[41]C[2]),0,IF(R13C3>=R[41]C[2],0, IF(AND(R[23]C[11]>=55,R[24]C[11]>=20),R53C3,0))))',
              '=IF(R[39]C[11]>65,R[25]C[42],ROUND((R[11]C[11]*IF(OR(AND(R[39]C[11]>=55, ' +
                  'R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")),R[44]C[11],R[43]C[11]))+(R[14]C[11] ' +
                  '*IF(OR(AND(R[39]C[11]>=55,R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")), ' +
                  'R[45]C[11],R[43]C[11])),0))',
              '=(propellor_charts!B22*(propellor_charts!E21+propellor_charts!D21*(engine_data!O16*D70+engine_data!P16)+propellor_charts!C21*(engine_data!O16*D70+engine_data!P16)^2+propellor_charts!B21*(engine_data!O16*D70+engine_data!P16)^3)^2)^(1/3)*(1*D70/5.33E-18)^(2/3)*0.0000000001*28.3495231*9.81/1000',
              '=(3600/1000)*E40*(E8/E39)*(E15/E19)*LN(E54/(E54-E48))',
              '=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))',
              '=LINEST(X5:X32,W5:W32^{1,2,3})',
              '=IF(configurations!$G$22=3,sizing!$C$303,M14)',
              '=0.000001042*E226^3-0.00004777*E226^2+0.0007646*E226-0.00075',
              '=LINEST(G2:G17,E2:E17,FALSE)',
              '=IF(AI119="","",E119)',
              '=LINEST(B32:(INDEX(B32:B119,MATCH(0,B32:B119,-1),1)),(F32:(INDEX(B32:F119,MATCH(0,B32:B119,-1),5)))^{1,2,3,4})',
              ]

    for i in inputs:
        print("**************************************************")
        print("Formula: ", i)

        e = shunting_yard(i)
        print("RPN: ",  "|".join([str(x) for x in e]))

        graph, root = build_ast(e)

        print("Python code: ", root.emit(graph, context=None))
        print("**************************************************")
