import re
from PySide6.QtCore import Qt

class ASTNode:
    def to_string(self) -> str:
        raise NotImplementedError

class NumberNode(ASTNode):
    def __init__(self, value):
        self.value = float(value)
    def to_string(self) -> str:
        if self.value == int(self.value):
            return str(int(self.value))
        return str(self.value)

class CellRefNode(ASTNode):
    def __init__(self, cell_name):
        self.cell_name = cell_name.upper()
    def to_string(self) -> str:
        return self.cell_name

class ErrorNode(ASTNode):
    def __init__(self, error_code: str):
        self.error_code = error_code.lstrip("=") 
    def to_string(self) -> str:
        return self.error_code

class RangeRefNode(ASTNode):
    def __init__(self, start_cell, end_cell):
        self.start_cell = start_cell.upper()
        self.end_cell = end_cell.upper()
    def to_string(self) -> str:
        return f"{self.start_cell}:{self.end_cell}"

class BinaryOpNode(ASTNode):
    def __init__(self, left, op, right):
        self.left = left
        self.op = op
        self.right = right
    def to_string(self) -> str:
        left_str = self.left.to_string()
        if isinstance(self.left, BinaryOpNode):
            left_str = f"({left_str})"
            
        right_str = self.right.to_string()
        if isinstance(self.right, BinaryOpNode):
            right_str = f"({right_str})"

        return f"{left_str}{self.op}{right_str}"


class FunctionNode(ASTNode):
    def __init__(self, func_name, args):
        self.func_name = func_name.upper()
        self.args = args
    def to_string(self) -> str:
        args_str = ",".join(arg.to_string() for arg in self.args)
        return f"{self.func_name}({args_str})"

class UnaryOpNode(ASTNode):
    def __init__(self, op, operand):
        self.op = op
        self.operand = operand
    def to_string(self) -> str:
        operand_str = self.operand.to_string()
        if isinstance(self.operand, BinaryOpNode):
             operand_str = f"({operand_str})"
        return f"{self.op}{operand_str}"

class ParsingError(Exception):
    pass

class CircularReferenceError(Exception):
    pass

class ReferenceError(Exception):
    pass

class Token:
    def __init__(self, type, value):
        self.type = type
        self.value = value
    def __repr__(self):
        return f"Token({self.type}, {self.value})"

class Lexer:
    TOKEN_SPECS = [
        ('NUMBER',   r'\d+(\.\d*)?'),
        ('REF_ERROR',r'#REF!'),
        ('NAME_ERROR',r'#NAME\?'),
        ('FUNCTION', r'[A-Z_]+(?=\()'), 
        ('CELL',     r'[A-Z]+[0-9]+'),
        ('PLUS',     r'\+'),
        ('MINUS',    r'-'),
        ('MUL',      r'\*'),
        ('DIV',      r'/'),
        ('POW',      r'\^'),
        ('LPAREN',   r'\('),
        ('RPAREN',   r'\)'),
        ('COMMA',    r','),
        ('COLON',    r':'),
        ('EQUALS',   r'='),
        ('SKIP',     r'[ \t]+'),
        ('MISMATCH', r'.'),
    ]
    TOKEN_REGEX = '|'.join(f'(?P<{name}>{regex})' for name, regex in TOKEN_SPECS)

    def tokenize(self, text):
        tokens = []
        formula_body = text.lstrip("=").upper()
        
        for mo in re.finditer(self.TOKEN_REGEX, formula_body):
            kind = mo.lastgroup
            value = mo.group()
            
            if kind == 'SKIP':
                continue
            if kind == 'MISMATCH':
                if value == '#' and len(formula_body) > mo.start() + 1:
                    continue 
                raise ParsingError(f"Невідомий символ: {value}")
                
            tokens.append(Token(kind, value))
        return tokens

class Parser:
    def __init__(self):
        self.tokens = []
        self.current_token = None
        self.pos = -1

    def _advance(self):
        self.pos += 1
        if self.pos < len(self.tokens):
            self.current_token = self.tokens[self.pos]
        else:
            self.current_token = None

    def _eat(self, token_type):
        if self.current_token and self.current_token.type == token_type:
            self._advance()
        else:
            raise ParsingError(f"Очікувався {token_type}, але знайдено {self.current_token.type if self.current_token else 'кінець'}")

    def parse(self, formula_string: str) -> ASTNode:
        if not formula_string.startswith("="):
             raise ParsingError("Формула має починатися з '='")
             
        try:
            self.tokens = Lexer().tokenize(formula_string)
        except ParsingError as e:
            return ErrorNode("#ERROR!")

        self.pos = -1
        self._advance()
        
        if not self.tokens:
            return NumberNode(0)

        ast_tree = self._parse_expression()
        if self.current_token is not None:
             raise ParsingError(f"Неочікуваний токен у кінці: {self.current_token}")
        return ast_tree

    def _parse_expression(self):
        node = self._parse_term()
        while self.current_token and self.current_token.type in ('PLUS', 'MINUS'):
            op_token = self.current_token
            self._eat(op_token.type)
            node = BinaryOpNode(left=node, op=op_token.value, right=self._parse_term())
        return node

    def _parse_term(self):
        node = self._parse_factor()
        while self.current_token and self.current_token.type in ('MUL', 'DIV'):
            op_token = self.current_token
            self._eat(op_token.type)
            node = BinaryOpNode(left=node, op=op_token.value, right=self._parse_factor())
        return node

    def _parse_factor(self):
        node = self._parse_primary()
        if self.current_token and self.current_token.type == 'POW':
            op_token = self.current_token
            self._eat('POW')
            node = BinaryOpNode(left=node, op=op_token.value, right=self._parse_factor())
        return node

    def _parse_primary(self):
        token = self.current_token
        if token is None: raise ParsingError("Неочікуваний кінець формули")

        if token.type == 'REF_ERROR':
            self._eat('REF_ERROR')
            return ErrorNode("#REF!")
        if token.type == 'NAME_ERROR':
            self._eat('NAME_ERROR')
            return ErrorNode("#NAME?")
            
        if token.type == 'NUMBER':
            self._eat('NUMBER')
            return NumberNode(token.value)
        
        if token.type == 'MINUS': 
            self._eat('MINUS')
            return UnaryOpNode(op='-', operand=self._parse_factor())

        if token.type == 'CELL':
            cell_token = token
            self._eat('CELL')
            if self.current_token and self.current_token.type == 'COLON':
                self._eat('COLON')
                end_token = self.current_token
                self._eat('CELL')
                return RangeRefNode(cell_token.value, end_token.value)
            return CellRefNode(cell_token.value)
        
        if token.type == 'FUNCTION':
            func_name = token.value
            self._eat('FUNCTION')
            self._eat('LPAREN')
            args = []
            if self.current_token and self.current_token.type != 'RPAREN':
                args.append(self._parse_expression())
                while self.current_token and self.current_token.type == 'COMMA':
                    self._eat('COMMA')
                    args.append(self._parse_expression())
            self._eat('RPAREN')
            return FunctionNode(func_name, args)

        if token.type == 'LPAREN':
            self._eat('LPAREN')
            node = self._parse_expression()
            self._eat('RPAREN')
            return node
            
        raise ParsingError(f"Неочікуваний токен: {token}")