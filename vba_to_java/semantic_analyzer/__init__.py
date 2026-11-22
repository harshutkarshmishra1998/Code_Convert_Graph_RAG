"""Semantic analysis and normalization for language-neutral Universal IR."""
from .normalize import normalize_ir, run_files
from .expression_parser import parse_expression, parse_vba_call_statement, ParseError
from .normalized_validator import validate_normalized_ir

__all__ = ['normalize_ir','run_files','parse_expression','parse_vba_call_statement','ParseError','validate_normalized_ir']
__version__ = '0.1.0'
