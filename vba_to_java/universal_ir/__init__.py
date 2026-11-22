"""Universal, language-neutral intermediate representation."""
from .builder import build_ir
from .io import IR_VERSION, load_ir, new_program, save_ir
from .validator import validate_program

__all__ = ["IR_VERSION", "build_ir", "load_ir", "new_program", "save_ir", "validate_program"]
