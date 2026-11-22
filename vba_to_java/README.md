# VBA to pseudocode

Convert a VBA module from the command line:

```bash
python vba_to_java/vba_to_pseudocode.py files/vba_sample_1.bas
```

By default, the output is written beside the input with a `.pseudocode.txt`
suffix. Choose another destination with `-o`.

The converter is also callable from Python:

```python
from vba_to_java.vba_to_pseudocode import main

main("files/vba_sample_1.bas", "files/vba_sample_1.pseudo.txt")
```

The converter handles common procedures, declarations, assignments, conditionals,
loops, `Select Case`, `With`, labels, and basic error-handling statements. It is
a deterministic source-to-explanation utility, not a complete VBA parser or
compiler: unrecognized statements are retained as `EXECUTE` steps, so review
generated pseudocode before relying on it for business logic or control-flow
equivalence. The five `files/vba_sample_*.bas` modules provide progressively
more complex examples.
