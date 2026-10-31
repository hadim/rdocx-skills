import sys

import docx_ops

src, out, expected = sys.argv[1:4]
docx_ops.replace_batch(src, out, [("Service Level Agreement", "Support Contract", int(expected))])
