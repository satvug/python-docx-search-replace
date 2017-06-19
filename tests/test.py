# Simple example of python-docx-search-replace package
import re
from dxsr import dxsr

# Find all occurences of "bunny" and "bunnies", ignoring case, and swap the case of all such occurences.
doc = dxsr("test.docx")
pattern = re.compile("bunn(y|ies)", re.IGNORECASE)
matches = doc.search_paragraphs(pattern)
dxsr.list_matches(matches)
doc.replace_all(matches, dxsr.replace_func_swapcase)

# Use fancy sub-method to replace "(anyword) World" with "(anyword) Bunny"
pat = re.compile(r"(\w+) (world)", re.I)
pat_r = r"\1 Bunny"
doc.sub(pat, pat_r)

# Other kind of replacement
doc.search_replace("strange", "super neat")

# Save changes to file
doc.save_docx("test-modified.docx")
