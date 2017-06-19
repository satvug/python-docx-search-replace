# Simple example of python-docx-search-replace package
import re
from dxsr import dxsr

# Find all occurences of "bunny" and "bunnies", ignoring case, and swap the case of all such occurences.
doc = dxsr("test.docx")
pattern = re.compile("bunn(y|ies)", re.IGNORECASE)
matches = doc.search_paragraphs(pattern)
doc.replace_all(matches, dxsr.replace_func_swapcase)

# Use fancy sub-method to replace "(anyword) World" with "(anyword) Bunny"
pat = re.compile(r"(\w+) (world)", re.I)
pat_r = r"\1 Bunny"
doc.sub(pat, pat_r)

# Other kind of replacement
doc.search_replace("strange", "super neat")

# Advanced: using our own replacement function
def replacefunc_example(text_match, hyperlinks, text_objects):
    replaced_text = text_match
    if text_match[0].lower() == "a":
        replaced_text = "hello"
    else:
        replaced_text = "bye"

    return (replaced_text, None) # None == a hyperlink if we desire it

matches = doc.search_paragraphs(re.compile("(raisin|apple)", re.I))
dxsr.list_matches(matches) # list matches for fun
a = doc.replace_all(matches, replacefunc_example)

# Demonstration of bug that occurs when several matches inside the same w:t object are replaced.
# "This short example."
# will become
# "This ultra ultra longample."
# whereas it should be "This ultra long ultra long."
doc.search_replace(re.compile("(short|example)"), "ultra long")

# Save changes to file
doc.save_docx("test-modified.docx")
