

# python-docx-search-replace (dxsr)
*Copyright (c) 2017, Gustav Jensen*

A Python library for performing search/replace-operations in Microsoft Word 2007+ (.docx) documents.

This small package allows you to search in .docx documents using regexes, and it allows you to replace matches according to your liking, by either regex, constant string or your own replacement function.

## Basic usage
Here is an example, showing most common usage:

```python
from dxsr import dxsr
import re

# Load a document.
doc = dxsr("test.docx")

# Search using a specific regex pattern
pattern = re.compile("bunn(y|ies)", re.IGNORECASE) # search for "bunny" and "bunnies"
matches = doc.search_paragraphs(pattern)

# Use built in replacement function to swap case of all found matches.
doc.replace_all(matches, dxsr.replace_func_swapcase)

# Simpler form of search-replacement
doc.search_replace(re.compile(".cow."), "cat") # search using regex ".cow.", replace with "cat"
doc.search_replace(".bee.", "dog") # replace raw string ".bee." with "dog"

# Advanced: using our own replacement function
# If the match starts with 'a', replace match to "hello", otherwise "bye"
def replacefunc_example(text_match, hyperlinks, text_objects):
    replaced_text = text_match
    if text_match[0].lower() == "a":
        replaced_text = "hello"
    else:
        replaced_text = "bye"

    return (replaced_text, None) # None == URL of hyperlink, only applicable if match had hyperlink attached

# search for "raisin" and "apple", ignoring case.
# using our custom replacement function, "raisin" will be replaced with "bye", and "apple" with "hello"
matches = doc.search_paragraphs(re.compile("(raisin|apple)", re.I))
dxsr.list_matches(matches) # list matches for fun
a = doc.replace_all(matches, replacefunc_example)

# Finally, let us save our changed document into a new file. The original is untouched.
doc.save_docx("test-modified.docx")
```

## Bug
There is currently a case that is not handled properly and needs to be fixed.
The bug only occurs when you're replacing matches that are somewhat close to each other (that is, they're inside the same w:t object).



Example sentence:

`This short example.`

If we run

`doc.search_replace(re.compile("(short|example)"), "ultra long")`

on the above content, the result will be messed up.
What happens is that after the first occurence `short` is modified to `ultra long`, the position of the second occurence `example` will have changed, but the function is unaware of that change in position, and so will insert the second replacement at a wrong position.

The result should be
`This ultra long ultra long.`, but turns out as `This ultra ultra longample.`
