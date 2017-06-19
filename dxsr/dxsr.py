#!/usr/bin/env python
# (C) Copyright 2017 by Gustav Jensen

"""
This class is used to search and replace in Microsoft Word 2007+ documents (.docx files).
Example usage:

    from dxsr import dxsr
    import re
    doc = dxsr("test.docx")
    pattern = re.compile("bunn(y|ies)", re.IGNORECASE)
    matches = doc.search_paragraphs(pattern)
    dxsr.list_matches(matches)
    doc.replace_all(matches, dxsr.replace_func_swapcase)
    doc.search_replace(re.compile(".cow."), "cat") # searches for regex ".cow.", replace with "cat"
    doc.save_docx("test-modified.docx")

"""
import zipfile, tempfile
from lxml import etree as ET
from shutil import copyfile
from pprint import pprint
import os
import re
import subprocess
import argparse
import json
import logging
import collections
import time


class dxsr:
    def __init__(self, infile, verbose=False):
        self.infile = infile
        self.verbose = verbose

        # default values
        self.peek_matches   = True
        self.peek_max_chars = 40

        self.log = logging.getLogger(__name__)
        logging.basicConfig(format="%(levelname)s: %(message)s")
        if self.verbose:
            self.log.setLevel(logging.DEBUG)
            self.log.info("Verbose output.")

        self.load_document()


    """ if any unwritten replacements were done, this method can be used to discard them and reload the document in its current state on disk """
    def load_document(self):
        self._load_docx()
        self._read_relationships()
        self._read_paragraphs()


    """
    Load the zip-file and load the document.xml file as well as relationships file inside the zip-file via lxml ElementTree.
    """
    def _load_docx(self):
        zf = zipfile.ZipFile(self.infile)
        if self.verbose:
            print("Files in archive", self.infile, ":", zf.namelist())

        # open relations document to read/modify hyperlinks/references in there
        relsdoc = zf.read("word/_rels/document.xml.rels")
        self.relsroot = ET.fromstring(relsdoc)

        # open document.xml file which stores main Word document info/text
        worddoc = zf.read("word/document.xml")
        self.docroot = ET.fromstring(worddoc)


    """
    Create a mapping of relationship ids to relationships.
    Example: some text in the document contains a hyperlink to a website.
    In practice, it means the text contains a reference to a relationship from the rels document with a certain id which then contains info about the hyperlink.
    The motivation is we want to be able to edit hyperlink URLs as well - not only the text as it appears in the document.
    Example dict:
        {'rId1': {
            'Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
            'rId': 'rId1',
            'Target': 'http://mylink.example.com',
            'TargetMode': 'External'
        }   }
    """
    def _read_relationships(self):
        children = self.relsroot.getchildren()
        self.rels_dict = {}
        for child in children:
            d = child.attrib
            if not 'Id' in d:
                raise Exception("Relationship object does not contain key 'Id'! Object:", d)
            self.rels_dict[ d['Id'] ] = child


    """
    Load paragraphs in the document so they can be searched through.
    Populates self.paragraph_map, which maps each paragraph object to a list of text objects
    """
    def _read_paragraphs(self):
        paragraphs = self.docroot.xpath('.//w:p', namespaces=self.docroot.nsmap)
        # Problem: the above xpath command only returns "top level" w:p elements.
        # If a w:p element contains sub-paragraphs i.e. more w:p elements, they are not included in above result.
        # We need to check if each found paragraph has sub-paragraphs and add those to the list.
        to_add = []
        for par in paragraphs:
            # Does this paragraph have any sub-paragraphs? This can be the case for text boxes.
            subpars = par.xpath('.//w:p', namespaces=self.docroot.nsmap)
            if len(subpars) > 0:
                # get number of w:t elements that are descendants of this paragraph
                n_texts = len(par.xpath('.//w:t', namespaces=self.docroot.nsmap))
                n_texts_in_subpars = 0
                to_add += subpars
                for subpar in subpars:
                    n_texts_in_subpars += len(subpar.xpath('.//w:t', namespaces=self.docroot.nsmap))
                if n_texts_in_subpars == n_texts:
                    # this paragraph can be safely removed since it contains no text objects that are not contained by child paragraphs
                    paragraphs.remove(par)
                else:
                    # We must keep this paragraph in the list because it contains text objects not contained by any child paragraph.
                    pass

        if len(to_add) > 0:
            paragraphs += to_add

        # find out which text objects belong to each paragraph
        # our dict must be ordered so the list of paragraphs we return is in correct order (start of document --> end of document)
        self.paragraph_map = collections.OrderedDict()
        for paragraph in paragraphs:
            text_objects = self.text_objects_in_paragraph(paragraph)
            self.paragraph_map[paragraph] = text_objects


    """ get nearest parent paragraph of an object. This might raise an exception if no parents are paragraphs! """
    @staticmethod
    def nearest_paragraph_parent(obj):
        iterator = obj
        while True:
            iterator = iterator.getparent()
            if re.match(".*}p$", iterator.tag):
                return iterator
        raise Exception("Should not be reached")


    """ given a paragraph object, return all text objects contained by it, not including text objects contained by sub-paragraphs! """
    def text_objects_in_paragraph(self, paragraph_object):
        # find all w:t children tags in this paragraph; w:t objects contain all text
        text_objects = paragraph_object.xpath('.//w:t', namespaces=self.docroot.nsmap)

        # If we received any text objects whose closest paragraph parent isn't THIS paragraph, remove them. Those objects belongs to a sub-paragraph.
        for obj in text_objects:
            nearest_paragraph = dxsr.nearest_paragraph_parent(obj)
            if nearest_paragraph != paragraph_object:
                text_objects.remove(obj)
                self.log.debug("Removed w:t object from paragraph that belongs to a sub-paragraph.")
            elif obj.text == None:
                text_objects.remove(obj)
                self.log.debug("Removed w:t object with no text!")

        return text_objects


    def get_paragraphs(self):
        # return a list of (paragraph, [text_objects]) pairs
        return self.paragraph_map.iteritems()


    """ return True if the given variable is of same type as re.compile() """
    static_re_object_type = type(re.compile(''))
    @staticmethod
    def is_sre_pattern(var):
        return type(var) == dxsr.static_re_object_type


    """ create directory if it does not exist already """
    @staticmethod
    def mkdir_if_needed(path):
        if not os.path.exists(path):
            os.makedirs(path)


    def default_new_filename(self):
        infile_noext, infile_ext = os.path.splitext(self.infile)
        outfile = "%s-new%s" % (os.path.basename(infile_noext), infile_ext)
        return outfile


    """
    Used to update an existing ZipFile with new files, overwriting if needed
    inspired from http://stackoverflow.com/a/25739108/6352803
    """
    @staticmethod
    def updateZip(zipname, files_data_dict):
        # generate a temp file
        tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(zipname))
        # print "Using temp name", tmpname
        os.close(tmpfd)

        # create a temp copy of the archive without the files to be updated
        with zipfile.ZipFile(zipname, 'r') as zin:
            with zipfile.ZipFile(tmpname, 'w') as zout:
                for item in zin.infolist():
                    if item.filename not in files_data_dict:
                        zout.writestr(item, zin.read(item.filename))

        # replace with the temp archive
        if os.path.exists(zipname):
            os.remove(zipname)
            time.sleep(0.2) # race condition on Windows? weird stuff happened
        os.rename(tmpname, zipname)

        # now add new files
        with zipfile.ZipFile(zipname, mode='a', compression=zipfile.ZIP_DEFLATED) as zf:
            for (filename, data) in files_data_dict.iteritems():
                zf.writestr(filename, data)



    """
    Save the loaded document to a new file; this will include all changes made by search-replacements or
    any other modifications to the XML trees.
    """
    def save_docx(self, outfile, overwrite=False):
        if outfile == "":
            outfile = self.default_new_filename()

        # check if we are overwriting a file
        exists = (self.infile == outfile or os.path.exists(outfile))
        if exists:
            if overwrite == True:
                print "Warning: overwriting file", outfile
            else:
                raise Exception("Cannot save to file", outfile, "- file with that name already exists. Use overwrite=True if you wish to overwrite the file.")

        # make copy of original docx file so we can update files inside it
        if self.infile != outfile:
            copyfile(self.infile, outfile)

        outxml = "word/document.xml"
        outrels = "word/_rels/document.xml.rels"
        d = { outxml:  ET.tostring(self.docroot),
              outrels: ET.tostring(self.relsroot) }
        dxsr.updateZip(outfile, d)

        print "Document written to file", outfile


    """
    Take a list of w:t objects and return the concatenated text.
    By default, no delimiter between each object is inserted, but this can be added.
    """
    @staticmethod
    def objects_to_text(text_objects, delim=""):
        txt = ""
        for obj in text_objects:
            if obj.text != None:
                txt += obj.text + delim
        return txt


    """
    This function takes a list of <w:t> objects from the document.xml file
    and concatenates all text (assumed to be ordered) so that it can be searched.
    Search is then performed on the concatenated text with the given search patterns.
    The method returns a list of dicts with useful info about each found match, such as
    start position and end position so that the match can be replaced easily later.

    Note that if text_objects span multiple paragraphs, the paragraphs are joined with "" i.e. nothing, like
    "this is paragraph1this is paragraph2".
    Matches spanning multiple paragraphs are therefore not supported yet.
    """
    def find_matches(self, text_objects, patterns, match_modifiers=[]):
        # Make a list of chars in the text objects along with the "object" id for every character, as well as position in that object.
        # This is so we can see at which object a match begins and which one it ends at
        txt_map = []
        for obj in text_objects:
            i=0
            if obj.text == None:
                # kan happen in documents that had replacements done via this module on matches spanning several objects - some objects can afterward contain no text
                # raise Exception("Object", obj, "does not have a .text attribute!")
                continue
            for char in obj.text:
                txt_map.append( {'char': char, # an actual character in the match
                    'text-object': obj, # the w:t object that the character belongs to
                    'charpos': i}  # the position in this text object where this char occurs
                )
                i += 1

        if len(txt_map) == 0:
            return []

        # Get full text by joining the w:t text elements. This is so we can search the text.
        # The text can be safely joined as the text objects already contain spaces etc.
        # (Except if the objects span more than one paragraph)
        txt = dxsr.objects_to_text(text_objects)

        # search for matches in the joined text
        len_txt_map = len(txt_map)
        matches_dict = []
        for pattern in patterns:
            for match in pattern.finditer(txt):
                if match.start() >= len_txt_map:
                    # this weird case can occur with ".*" regex
                    continue
                match_info = self._get_match_info(match, txt_map, match_modifiers)
                matches_dict.append(match_info)
        return matches_dict



    """ return a list of unique text-objects in order of appearance in indexed part of txt_map """
    @staticmethod
    def objects_from_txt_map(i_start, i_end, txt_map):
        objects = []
        for t in txt_map[i_start:i_end+1]:
            obj = t['text-object']
            if not obj in objects:
                objects.append(obj)
        return objects

    """ extract text from txt_map, given index range [i_start; i_end] """
    @staticmethod
    def text_from_txt_map(i_start, i_end, txt_map):
        return "".join( [ i['char'] for i in txt_map[i_start:i_end+1] ] )


    """ return a list of hyperlinks given a list of w:t objects. The hyperlinks returned are connected to one/more of these w:t objects """
    def hyperlinks_for_text_objects(self, text_objects):
        # find rId in parents of the w:t objects
        hyperlink_ids = []
        for obj in text_objects:
            # is the parent-parent object a <w:hyperlink>? If so, get the rId of it
            pp = obj.getparent().getparent()
            if re.match(".*hyperlink$", pp.tag):
                rIds = []
                for val in pp.values():
                    for rId in re.findall("^rId\d*$", val):
                        rIds.append(rId)
                if len(rIds) == 1:
                    if not rIds[0] in hyperlink_ids:
                        hyperlink_ids.append(rIds[0])
                elif len(rIds) > 1:
                    raise Exception("Strange, a hyperlink parent object had more than one rId:", rIds, " - this is an unknown case and not handled.")
                else:
                    pass # might be a ToC reference, these have no rId

        hyperlink_objects = []
        # for each found rId reference, get a reference to our own hyperlinks dictionary
        for hl_id in hyperlink_ids:
            if hl_id in self.rels_dict:
                hyperlink_objects.append( self.rels_dict[hl_id] )
            else:
                print "Relationships dict:"
                pprint(self.rels_dict)
                raise Exception("Hyperlink rId", hl_id, ", was not found in list of ids read from rels file! Apparently some text contains a reference to a hyperlink which hasn't been defined in the document.")
        return hyperlink_objects


    @staticmethod
    def clamp_max(n, n_max):
        return n_max if n > n_max else n


    @staticmethod
    def check_txt_map_bounds(i_start, i_end, txt_map):
        if i_start < 0 or i_start >= len(txt_map) or i_end < 0 or i_end >= len(txt_map) or i_end < i_start:
            raise Exception("Bad i_start or i_end - i_start: %d, i_end: %d, len(txt_map): %d" %
                (i_start, i_end, len(txt_map)))

    """
    This function takes a regex match object and the text-map of the paragraph containing this match,
    and returns a dict containing the text of the match, a list of text-objects containing (all parts of)
    the match, startpos and endpos and a reference to hyperlink(s) if found.

    See dict structure below.

    returned dict structure of a match found in document
    {'text': 'http:/subversion/svn/proj/test.docx', # the actual text match
     'objects': [0x001, 0x002, 0x004], # list of w:t objects in order of appearance which contain the match text
     'startpos': 20,  # the index in the first object.text where match begins
     'endpos': 15,   # the index in the last object.text where the match ends
     'hyperlinks': [0x133],  # a list of references to object in rels document, so hyperlinks can be viewed/changed if wanted.
     're_object': sre.SRE_Match object # match object, containing all info about the match like groups, etc.
     'context': a larger string showing part of what comes before and after the matched text
    """
    def _get_match_info(self, re_match, txt_map, match_modifiers):
        i_start = re_match.start()   # start index of matched text
        i_end   = re_match.end() - 1 # end index of matched text

        if len(txt_map) == 0:
            raise Exception("Empty txt_map!", re_match)

        dxsr.check_txt_map_bounds(i_start, i_end, txt_map)

        match_text = dxsr.text_from_txt_map(i_start, i_end, txt_map)
        if match_text != re_match.group():
            print "match_from_txt_map:", match_text
            print "match.group():", re_match.group()
            raise Exception("Fatal error, match_from_txt_map != re_match.group()!")

        self.log.info("Found match: '%s'" % (match_text))

        # Add all text objects in the match to a list. This must be in correct order!
        match_objects = dxsr.objects_from_txt_map(i_start, i_end, txt_map)

        # Check if there are hyperlinks connected to this text
        hyperlink_objects = self.hyperlinks_for_text_objects(match_objects)

        # if for some reason we need to shrink/grow the match by cutting off or adding characters in the end or beginning, this can be done here via 'match modifier functions'
        for modifier in match_modifiers:
            (new_i_start, new_i_end) = modifier(i_start, i_end, txt_map, hyperlink_objects)
            if new_i_start != i_start or new_i_end != i_end:
                # The modifier function expanded/shrinked the match
                # We therefore have to rebuild the object and hyperlink list.
                dxsr.check_txt_map_bounds(new_i_start, new_i_end, txt_map)
                (i_start, i_end)  = (new_i_start, new_i_end)
                match_objects     = dxsr.objects_from_txt_map(i_start, i_end, txt_map)
                hyperlink_objects = self.hyperlinks_for_text_objects(match_objects)
                match_text        = dxsr.text_from_txt_map(i_start, i_end, txt_map)
                self.log.info("Match was modified to '%s'" % match_text)

        start_object_startpos = txt_map[i_start]['charpos']
        last_object_endpos = txt_map[i_end]['charpos']
        dict_entry = {
            'text':         match_text,
            'objects':      match_objects,
            'startpos':     start_object_startpos,
            'endpos':       last_object_endpos,
            'hyperlinks':   hyperlink_objects,
            're_object':    re_match # useful so replacement functions can check match groups etc.
        }

        # Get text coming before and after the match. Useful to see where in the document this match is.
        # However, we can only peek into the same paragraph as the match was found in.
        if self.peek_matches:
            max_peek_chars = self.peek_max_chars
            max_before  = dxsr.clamp_max(i_start, max_peek_chars)
            max_after   = dxsr.clamp_max(len(txt_map) - i_end, max_peek_chars)
            text_before = text_after = ""

            if i_start > 0:
                text_before = dxsr.text_from_txt_map(i_start-max_before, i_start-1, txt_map)
            if len(txt_map) > i_end:
                text_after = dxsr.text_from_txt_map(i_end+1, i_end+max_after, txt_map)
            dict_entry['context'] = text_before + match_text + text_after

        return dict_entry


    """ helper function that can insert a string into a string at a given position (index) """
    @staticmethod
    def insert_str(string, str_to_insert, index):
        return string[:index] + str_to_insert + string[index:]


    """
    This function can be used as argument to replace_match, as an example.
    It simply swaps case of the text match and hyperlink if a hyperlink is associated with this text match
    """
    @staticmethod
    def replace_func_swapcase(text_match, hyperlinks, text_objects):
        if hyperlinks:
            hl = hyperlinks[0].swapcase()
        else:
            hl = None
        return (text_match.swapcase(), hl)

    """
    Replace all given search matches via replace_func.
    If max_replacements > 0, only perform replacements on [max_replacements] first matches.
    """
    def replace_all(self, matches, replace_func, max_replacements=0):
        if max_replacements < 0:
            raise ValueError("max_replacements must be >= 0")

        n_replacements = 0
        for match in matches:
            self.replace_match(match, replace_func)
            n_replacements += 1
            if n_replacements == max_replacements:
                print "Reached maximum of %d replacements." % max_replacements
                break
        return n_replacements

    """ replace a single search match, by using replacement function replace_func """
    def replace_match(self, match, replace_func):
        self.log.debug("Replace match '%s'.." % match['text'])

        match_objects = match['objects']
        self.log.debug("OBJECTS before mod, %d objects:" % len(match_objects))
        self.log.debug("'%s'" % dxsr.objects_to_text(match_objects, delim="|"))

        # Delete the entire match from the text objects, not touching the non-match parts.
        # Save the object and position in which to insert replacement text.
        # Note: if this match spans multiple objects and the objects contain different formatting, the object in which we decide to insert the replacement is of importance! Right now this is not handled at all - replacement is always inserted into final object of the matched text.
        startpos  = match['startpos']
        endpos    = match['endpos']
        first_obj = match_objects[0]
        last_obj  = match_objects[len(match_objects)-1]
        for obj in match_objects:
            original_obj_text = obj.text
            if obj == first_obj:
                # first object; take text up until beginning of match
                obj.text = original_obj_text[:startpos]
            if obj == last_obj:
                # this is the last text object that the match is part of
                obj_with_replacement = obj
                if obj == first_obj:
                    # if first object is also last object, append text to not overwrite what we added previously
                    replace_index = len(obj.text)
                    obj.text += original_obj_text[endpos+1:]
                else:
                    # this is last but not first object, so replacement can be inserted at the beginning
                    replace_index = 0
                    obj.text = original_obj_text[endpos+1:]
            elif obj != first_obj and obj != last_obj:
                # This object is neither first nor last object, so that means all the text
                # in this object is part of the match. Delete it.
                obj.text = ""


        # Does this match have a connection to a hyperlink?
        match_hls = [ hl.attrib['Target'] for hl in match['hyperlinks'] ]

        # Find out what to replace the match with
        (replacement_text, new_hl_target) = replace_func(match['text'], match_hls, match_objects)

        # Do replacement of match text if we got a replacement
        modified = False
        if replacement_text != None:
            obj_with_replacement.text = dxsr.insert_str(obj_with_replacement.text, replacement_text, replace_index)
            modified = True # note: should probably only be true if replacement isn't equal to match

        # replace hyperlink if we received replacement
        # note: this replaces all hyperlinks to the one and same; this might not be intended in some cases.
        if new_hl_target != None and len(match_hls) > 0:
            for hl in match['hyperlinks']:
                hl.attrib['Target'] = new_hl_target
                modified = True

        self.log.debug("Objects after replacing match:\n'%s'" % dxsr.objects_to_text(match_objects, delim="|"))

        return modified

    """
    helper function that returns a replacement function from given argument
    e.g. if you give it a string, it returns a function that always returns that string
    """
    @staticmethod
    def to_replacement_func(replacer):
        if callable(replacer):
            return replacer

        if type(replacer) == str:
            # replacement is a string; define a replacement function returning this string
            replace_func = lambda *_: (replacer, None) # None == new hyperlink
            return replace_func

        raise Exception("Can not create replacement function from type", type(replacer))

    """
    Method to do substitutions, in same style as re.sub. Example:
      doc.sub( r"(\w+) World", r"\1 Earth )
    will replace e.g. "Hello World" with "Hello Earth", or "Hi World" with "Hi Earth"
    Warning: does not handle match overlaps properly, nor lookahead (?=) and lookbehind (?<=)
    """
    def sub(self, patterns, repl_pattern):
        patterns = dxsr.make_patterns(patterns)
        matches  = self.search_paragraphs(patterns)
        for match in matches:
            # use re.sub to figure out what replacement should be
            replacement_str = re.sub(match['re_object'].re, repl_pattern, match['text'])
            replace_func    = dxsr.to_replacement_func(replacement_str)
            self.replace_match(match, replace_func)
        return matches

    """
    A simple search/replace function, just like Python's str.replace().
    This can either receive one or more search patterns, or a raw string which is then translated into a pattern.
    It is possible to specify how many replacements to perform at most.

    Examples:
       doc.search_replace(".cow.", "cat")             <-- searches for raw string ".cow.", replace with "cat"
       doc.search_replace(re.compile(".cow."), "cat") <-- searches for regex ".cow.", replace with "cat"
    """
    def search_replace(self, patterns, replacement, max_replacements=0):
		# TODO: fix problem that occurs when multiple matches within same object is found - that
		# can mess things up because position in the object is bad after first replacement in the object is done.
        patterns     = dxsr.make_patterns(patterns)
        replace_func = dxsr.to_replacement_func(replacement)
        matches      = self.search_paragraphs(patterns)
        self.replace_all(matches, replace_func, max_replacements)
        return matches

    """
    Do search/replace from a dictionary mapping search regexes to replacements e.g.
        {'coolstring': 'replacedstring',
         re.compile(".coolregex.", re.IGNORECASE): 'blabla',
         re.compile("(bunny|cat)"), my_replacer_func
         }
    This allows searching for raw strings as well as regexes.
    Raw strings are treated as case sensitive.
    """
    def search_replace_dict(self, dictionary):
        # build a map of SRE_Pattern --> replacement.
        for (searchkey, replacement) in dictionary.iteritems():
            # make replacement function
            dictionary[searchkey] = dxsr.to_replacement_func(replacement)
            # translate strings to regex patterns
            if type(searchkey) == str:
                # translate string key into a pattern object and update the key
                pattern = re.compile(re.escape(searchkey), 0)
                dictionary[pattern] = dictionary.pop(searchkey)
            elif dxsr.is_sre_pattern(searchkey) == False:
                raise Exception("Bad search key type:", type(searchkey), " - must be either string or SRE_pattern")

        matches = self.search_paragraphs( dictionary.keys() )
        # replace one match at a time, using the right replacement function for each match
        for match in matches:
            match_re_pattern = match['re_object'].re
            replace_func     = dictionary[match_re_pattern]
            self.replace_match(match, replace_func)

        print "Replaced %d matches." % len(matches)
        return matches

    """
    This method is used to translate a list of string regexes to SRE patterns.
    Called by all search functions so they can receive several kinds of input, e.g.:
        search_paragraphs( ["str", "string2", re.compile(".test$")] )
        search_paragraphs("mystring")
    """
    @staticmethod
    def make_patterns(patterns, flags=0):
        if type(patterns) == list:
            ret_patterns = []
            for pattern in patterns:
                ret_patterns += dxsr.make_patterns(pattern, flags)
            return ret_patterns
        elif type(patterns) == str:
            return [re.compile(re.escape(patterns), flags)]
        elif dxsr.is_sre_pattern(patterns):
            return [patterns]
        else:
            raise Exception("got type", type(patterns), "expected str or SRE object")


    """
    Search through the document one paragraph at a time and return a list of matches.
    Searching one paragraph at a time means that a match spanning several paragraphs cannot be found.

    @param patterns, a list of regex patterns to use for searching
    """
    def search_paragraphs(self, patterns, match_modifiers=[]):
        matches = []
        patterns = dxsr.make_patterns(patterns)
        if self.verbose:
            self.log.info("Search through all paragraphs for patterns: %s" % patterns)

        for (paragraph, text_objects) in self.get_paragraphs():
            matches += self.find_matches(text_objects, patterns, match_modifiers)
        return matches


    """
    Search for matches throughout all text in the document. This means paragraphs aren't separated, but are joined with ""
    In most cases, it is better to use search_paragraphs so you don't get bad matches due to
    paragraphs being pasted right after each other.
    """
    def search_all(self, patterns, match_modifiers=[]):
        all_text_objects = self.docroot.xpath('.//w:t', namespaces=self.docroot.nsmap)
        patterns = dxsr.make_patterns(patterns)
        # Is it possible that some w:t elements do not have a text property? Maybe!
        for obj in all_text_objects:
            if obj.text == None:
                self.log.debug(" ! search_all: Removed object without text: '%s'" % obj)
                all_text_objects.remove(obj)

        return self.find_matches(all_text_objects, patterns, match_modifiers)


    """
    Return all text in the document. It is possible to choose how paragraphs are separated.
    By default, they are separated by a newline.
    """
    def all_text(self, paragraph_sep=2):
        if paragraph_sep == 0:
            # Do not use any delimiter between paragraphs; this will join paragraphs without any space/newline or anything.
            text_objects = self.docroot.xpath('.//w:t', namespaces=self.docroot.nsmap)
            return dxsr.objects_to_text(text_objects)

        if paragraph_sep != 1 and paragraph_sep != 2:
            raise Exception("paragraph_sep must be 0, 1 or 2.")

        # Text objects should have some indicator when changing from one paragraph to a new.
        # Otherwise, the two paragraphs will not have a separator, e.g. like "this is paragraph1this is paragraph2"
        if paragraph_sep == 1:
            par_prefix  = "<!paragraph!>"
            par_suffix = "</!paragraph!>\n"
        elif paragraph_sep == 2:
            par_prefix = ""
            par_suffix = "\n"

        txt = ""
        for (paragraph, text_objects) in self.get_paragraphs():
            par_txt = dxsr.objects_to_text(text_objects)
            txt += par_prefix + par_txt + par_suffix
        return txt


    """ pretty print search matches """
    @staticmethod
    def list_matches(matches):
        if len(matches) > 0:
            pprint(matches)
        print "%d matches listed." % len(matches)

    """ save search matches to file """
    def save_matches(self, filename, matches):
        # TODO: get line number(s) (in document.xml file) for each match and save so the user can open document.xml and find the matches on the right lines/columns
        matches_pretty = []
        for match in matches:
            matches_pretty.append( {
                'match': match['text'],
                'context': match['context'],
                'n_objects': len(match['objects']),
                'pattern': {
                        'regex': match['re_object'].re.pattern,
                        'flags': match['re_object'].re.flags
                    }
            } )

        with open(filename, 'w') as filehandle:
            dic = { 'infile': self.infile,
                    'n_matches': len(matches),
                    '_matches': matches_pretty }
            json.dump(dic, filehandle, indent=2, sort_keys=True)

        print "Wrote matches via json.dump to file", filename





"""
    What is a match modifier? It is a function which is called on each found search match.
    The 'match' can then be shrunk or grown to cover more or less text.
    This is to add functionality that is hard to realize with just regexes.
    This is meant for advanced use on complex problems.

    Example: a regex is used to match URLs, but naturally it can't match URLs with spaces in them.
    So every time we match a URL, we use a match modifier function to look beyond the matched text to
    see if what comes next should be considered part of the URL as well. In that case, we can "grow"
    the match so it covers the next few words ahead.
"""



# TODO
if __name__ == "__main__":
    print "not implemented"
    exit(1)

