import re

def _substr_after(s, delim):
    return s.partition(delim)[2]

passive_package_regex = re.compile('(^.*\s+[0-9]+)\s+[0-9]+(M|m)etric$')

def translate_fp(fp_string):
    """
    Translate a footprint to a simpler human-readable format

    The goal is to make the BOM as clean and readable as possible. Note
    that the translated footprints are l
    still keeping the footprints unique enough that they can be used to
    correctly group parts based on them
    """

    if not fp_string:
        return ""

    if not isinstance(fp_string, str):
        fp_string = str(fp_string)
    result = fp_string

    # Try to remove the library prefix
    lib_prefix_removed = _substr_after(fp_string, ':')
    if lib_prefix_removed:
        result=lib_prefix_removed

    # Underscore to space for better readability
    result = result.replace('_', ' ').strip()

    match = passive_package_regex.match(result)
    if match:
        result = match.group(1)


    return result


