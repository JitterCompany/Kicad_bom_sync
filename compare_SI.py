import re
numeric_regex = re.compile('(^[0-9.]+)\s*([^0-9]?)')


def _to_numeric(SI_str):

    SI_multipliers = {
            'y': 1e-24,
            'z': 1e-21,
            'a': 1e-18,
            'f': 1e-15,
            'p': 1e-12,
            'n': 1e-9,
            'µ': 1e-6,
            'u': 1e-6,
            'm': 1e-3,
            'c': 1e-2,
            'd': 1e-1,
            'h': 1e2,
            'k': 1e3,
            'K': 1e3,
            'M': 1e6,
            'G': 1e9,
            'T': 1e12,
            'P': 1e15,
            'E': 1e18,
            'Z': 1e21,
            'Y': 1e24,
            }

    match = numeric_regex.match(SI_str)
    if not match:
        return None
    
    num = match.group(1)
    suffix = match.group(2)
    if suffix in SI_multipliers:
        mult = SI_multipliers[suffix]
    else:
        mult = 1

    return (float(num) * mult)


def compare_SI(a,b):

    # Try to compare numerically by applying SI suffix
    a_num = _to_numeric(a)
    b_num = _to_numeric(b)
    if (a_num is not None) and (b_num is not None):

        result = a_num-b_num
        #print("Numeric: {} < {}".format(a,b), (True if result < 0 else False ))
        return result

    # Fallback to default string sorting (alphabetical order)
    if a == b:
        return 0

    result = (a < b)
    #print("Fallback: {} < {}".format(a,b), result)

    return -1 if result else 1


