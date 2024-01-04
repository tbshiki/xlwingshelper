def alpha_to_num(letters):
    """
    Converts a string of letters to its corresponding Excel column number.
    For example, 'A' becomes 1, 'Z' becomes 26, 'AA' becomes 27, 'BC' becomes 55, etc.

    Args:
        letters (str): The string of letters to be converted.

    Returns:
        int: The corresponding Excel column number.

    Examples:
        >>> alphabet_position('A')
        1
        >>> alphabet_position('Z')
        26
        >>> alphabet_position('AA')
        27
        >>> alphabet_position('BC')
        55
    """
    result = 0
    for char in letters.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


def num_to_alpha(num, uppercase=True):
    """
    Converts a given integer to an alphabetical string representation, similar to Excel column names.
    For example, 1 converts to 'A', 27 to 'AA', and 52 to 'AZ'.

    Args:
        num (int): The integer to be converted. Must be 1 or higher.
        uppercase (bool): If True (default), the output is in uppercase letters.
                          If False, the output is in lowercase letters.

    Returns:
        str: The corresponding alphabetical string representation of the given number.
             For example, 1 becomes 'A', 27 becomes 'AA'.

    Raises:
        ValueError: If `num` is less than 1.

    Examples:
        >>> num_to_alpha(1)
        'A'
        >>> num_to_alpha(27)
        'AA'
        >>> num_to_alpha(702, uppercase=False)
        'zz'
    """
    if num < 1:
        raise ValueError("Number must be 1 or higher")

    result = ""
    while num > 0:
        num, remainder = divmod(num - 1, 26)
        if uppercase:
            result = chr(65 + remainder) + result
        else:
            result = chr(97 + remainder) + result
    return result
