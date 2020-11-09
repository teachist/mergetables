import os


def verifyIDcardValid(ID):
    """verifyIDcardValid(ID)

    Check ID card whethear it valid
    The checksum is 7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2
    Using the first 17 number to multiply the corresponding checksum and
    add them all together, then module it to 11 to get the remainder, the
    remainders are 1 throught 10, using the remainder as the list position
    to retrive the value. Finally, using the value to match the 18th number
    with it, if it matches, then the ID is valide, or not.
    """
    if len(ID.strip()) != 18:
        return "NOT A ID"
    else:
        ID_check = ID[17]
        W = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
        ID_CHECK = ['1', '0', 'X', '9', '8', '7', '6', '5', '4', '3', '2']
        ID_aXw = 0
        for i in range(len(W)):
            ID_aXw = ID_aXw + int(ID[i]) * W[i]
        ID_Check = ID_aXw % 11
        if ID_check != ID_CHECK[ID_Check]:
            return 'INVALID'
        else:
            return ID


def get_filelist(folder_name):
    """get_filelist(folder_name)

    Return all the files include the given folder name after sorted.

    Folder: the folder include the spreadsheet files the function are
    going to process.
    Also it will do some clearnings, like the hidden file for specified
    operating system.
    """
    # print(self._folder_name)
    results = []
    for home, dirs, files in os.walk(folder_name):
        for filename in files:
            _file = os.path.join(home, filename)
            if filename == '.DS_Store':
                os.remove(_file)
            # print(filename)
            results.append(_file)

    return sorted(results)


def _chinese_str_len(csl):
    try:
        built_in_len = len(csl)
        utf8_len = len(csl.encode('utf-8'))
        return (utf8_len-built_in_len) // 2 + built_in_len
    except:
        return None
    return None


def print_data(rows):
    column_lenght = len(list(list(rows)[0]))
    print("-" * 20 * (column_lenght+1))
    print('行号', end='\t')
    print('| ', end='')
    for x in range(column_lenght):
        text = f'列{x+1:02}'
        space = ' ' * (20 - _chinese_str_len(text))
        print(text+space + '|', end='')
    print()
    print("-" * 20 * (column_lenght+1))
    for index, row in enumerate(rows):
        if index > 15:
            return
        print(f'{index+1:03}', end='\t')
        print('| ', end='')
        for value in row:
            value = value.strip()
            space = ' ' * (20 - _chinese_str_len(value))
            print(value+space + '|', end='')
        print()
