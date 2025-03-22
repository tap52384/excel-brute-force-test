import openpyxl
from openpyxl.utils.exceptions import InvalidFileException
from itertools import product

from itertools import product, chain

def case_variations(word):
    """
    Generates all possible case variations of a given word.
    For example, 'abc' -> ['abc', 'Abc', 'aBc', 'ABc', 'abC', 'AbC', 'aBC', 'ABC']
    """
    if not word:
        return [""]
    return map("".join, product(*((char.lower(), char.upper()) for char in word)))

def generate_passwords(base_passwords, prefixes=None, suffixes=None):
    """
    Generates a list of passwords based on combinations of base passwords, prefixes, suffixes,
    and their case variations.

    :param base_passwords: List of base passwords.
    :param prefixes: List of prefixes to prepend to the base passwords (optional).
    :param suffixes: List of suffixes to append to the base passwords (optional).
    :return: A generator yielding password combinations.
    """
    prefixes = prefixes or [""]
    suffixes = suffixes or [""]

    # Generate case variations for prefixes, base passwords, and suffixes
    prefixes_variations = chain.from_iterable(case_variations(prefix) for prefix in prefixes)
    base_variations = chain.from_iterable(case_variations(base) for base in base_passwords)
    suffixes_variations = chain.from_iterable(case_variations(suffix) for suffix in suffixes)

    # Generate all combinations of prefixes, base passwords, and suffixes
    for prefix, base, suffix in product(prefixes_variations, base_variations, suffixes_variations):
        yield f"{prefix}{base}{suffix}"

def test_passwords(excel_file, password_generator):
    """
    Attempts to open a password-protected Excel file using a generator of passwords.

    :param excel_file: Path to the Excel file.
    :param password_generator: A generator yielding passwords to test.
    """
    for password in password_generator:
        try:
            print(f"Trying password: {password}")
            workbook = openpyxl.load_workbook(excel_file, read_only=True, password=password)
            print(f"Success! The correct password is: '{password}'")
            return
        except InvalidFileException:
            print(f"Invalid password: {password}")
        except Exception as e:
            print(f"An error occurred: {e}")

    print("None of the passwords worked.")

# Example usage
if __name__ == "__main__":
    # User-provided base passwords
    base_passwords = ["unccu", "baritone", "admin"]

    # Optional prefixes, suffixes, and numbers
    prefixes = ["",]
    suffixes = ["", "2021", "202!", "131", "13!", "1819", "1819!"]
    # numbers = ["",]

    # Path to the Excel file
    excel_file_path = "/Volumes/Seagate Backup Plus Drive/MBP/ANTV2.xlsx"  # Replace with your Excel file path

    # Generate passwords dynamically
    password_generator = generate_passwords(base_passwords, prefixes, suffixes)

    # Test the generated passwords
    test_passwords(excel_file_path, password_generator)
