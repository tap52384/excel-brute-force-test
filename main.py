from itertools import product, chain
import csv
import os
import string
import time
from typing import Generator
from msoffcrypto.format.ooxml import OOXMLFile
import msoffcrypto


OUTPUT_FOLDER = "checked_files"

def case_variations(word):
    """
    Generates all possible case variations of a given word.
    For example, 'abc' -> ['abc', 'Abc', 'aBc', 'ABc', 'abC', 'AbC', 'aBC', 'ABC']
    """
    if not word:
        return [""]
    return map("".join, product(*((char.lower(), char.upper()) for char in word)))

def generate_passwords(password_bases: list[str], prefixes: list[str], suffixes: list[str]) -> Generator[str, None, None]:
    """
    Generates a list of passwords based on combinations of base passwords, prefixes, suffixes,
    and their case variations.

    :param base: List of base passwords (strings).
    :param prefixes: List of prefixes to prepend to the base passwords
        (optional, defaults to an empty list).
    :param suffixes: List of suffixes to append to the base passwords
        (optional, defaults to an empty list).
    :return: A generator yielding password combinations.

    :param base: List of base passwords.
    :param prefixes: List of prefixes to prepend to the base passwords (optional).
    :param suffixes: List of suffixes to append to the base passwords (optional).
    :return: A generator yielding password combinations.
    """
    prefixes = prefixes or [""]
    suffixes = suffixes or [""]
    # The set() function is used to remove duplicates
    prefixes_variations = set(chain.from_iterable(
        case_variations(prefix) for prefix in prefixes))
    base_variations = set(chain.from_iterable(
        case_variations(base) for base in password_bases))
    suffixes_variations = set(chain.from_iterable(
        case_variations(suffix) for suffix in suffixes))

    # Generate all combinations of prefixes, base passwords, and suffixes
    for prefix, base, suffix in product(prefixes_variations, base_variations, suffixes_variations):
        yield f"{prefix}{base}{suffix}"

def generate_all_passwords(max_length: int = 10) -> Generator[str, None, None]:
    """
    Generates all possible passwords starting with a letter and up to a given length.
    Includes all ASCII printable characters except whitespace and control characters.

    :param max_length: Maximum length of the passwords to generate (default is 10).
    :return: A generator yielding password combinations.
    """

    # Define the character set: all printable ASCII characters except whitespace and control characters
    valid_chars = string.ascii_letters + string.digits + string.punctuation

    # Generate passwords starting with a letter and up to max_length characters
    for length in range(1, max_length + 1):
        for password in product(valid_chars, repeat=length):
            if password[0].isalpha():  # Ensure the password starts with a letter
                yield ''.join(password)



def file_is_encrypted(file_path:str) -> bool:
    """
    Check if the file is encrypted.
    """

    # Stop if the file does not exist
    if not file_path:
        raise FileNotFoundError("Excel file path not provided.")

    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Excel file not found: {file_path}")

    # To check if the file is encrypted from the command line, use the -t flag:
    # msoffcrypto-tool file.xlsx --test -v
    # Opening a file in Python as "rb" means read-only for a binary file
    with open(file_path, "rb") as file:
        try:
            # Attempt to decrypt the file
            officefile = OOXMLFile(file)
            return officefile.is_encrypted()

        except msoffcrypto.exceptions.DecryptionError:
            print("DecryptionError: Failed to confirm if the file is encrypted.")
        except msoffcrypto.exceptions.FileFormatError:
            print("FileFormatError: Failed to confirm if the file is encrypted.")
        except msoffcrypto.exceptions.EncryptionError:
            print("EncryptionError: Failed to confirm if the file is encrypted.")


        # If the file could not be confirmed as encrypted, return False to prevent
        # unnecessary attempts to test the password
        return False

def get_checked_passwords(file_path:str) -> set[str]:
    """
    Get a list of passwords already checked for the given file to save time.
    """

    file_name = os.path.basename(file_path)
    checked_file_path = os.path.join(OUTPUT_FOLDER, f"{file_name}.csv")

    # Create the output folder if it doesn't exist
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)
        return set()

    checked_passwords = set()
    if os.path.exists(checked_file_path):
        with open(checked_file_path, mode='r', encoding='utf-8') as file:
            reader = csv.reader(file)
            for row in reader:
                checked_passwords.add(row[0])

    return checked_passwords


def test_passwords(excel_file, password_list: Generator[str]):
    """
    Attempts to open a password-protected Excel file using a list of passwords.

    :param excel_file: Path to the Excel file.
    :param password_list: A list of passwords to test.
    """

    # Get the list of passwords already checked; this can save time by not checking passwords again
    checked_passwords = get_checked_passwords(excel_file)

    file_name = os.path.basename(excel_file)
    checked_file_path = os.path.join(OUTPUT_FOLDER, f"{file_name}.csv")

    num_checked = 0
    num_skipped = 0

    # Used to determine how long it takes to check all of the passwords
    start_time = time.perf_counter()

    # This opens both the Excel file we will be attempting to find the password for
    # and the file we will be writing the checked passwords to
    with open(excel_file, "rb") as f, \
        open(checked_file_path, mode='a', newline='', encoding='utf-8') as checked_file:
        file = OOXMLFile(f)

        for password in password_list:
            try:

                # Skip the password if it has already been checked
                if password in checked_passwords:
                    num_skipped += 1
                    continue

                # This will attempt to verify the password
                # https://msoffcrypto-tool.readthedocs.io/en/latest/index.html#id1
                num_checked += 1
                file.load_key(password=password, verify_password=True)

                end_time = time.perf_counter()
                elapsed_time = end_time - start_time

                print(f"***SUCCESS*** The correct password is: '{password}'")
                print(f"\nChecked {num_checked} passwords, skipped {num_skipped} passwords in {elapsed_time:.2f} seconds.")
                return

            except msoffcrypto.exceptions.DecryptionError as e:
                print(f"DecryptionError: Failed to verify password '{password}'. {e}")
            except Exception as e:
                print(f"An unknown error occurred: {e}")

            # Write the password to the checked file
            checked_file.write(f"{password}\n")

    print("None of the passwords worked.")
    end_time = time.perf_counter()
    elapsed_time = end_time - start_time
    print(f"\nChecked {num_checked} passwords, skipped {num_skipped} passwords in {elapsed_time:.2f} seconds.")

# Example usage
if __name__ == "__main__":
    # User-provided base passwords
    BASE_PASSWORDS = ["unccu", "baritone", "admin"]

    # Optional prefixes, suffixes, and numbers
    PREFIXES = ["",]
    SUFFIXES = ["", "2021", "202!", "131", "13!", "1819", "1819!", "!31", "!3!", "131!", "!131", "!131!", "1E!", "!31!", "1331", "!331", "1331!"]
    # numbers = ["",]

    # Path to the Excel file
    EXCEL_FILE_PATH = "ANTV2.xlsx"  # Replace with your Excel file path

    # 1. Test if the file is encrypted
    if not file_is_encrypted(EXCEL_FILE_PATH):
        print("The file is not encrypted.")
        exit()

    print(f"File is encrypted: {EXCEL_FILE_PATH}")

    # 2. Now that we have confirmed the file is encrypted, we can generate passwords
    all_passwords = generate_passwords(BASE_PASSWORDS, prefixes=PREFIXES, suffixes=SUFFIXES)

    # 3. Attempt to generate all passwords starting with a letter and up to a given length
    comprehensive_passwords = generate_all_passwords(max_length=10)

    # 3. Loop through the password list and test all of the passwords until we find the right one
    test_passwords(excel_file=EXCEL_FILE_PATH, password_list=comprehensive_passwords)
