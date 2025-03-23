import itertools
import csv
import os
import string
import time
from typing import Generator, List, Optional
from msoffcrypto.format.ooxml import OOXMLFile
import msoffcrypto


OUTPUT_FOLDER = "checked_files"

 # Helper: produce all case variations for a string
def case_variations(s: str):
    variants = [(ch.lower(), ch.upper()) if ch.isalpha() else (ch,) for ch in s]
    return (''.join(candidate) for candidate in itertools.product(*variants))

def generate_passwords(
    prefixes: Optional[List[str]] = None,
    suffixes: Optional[List[str]] = None,
    max_length: int = 10
) -> Generator[str, None, None]:
    """
    Generates all possible passwords based on the provided prefixes, suffixes, and maximum length.
    The first character of the password is always a letter if no prefix is provided. Therefore, if
    you want to generate passwords that start with a specific letter, you must provide a prefix.

    :param prefixes: List of prefixes to use (default is None).
    :param suffixes: List of suffixes to use (default is None).
    :param max_length: Maximum length of the passwords to generate (default is 10).
    :return: A generator yielding password combinations.
    """

    # Letters for the first character if no prefix is provided
    letters = string.ascii_letters

    # Valid ASCII characters (excluding whitespace and control characters)
    valid_chars = letters + string.digits + string.punctuation

    # Prepare prefix variations (or just an empty string if no prefixes)
    prefixes = prefixes or [""]
    prefix_variations = set()
    for prefix in prefixes:
        prefix_variations.update(case_variations(prefix))

    # Prepare suffix variations (or just an empty string if no suffixes)
    suffixes = suffixes or [""]
    suffix_variations = set()
    for suffix in suffixes:
        suffix_variations.update(case_variations(suffix))

    # Set to track unique passwords
    seen_passwords = set()

    # Generate passwords by combining prefixes, generated body, and suffixes
    for prefix in prefix_variations:
        for suffix in suffix_variations:
            total_fixed_length = len(prefix) + len(suffix)

            # If prefix + suffix fills max_length already
            if total_fixed_length > max_length:
                continue

            remaining_length = max_length - total_fixed_length

            # If there's no space for a body, just yield prefix + suffix
            if remaining_length == 0:
                full_password = prefix + suffix
                if full_password not in seen_passwords:
                    seen_passwords.add(full_password)
                    yield full_password
                continue

            # Character pools: first char special if no prefix
            first_chars = letters if prefix == "" else valid_chars

            # Generate all possible combinations for the body of the password
            for length in range(1, remaining_length + 1):
                if length == 1:
                    char_pool = [first_chars]
                else:
                    char_pool = [first_chars] + [valid_chars] * (length - 1)

                for body_chars in itertools.product(*char_pool):
                    body = ''.join(body_chars)

                    # Generate all case variations for the body
                    for body_variation in case_variations(body):
                        full_password = prefix + body_variation + suffix
                        if full_password not in seen_passwords:
                            seen_passwords.add(full_password)
                            yield full_password

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
        for password in itertools.product(valid_chars, repeat=length):
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
                print(f"\nChecked {num_checked} passwords, skipped {num_skipped} passwords in {elapsed_time:.2f} seconds ({(num_checked+num_skipped)/elapsed_time} pwds/sec).")
                return

            except msoffcrypto.exceptions.DecryptionError:
                # Even printing the error message can be a security risk and slow things down, so
                # do nothing
                pass
                # print(f"DecryptionError: Failed to verify password '{password}'. {e}")
            except KeyboardInterrupt:
                print("KeyboardInterrupt: Stopping password check because CTRL + C was pressed.")
                # This break is required to stop the program! Otherwise, it will continue to
                # check the next password
                break
            except Exception as e:
                print(f"An unknown error occurred: {e}")

            # Write the password to the checked file
            checked_file.write(f"{password}\n")

    print("None of the passwords worked.")
    end_time = time.perf_counter()
    elapsed_time = end_time - start_time
    print(f"\nChecked {num_checked} passwords, skipped {num_skipped} passwords in {elapsed_time:.2f} seconds ({(num_checked+num_skipped)/elapsed_time} pwds/sec).")

# Example usage
if __name__ == "__main__":
    # User-provided base passwords
    # BASE_PASSWORDS = ["unccu", "baritone", "admin"]
    BASE_PASSWORDS = ["baritone"]

    # Optional prefixes, suffixes, and numbers
    # PREFIXES = ["",]
    # SUFFIXES = ["", "2021", "202!", "131", "13!", "1819", "1819!", "!31", "!3!", "131!", "!131", "!131!", "1E!", "!31!", "1331", "!331", "1331!"]
    # numbers = ["",]

    # Path to the Excel file
    EXCEL_FILE_PATH = "ANTV2.xlsx"  # Replace with your Excel file path

    # 1. Test if the file is encrypted
    if not file_is_encrypted(EXCEL_FILE_PATH):
        print("The file is not encrypted.")
        exit()

    print(f"File is encrypted: {EXCEL_FILE_PATH}")

    # 2. Now that we have confirmed the file is encrypted, we can generate passwords
    # By specifying only prefixes, we can generate all passwords with the given prefixes
    # Therefore, to generate all passwords starting with a letter, add the letter as an element of
    # the BASE_PASSWORDS list
    all_passwords = generate_passwords(BASE_PASSWORDS, [], max_length=11)

    # 3. Loop through the password list and test all of the passwords until we find the right one
    test_passwords(excel_file=EXCEL_FILE_PATH, password_list=all_passwords)
