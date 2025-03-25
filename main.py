import itertools
import argparse
import csv
import os
import string
import sys
import time
from typing import Generator, List, Optional
from msoffcrypto.format.ooxml import OOXMLFile
import msoffcrypto


OUTPUT_FOLDER = "checked_files"

 # Helper: produce all case variations for a string
def case_variations(s: str):
    """
    Generate all case variations for a given string.
    :param s: The input string.
    :return: A generator yielding all case variations of the string.
    """
    variants = [(ch.lower(), ch.upper()) if ch.isalpha() else (ch,) for ch in s]
    return (''.join(candidate) for candidate in itertools.product(*variants))

def positive_int(value: str) -> int:
    """
    Argument type for positive integers. Used by the min_length and max_length parameters.

    :param value: The input value to check.
    :return: The positive integer value.
    :raises argparse.ArgumentTypeError: If the value is not a positive integer.
    """
    try:
        ivalue = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"{value} is not a valid integer") from exc
    if ivalue <= 0:
        raise argparse.ArgumentTypeError("Value must be a positive integer greater than 0")
    return ivalue

def process_arg(arg_str: str) -> list[str]:
    """
    Processes a comma-separated argument string:
      - Converts the entire string to lowercase.
      - Splits it into a list on commas.
      - Strips whitespace from each element.
      - Deduplicates the list while preserving order.
    """
    # Convert to lowercase and split
    items = [item.strip() for item in arg_str.casefold().split(",") if item.strip()]
    # Remove duplicates while preserving order
    return list(dict.fromkeys(items))


def parse_args():
    """
    Parses command-line arguments. This sets the help text that is provided to the console.
    :return: Parsed arguments.
    """
    parser = argparse.ArgumentParser(
        description="Generates and checks passwords to unlock a password-protected Excel workbook."
    )
    # This is a positional argument and not an optional command line argument
    # This means that when parser.parse_args() is called, the user must provide a file
    # Otherwise, the script will not run and an error will be shown
    parser.add_argument(
        "file",
        type=str,
        help="The path to the password-protected Microsoft Office file. (e.g., file.xlsx)"
    )
    parser.add_argument(
        "--prefixes",
        type=str,
        default="",
        help="A string of comma-separated prefixes (e.g., 'Admin,root,User')."
    )
    parser.add_argument(
        "--suffixes",
        type=str,
        default="",
        help="A string of comma-separated suffixes (e.g., '123,XYZ')."
    )
    parser.add_argument(
        "--max_length",
        type=positive_int,
        default=10,
        help="Max length of the password (positive integer greater than 0)."
    )

    # If no arguments are provided, print the help message
    if len(sys.argv) == 1:
        parser.print_help()
        sys.exit(0)

    return parser.parse_args()

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

    # Define the character set: all printable ASCII characters except whitespace and control chars
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
    # To make sure that we do not miss the correct password, we create a success file
    # that will contain the correct password if found
    success_file_path = os.path.join(OUTPUT_FOLDER, f"{file_name}_success.csv")

    num_checked = 0
    num_skipped = 0

    # Used to determine how long it takes to check all of the passwords
    start_time = time.perf_counter()

    # This opens both the Excel file we will be attempting to find the password for
    # and the file we will be writing the checked passwords to
    with open(excel_file, "rb") as f, \
        open(checked_file_path, mode='a', newline='', encoding='utf-8') as checked_file, \
        open(success_file_path, mode='a', newline='', encoding='utf-8') as success_file:
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

                # If the password is correct, write it to the success file and print the result
                success_file.write(f"{password}\n")
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

def main():
    """
    Main entry point for this script. Using main() prevents the script from running when imported
    and also prevents cluttering the global namespace.
    """
    args = parse_args()

    # Set the prefixes and suffixes to the user-provided values or defaults
    prefixes = process_arg(args.prefixes) if args.prefixes else ["unccu"]
    suffixes = process_arg(args.suffixes) if args.suffixes else []

    office_file = args.file if args.file else ""

    max_length = args.max_length

    print("Prefixes:", prefixes)
    print("Suffixes:", suffixes)
    print("Max Length:", max_length)

    # 1. Test if the file is encrypted
    if not file_is_encrypted(office_file):
        print(f"The file is not encrypted: {office_file}")
        exit()

    print(f"File is encrypted: {office_file}")

    # 2. Now that we have confirmed the file is encrypted, we can generate passwords
    # By specifying only prefixes, we can generate all passwords with the given prefixes
    # Therefore, to generate all passwords starting with a letter, add the letter as a prefix
    all_passwords = generate_passwords(prefixes, suffixes, max_length=max_length)

    # 3. Loop through the password list and test all of the passwords until we find the right one
    test_passwords(excel_file=office_file, password_list=all_passwords)

# Example usage
if __name__ == "__main__":
    main()
