#!/usr/bin/env python
"""Convert ICS bank statements from PDF to CSV, TSV or JSON.

Annoyingly, ICS (icscards.nl) only allows customers to download bank statements in PDF, 
not in computer readable formats. This Python script translates PDF to CVS.

This script heavily relies on the exact layout of the ICS bank statement.
If that changes, this script will break.
It has been tested on bank statements from the last 10 years (2016-2025).

Requirements:
- pdfplumber

Homepage: https://www.github.com/macfreek/icscards-parser/

This script is copyright 2025-2026 Freek Dijkstra
License: MIT-license

In short: You can use this script in any way you like, but there is no support or warranty.
I appreciate that you give credit if you use (large parts of) this script.
"""
from csv import DictWriter
from dataclasses import dataclass, fields, asdict
from datetime import date, timedelta
import json
import logging
from pathlib import Path
import re
import sys

logger = logging.getLogger(__name__)
if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

try:
    import pdfplumber
except ImportError:
    logger.critical("pdfplumber not installed. Run `pip install pdfplumber`.")
    if __name__ == '__main__':
        sys.exit(1)
    raise


DEFAULT_OUTPUT = "TSV"  # CSV or TSV or JSON
DESTINATION_DIR = None  # The export directory. Use None for the same as the source directory.

@dataclass
class Transaction:
    """A single line on the bank statement."""
    transaction_date: date
    booking_date: date
    description: str
    location: str
    country: str
    amount: float
    foreign_amount: float | None = None
    foreign_currency: str | None = None
    foreign_exchange_rate: float | None = None
    card_last_numbers: str = ""
    card_owner: str = ""

    def as_dict(self)-> dict:
        """Return the transaction as a dict."""
        return asdict(self)

    def as_json_dict(self)-> dict:
        """Return the transaction as a dict.
        null values are skipped (empty value are retained),
        dates are converted to ISO8601 string."""
        result = {}
        for key, value in asdict(self).items():
            if value is None:
                continue
            if not isinstance(value, (str, float, int)):
                value = str(value)
            result[key] = value
        return result


@dataclass
class BankStatement:
    """Represents one bank statement, containing the transactions of one month."""
    date: date
    customer_number: str
    serial_number: int
    previous_balance: float
    total_received_payment: float
    total_new_expenses: float
    new_balance: float
    transactions: list[Transaction]

    def properties(self) -> dict:
        """Return a dict with the above defined attributes, except the transactions"""
        props = self.as_dict()
        del props['transactions']
        return props

    def as_dataframe(self) -> "pandas.DataFrame":
        """Return the transactions as a pandas DataFrame"""
        import pandas
        raise NotImplemented("exporting to dataframe is not yet implemented")

    def as_dict(self)-> dict:
        """Return the statement as dict, with the transactions as list of dicts"""
        return asdict(self)

    def as_json_dict(self)-> dict:
        """Return the statement as dict, with the transactions as list of dicts.
        null values are skipped (empty value are retained),
        dates are converted to ISO8601 string."""
        result = {}
        for key, value in asdict(self).items():
            if value is None:
                continue
            if key == "transactions":
                value = [item.as_json_dict() for item in self.transactions]
            elif not isinstance(value, (str, float, int)):
                value = str(value)
            result[key] = value
        return result

    def save_as_csv(self, filename: Path | str, dialect='excel'):
        """Saves the transactions as comma separated values (CSV) file"""
        if isinstance(filename, str):
            filename = Path(filename)
        with filename.open('w', newline='') as fp:
            extras = {'month': f"{self.date.year:04}-{self.date.month:02}"}
            fieldnames = list(extras.keys())
            fieldnames.extend(field.name for field in fields(Transaction))
            writer = DictWriter(fp, dialect=dialect, fieldnames=fieldnames)
            writer.writeheader()
            for transaction in self.transactions:
                writer.writerow(extras | transaction.as_json_dict())

    def save_as_tsv(self, filename: Path | str):
        """Saves the transactions as tab separated values (TSV) file"""
        self.save_as_csv(filename, dialect='excel-tab')

    def save_as_json(self, filename: Path | str):
        """Saves the whole statement as JSON file"""
        if isinstance(filename, str):
            filename = Path(filename)
        with filename.open('w') as fp:
            json.dump(self.as_json_dict(), fp, indent=4)

    def print(self):
        """Debug function: print this statement in human readable form"""
        for key, value in self.properties().items():
            print(f"{key:25} {value!s:>15}")
        print(82*'-')
        print(f"{'Date':10}  {'Description':25}  {'Location':13}  {'CC':2}  {'Amount':9}  {'Foreign'}")
        print(82*'-')
        for ta in self.transactions:
            if ta.foreign_amount is None:
                print(f"{ta.transaction_date!s}  {ta.description:<25}  {ta.location:<13}  {ta.country:<2}  {ta.amount:9.2f}")
            else:
                print(f"{ta.transaction_date!s}  {ta.description:<25}  {ta.location:<13}  {ta.country:<2}  {ta.amount:9.2f}  {ta.foreign_amount:9.2f} {ta.foreign_currency}")
        print(82*'=')


class ParseError(ValueError):
    pass


def print_words(words: list[dict]):
    """Debug function. Print found words"""
    for word in words:
        print(f"{round(word['top'],3)}  {round(word['x0'],3)}-{round(word['x1'],3)}  {word['text']}")


def print_lines(lines: list[list[dict]]):
    """Debug function. Print found words, by line"""
    first_line = True
    for line in lines:
        print_words(line)
        if first_line:
            print(40*'-')


def parse_date(date_str: str, year_hint: date | None = None) -> date:
    """Parse a date. For partial dates (day+month), provide a year_hint.
    The result will be a date before year_hint."""
    year_hint_tolerance = 40  # allow 40 days after year_hint.
    def parse_month(month_str: str) -> int:
        MONTHS = {'jan': 1, 'feb': 2, 'maa': 3, 'apr': 4, 'mei': 5, 'jun': 6,
            'jul': 7, 'aug': 8, 'sep': 9, 'okt': 10, 'nov': 11, 'dec': 12,
            'mar': 3, 'mrt': 3, 'may': 5, 'oct': 10}
        try:
            return MONTHS[month_str[:3].lower()]
        except IndexError:
            raise ParseError(f"Unknown month {month_str!r}.")
    date_items = date_str.split()
    if len(date_items) not in (2, 3):
        raise ParseError(f"Invalid date {date_str!r}.")
    try:
        day = int(date_items[0])
    except ValueError:
        raise ParseError(f"Invalid date {date_str!r}.") from None
    month = parse_month(date_items[1])
    if len(date_items) > 2:
        try:
            year = int(date_items[2])
        except ValueError:
            raise ParseError(f"Invalid date {date_str!r}.") from None
    elif year_hint is None:
        raise ParseError(f"Impartial date {date_str!r}.")
    else:
        year = year_hint.year
        if date(year, month, day) > year_hint + timedelta(days=year_hint_tolerance):
            year -= 1
    return date(year, month, day)


def parse_amount(amount: str, BijAf: str = "Bij", *args) -> float:
    """Parse an amount to a float.
    e.g. '€ 1.234,56', 'Af' becomes -1234.56"""
    if args:
        raise ParseError(f"Too many words in amount: {amount} {BijAf} {' '.join(args)}")
    intl_amount = amount.strip('€ ').replace('.', '').replace(',', '.')
    try:
        num_amount = float(intl_amount)
    except ValueError:
        raise ParseError(f"{amount!r} can't be parsed as a numeric amount") from None
    BijAf = BijAf.strip()
    if BijAf not in ('Af', 'Bij'):
        raise ParseError(f"Expected 'Bij' or 'Af', found {BijAf!r}") from None
    if BijAf == 'Af' and num_amount > 0:
        num_amount = -num_amount
    return num_amount


def group_by_lines(words: list[dict], line_tolerance=3) -> list[list[dict]]:
    """Groups chars or words together that belong to the same transaction line.
    Input: list of all words (each a dict from pdfplumber).
    Output: a list of lines, with each line a list of words."""
    if not words:
        return []
    line_top = words[0]['top']
    lines = []
    line = []
    def new_line():
        nonlocal line, lines
        if line:
            lines.append(line)
        line = []
    for word in words:
        if word.get('upright', True) != True:
            raise ParseError(f"Word {word['text']} is not upright")
        if word.get('direction', 'ltr') != 'ltr':
            raise ParseError(f"Word {word['text']} is not left-to-right")
        # A new line starts if the line height jumps 3 or more pixels.
        if word['top'] >= line_top + line_tolerance:
            new_line()
            line_top = word['top']
        line.append(word)
    new_line()
    return lines


def header_words_to_dict(words: list[dict]) -> dict[str, list[str]]:
    """Correlate keys and values in the header to a dict.
    Input: list of all words (each a dict from pdfplumber).
    Output: dict of key -> value(s), with each value a list of string.
    Most values are just one string, but monetary amount have two strings
    (e.g. ["€120,00", "Bij"]) and also page number are two str (e.g. ["1", "van 2"])."""
    if not words:
        return {}
    lines = group_by_lines(words, line_tolerance=3)
    if len(lines) % 2:
        # not an even number of lines
        # print_lines(lines)
        raise ParseError(f"Expected 4 lines in header. Found {len(lines)} lines.")
    values: dict[str, list[str]] = {}
    x_tolerance = 2
    while lines:
        header_line = lines.pop(0)
        key_positions: list[tuple[float, str]]  # x0 (horizontal) position, name of the header
        key_positions = [(word['x0']-x_tolerance, word['text'].lower()) for word in header_line]
        key_positions.sort(reverse=True)
        for key_x0, key_text in key_positions:
            if key_text in values:
                logger.error(f"Duplicate header {key_text!r}")
            else:
                values[key_text] = []
        values_line = lines.pop(0)
        for word in values_line:
            keyword = None
            for key_x0, key_text in key_positions:
                if word['x0'] >= key_x0:
                    keyword = key_text
                    break
            if not keyword:
                keyword = key_positions[-1][1]
            values[keyword].append(word['text'])
    return values


def get_boundaries_from_first_line(words: list[dict]):
    """Parse the first line, and return column boundaries."""

    # Many columns in the header are split into multiple subcolumns.
    # relative_boundaries specifies these boundaries relative to the headers.
    relative_boundaries = {
        'Datum transactie': [-1],  # 61.0 -> 60.0
        'Datum boeking': [-2],  # 105.0 -> 60.0
        'Omschrijving': [-4, 122, 209],  # 154.0 -> 150.0, 276.0, 363.0
        'Bedrag in vreemde valuta': [-3, 43],  # 402.0 -> 399.0, 445.0
        "Bedrag in euro's": [-3, 56],  # 479.0 -> 476.0, 535.0
    }

    # join words spanning two lines (with same x0, just underneath each other)
    # headers contains the left boundary as key and title as value
    headers: dict[float, str] = {}
    for word in words:
        if word['x0'] not in headers:
            headers[word['x0']] = word['text']
        else:
            headers[word['x0']] = headers[word['x0']] + " " + word['text']

    # check if we have the headers as expected
    expected_first_line = list(relative_boundaries.keys())
    if list(headers.values()) != expected_first_line:
        raise ParseError(f"Unexpected headers {list(headers.values())}. Expected {expected_first_line}.")

    boundaries = []
    for x0, header_title in headers.items():
        for relative_boundary in relative_boundaries[header_title]:
            boundaries.append(x0 + relative_boundary)

    expected_boundaries = [60.0, 103.0, 150.0, 276.0, 363.0, 399.0, 445.0, 476.0, 535.0]
    if boundaries != expected_boundaries:
        logging.warning("Header positions have changed. Adjusting columns accordingly, but you may get unexpected results.")
        logging.info(f"Expected columns boundaries {expected_boundaries}. Found {boundaries}.")
    return boundaries



def is_sentence_line(chars: list[dict], boundaries: list[int]):
    if not chars:
        return False
    else:
        return chars[0]['fontname'].lower().endswith('-bold')


def group_by_columns(chars: list[dict], boundaries: list[float]) -> list[str]:
    """Given a line of chars, combine chars that are in the same column.
    A new column starts at the given boundaries.
    Input: list of chars on a line (each a dict from pdfplumber).
    Output: list of the text of each word."""
    boundaries_dict = [(left, index) for index, left in enumerate(boundaries)]
    boundaries_dict.sort(reverse=True)
    def get_column_index(x0):
        nonlocal boundaries_dict
        for left, index in boundaries_dict:
            if x0 >= left:
                return index
        return 0
    row = len(boundaries) * ['']
    for char in chars:
        column = get_column_index(char['x0'])
        row[column] = row[column] + char['text']
    return row


def parse_transaction(line: list[str], year_hint: date) -> Transaction:
    """Given a line with the words in a transaction, return a Transaction object.
    Assumes that the words are in the appropriote order."""
    transaction_date = parse_date(line[0], year_hint)
    booking_date = parse_date(line[1], year_hint)
    description = line[2].strip()
    location = line[3].strip()
    country = line[4].strip()
    amount = parse_amount(line[7], line[8])
    foreign_amount = parse_amount(line[5]) if line[5] else None
    foreign_currency = line[6] if line[6] else None
    foreign_exchange_rate = None
    return Transaction(
        transaction_date = transaction_date,
        booking_date = booking_date,
        description = description,
        location = location,
        country = country,
        amount = amount,
        foreign_amount = foreign_amount,
        foreign_currency = foreign_currency,
    )

def parse_ics_pdf(path: Path) -> BankStatement:
    """This is the main function.
    It parses the PDF file and outputs it as a data structure.
    returns a BankStatement, or raises a ParseError or IOError."""

    # Location where info is found on the PDF
    # pdfplumber uses (left, top, right, bottom) for a bounding box (bbox).
    INFO_BOX_BBOX = (56,112,526,160)
    TABLE_HEADER_BBOX = (56,186,550,213)
    FIRST_PAGE_BODY_BBOX = (56,212,550,564)
    LATER_PAGE_BODY_BBOX = (56,210,550,700)

    with pdfplumber.open(path) as pdf:
        # Parse header
        headers = pdf.pages[0].within_bbox(INFO_BOX_BBOX)
        words = headers.extract_words(use_text_flow=False, keep_blank_chars=True)
        properties = header_words_to_dict(words)
        try:
            date = parse_date(' '.join(properties['datum']))
            customer_number = (' '.join(properties['ics-klantnummer'])).strip()
            if len(properties['volgnummer']) > 1:
                raise ParseError(f"Serial number has multiple words {' '.join(properties['ics-klantnummer'])}")
            try:
                serial_number = int(properties['volgnummer'][0])
            except ValueError:
                raise ParseError(f"Serial number {properties['ics-klantnummer'][0]} is not an integer") from None
            if 'vorig tegoed' in properties:
                previous_balance = parse_amount(*properties['vorig tegoed'])
            else:
                previous_balance = parse_amount(*properties['vorig openstaand saldo'])
            total_received_payment = parse_amount(*properties['totaal ontvangen betalingen'])
            total_new_expenses = parse_amount(*properties['totaal nieuwe uitgaven'])
            if 'nieuw tegoed' in properties:
                new_balance = parse_amount(*properties['nieuw tegoed'])
            else:
                new_balance = parse_amount(*properties['nieuw openstaand saldo'])
        except KeyError as err:
            # print(properties)
            raise ParseError(f"{err} not found in header") from None

        lines = []
        for page in pdf.pages:
            # Determine column boundaries from the table header
            header = page.within_bbox(TABLE_HEADER_BBOX)
            words = header.extract_words(use_text_flow=True, keep_blank_chars=True)
            boundaries = get_boundaries_from_first_line(words)
            assert len(boundaries) > 4

            # Extract lines of characters from the main table
            if page.page_number == 1:
                chars = page.within_bbox(FIRST_PAGE_BODY_BBOX).chars
            else:
                chars = page.within_bbox(LATER_PAGE_BODY_BBOX).chars
            char_lines = group_by_lines(chars)
            if not char_lines:
                raise ParseError(f"No text found on page {page.page_number}")

            # Extract cell content (words) for each line
            while char_lines:
                char_line = char_lines.pop(0)
                if is_sentence_line(char_line, boundaries):
                    text = ''.join(char['text'] for char in char_line)
                    lines.append([text])
                else:
                    line = group_by_columns(char_line, boundaries)
                    assert len(line) == len(boundaries)
                    lines.append(line)

        # end of with open(): the PDF file is closed here

    # Turn lines into Transaction records,
    # and parse sentences that appear in between transactions.
    RE_CARD_NO = re.compile(r'Uw Card met als laatste vier cijfers (\d+)', re.IGNORECASE)
    transactions: list[Transaction] = []
    card_last_numbers = ""
    card_owner = ""
    transaction: Transaction | None = None
    while lines:
        line = lines.pop(0)
        if len(line) == 1:
            if m := RE_CARD_NO.match(line[0]):
                card_last_numbers = m.group(1)
                if lines:
                    line = lines.pop(0)
                    if len(line) == 1:
                        card_owner = line[0].strip()
                    else:
                        lines.insert(0, line)
                        card_owner = ""
            else:
                logger.warning(f"Unknown line {' '.join(line)!r}")
            transaction = None
        elif line[0] == '' and line[1] == '' and line[2].startswith("Wisselkoers"):
            if transaction:
                transaction.foreign_exchange_rate = parse_amount(line[3])
                transaction = None  # we do not expect any more lines about this transaction
            else:
                raise ParseError(f"Found {line[2]} {line[3]}, without prior transaction line")
        else:
            transaction = parse_transaction(line, date)
            transaction.card_last_numbers = card_last_numbers
            transaction.card_owner = card_owner
            transactions.append(transaction)
    return BankStatement(
        date = date,
        customer_number = customer_number,
        serial_number = serial_number,
        previous_balance = previous_balance,
        total_received_payment = total_received_payment,
        total_new_expenses = total_new_expenses,
        new_balance = new_balance,
        transactions = transactions
    )


def destination_path(path: Path) -> Path:
    """Given a source path, return the destination Path.
    May raise a IOError."""
    # match case of the extension
    if path.suffix == '.pdf':
        output_suffix = DEFAULT_OUTPUT.lower()
    else:
        output_suffix = DEFAULT_OUTPUT
    if DESTINATION_DIR:
        dest_dir = Path(DESTINATION_DIR)
        if not dest_dir.is_dir():
            raise NotADirectoryError(f"Not a directory: {DESTINATION_DIR}")
    else:
        dest_dir = path.parent
    return dest_dir / (path.stem + '.' + output_suffix)


def sanity_checks(statement: BankStatement, path: Path):
    """Make some validity checks, and log a warning if they fail."""
    if m := re.search(r'\b(\d\d\d\d)[\-_](\d\d)\b', path.name):
        if int(m.group(1)) != statement.date.year or int(m.group(2)) != statement.date.month:
            logger.warning(f"{path.name} is dated {statement.date}, " \
                           f"expected in {m.group(1)}-{m.group(2)}.")
    max_rounding_error = 0.001
    new_expenses = sum(ta.amount for ta in statement.transactions if ta.amount < 0)
    if new_expenses - statement.total_new_expenses > max_rounding_error:
        logger.warning(f"Total new expenses is listed as " \
                       f"{statement.total_new_expenses:.2f}, " \
                       f"but sum of negative transactions is {new_expenses:.2f}")
    received_payments = sum(ta.amount for ta in statement.transactions if ta.amount > 0)
    if received_payments - statement.total_received_payment > max_rounding_error:
        logger.warning(f"Total received payments is listed as " \
                       f"{statement.total_received_payment:.2f}, " \
                       f"but sum of negative transactions is {received_payments:.2f}")
    expected_new_balance = (statement.previous_balance 
                          + statement.total_received_payment
                          + statement.total_new_expenses)
    if expected_new_balance - statement.new_balance > max_rounding_error:
        logger.warning(f"Expected new balance in {path.name} to be " \
                       f"{statement.previous_balance:0.2f} + " \
                       f"{statement.total_received_payment:0.2f} + " \
                       f"{statement.total_new_expenses:0.2f} = " \
                       f"{expected_new_balance:0.2f}. " \
                       f"Found {statement.new_balance:0.2f}")


def process_file(path: Path):
    """Parse a ICS PDF file and export to CSV, TSV or JSON,
    based on the value of DEFAULT_OUTPUT.
    Overwrites existing output files.
    May raise an IOError or ParsingError."""
    logger.info(f"Parsing {path.name}")
    dest_file = destination_path(path)
    statement = parse_ics_pdf(path)
    sanity_checks(statement, path)
    logger.debug(f"Writing {dest_file.name}")
    if DEFAULT_OUTPUT == 'JSON':
        statement.save_as_json(dest_file)
    elif DEFAULT_OUTPUT == 'TSV':
        statement.save_as_tsv(dest_file)
    else:
        statement.save_as_csv(dest_file)
    logger.info(f"Written {len(statement.transactions)} transactions to {dest_file.name}")
    return True


def process_folder(path: Path):
    """Convert all PDF files in a directory to CSV.
    Only consider PDF files matching *yyyy-mm*.pdf file name pattern,
    and skip files if there is already a CSV file."""
    counter = 0
    file = None
    try:
        files = list(path.glob("*[-_][0-9][0-9][0-9][0-9]-[0-9][0-9]*.pdf"))
        if not files:
            logger.warning("No PDF files found (matching *yyyy-mm*.pdf pattern)")
            return
        for file in files:
            dest_file = destination_path(path)  # may raise an IOError
            if dest_file.exists():
                logging.info(f"{dest_file.name} already exists")
                continue
            try:
                process_file(file)
                counter += 1
            except ParseError as err:
                logger.error(f"{file.name}: {err}")
    except IOError as err:
        if file:
            logger.error(f"{file.name}: {err}")
        else:
            logger.error(f"{err}")
    logger.info(f"{counter} files converted to {DEFAULT_OUTPUT}")


def usage():
    """Print usage information"""
    print(f"{__file__} [ICS_FILE.PDF|DIRECTORY]")
    print(f"Convert ICS bank statements from PDF to {DEFAULT_OUTPUT}.")
    print(f"If a directory is specified, process all PDF files in that directory,")
    print(f"but skip files were a {DEFAULT_OUTPUT} file already exists.")


def main(args):
    """Process the first given argument, either a file or a directory.
    If no arguments are given, process the directory of this script."""
    if args:
        path = Path(args[0])
    else:
        path = Path(__file__).parent
    if path.is_dir():
        process_folder(path)
    elif path.is_file() and path.suffix.lower == 'pdf':
        try:
            process_file(path)
        except (ParseError, IOError) as err:
            logging.error(f"{path.name}: {err}")
    else:
        logging.error(f"{path} is not a valid file or directory")


def test():
    """Parse some example files"""
    file_dir = Path(__file__).parent
    filename_prefix = "Rekeningoverzicht-"
    filename_suffix = ".pdf"

    # nothing special
    path1 = file_dir / (filename_prefix + "2025-08" + filename_suffix)

    # year roll-over
    path2 = file_dir / (filename_prefix + "2025-01" + filename_suffix)

    # multiple page (7 pages)
    path3 = file_dir / (filename_prefix + "2023-08" + filename_suffix)

    # description text overlaps with location text
    path4 = file_dir / (filename_prefix + "2018-06" + filename_suffix)

    # more card numbers on one overview
    path5 = file_dir / (filename_prefix + "2020-06" + filename_suffix)

    # old logo (same layout); serial number does not match month
    path6 = file_dir / (filename_prefix + "2016-02" + filename_suffix)

    # positive balance (usually this is negative)
    path7 = file_dir / (filename_prefix + "2024-04" + filename_suffix)

    # credit correction (description + location, no country)
    path8 = file_dir / (filename_prefix + "2017-06" + filename_suffix)

    # process_file(path1)
    for path in (path1, path2, path3, path4, path5, path6, path7, path8):
        process_file(path)


if __name__ == '__main__':
    main(sys.argv[1:])
    # test()
