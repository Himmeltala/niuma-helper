import re
from common.validate_util import check_url, check_mission
from config.regex import WXWORK_FILL_URL

TEXT = """
"""


def parse_text(text):
    matches = []
    unrecognized_lines = []

    for line in text.strip().split('\n'):
        if line.strip():
            match = re.match(
                WXWORK_FILL_URL, line)

            if match:
                matches.append(match)
            else:
                unrecognized_lines.append(line)
                check_url(line)

    return matches, unrecognized_lines


def main():
    matches, _ = parse_text(TEXT)
    for match in matches:
        check_mission(match)


if __name__ == "__main__":
    main()
