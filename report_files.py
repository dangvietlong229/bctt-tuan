"""Shared report-date and input selection rules for both processors."""
import datetime as dt
import glob
import os
import re


def friday_on_or_before(day):
    return day - dt.timedelta(days=(day.weekday() - 4) % 7)


def get_report_as_of():
    value = os.environ.get("REPORT_AS_OF", "").strip()
    if value:
        try:
            return dt.date.fromisoformat(value)
        except ValueError as exc:
            raise ValueError(f"Invalid REPORT_AS_OF: {value}; expected YYYY-MM-DD") from exc
    return friday_on_or_before(dt.date.today())


def get_update_suffix():
    return f"_update {get_report_as_of():%d%m%y}"


def date_from_filename(path):
    name = os.path.basename(path)
    # The processed date takes precedence over the original export date.
    for pattern, fmt in (
        (r"(?i)update[ _-]*(\d{6})(?!\d)", "%d%m%y"),
        (r"(?<!\d)(20\d{6})(?!\d)", "%Y%m%d"),
        (r"(?<!\d)(\d{8})(?!\d)", "%d%m%Y"),
    ):
        for token in re.findall(pattern, name):
            try:
                return dt.datetime.strptime(token, fmt).date()
            except ValueError:
                continue
    return None


def select_file_for_as_of(paths, as_of, require_exact=False):
    dated = [(path, date_from_filename(path)) for path in paths
             if os.path.isfile(path) and not os.path.basename(path).startswith(("~$", "."))]
    eligible = [(path, day) for path, day in dated
                if day is not None and (day == as_of if require_exact else day <= as_of)]
    if eligible:
        return max(eligible, key=lambda item: (item[1], os.path.getmtime(item[0]), str(item[0])))[0]
    if require_exact:
        return None
    return max((path for path, day in dated if day is None),
               key=lambda path: (os.path.getmtime(path), str(path)), default=None)


def find_latest_template(directory, pattern, fallback_name):
    return select_file_for_as_of(glob.glob(os.path.join(directory, pattern)), get_report_as_of()) or os.path.join(directory, fallback_name)


def filter_raw_files(paths):
    return [path for path in paths if not os.path.basename(path).startswith(("~$", "."))
            and "_update" not in os.path.basename(path).lower()]


def parse_module_exclusions(value):
    if not value.strip():
        return set()
    numbers = {item.strip() for item in value.split(",")}
    if not numbers <= {str(number) for number in range(2, 11)}:
        raise ValueError("Module numbers must be 2-10, separated by commas")
    return numbers
