import pathlib
from datetime import datetime
from pathlib import Path


def get_datetime(rev: bool = False):
    if rev:
        date_time = str(datetime.now().strftime("%Y.%m.%d_%H.%M"))
    else:
        date_time = str(datetime.now().strftime("%d.%m.%Y_%H.%M"))
    return date_time


def prepare_filename(
    result_filename: str or None = None, ext: str or None = None, datetime_reverse=True
) -> pathlib.Path:
    """
    :param result_filename: str text for filename
    :param ext: str suffix '.xxx' or 'xxx'
    :return: Path(filename.ext)
    :param datetime_reverse: True for '2020.10.23_18.15', False for reverse date
    """
    if not result_filename:
        result_filename = "выгрузка"
    if not ext:
        ext = "xlsx"
    elif ext.startswith("."):
        ext = ext[1:]
    result_filename = "_".join([result_filename, get_datetime(rev=datetime_reverse)])
    return Path(".".join([result_filename, ext]))


if __name__ == "__main__":
    result_file_name = "panda"
    result_ext = "docx"

    full_result_filename = prepare_filename(
        result_filename=result_file_name, ext=result_ext
    )
    print(full_result_filename.name)
