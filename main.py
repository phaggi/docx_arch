import shutil

from pathlib import Path
from src.file_manager import Payload, TargetDocx
from src.fileman import prepare_filename


def test_file_manager():
    test_source_name = "DB Create Table As_2.knwf"

    payload_folder_name = "knimedb/"
    payload_filename = "DB Create Table As.knwf"
    target_docx_filename = prepare_filename(
        result_filename="panda", ext="docx", datetime_reverse=True
    )
    print()
    my_path = Path.cwd() / payload_folder_name
    source = my_path / test_source_name
    payload_source = my_path / payload_filename

    if not payload_source.exists():
        shutil.copyfile(source, payload_source)
    my_payload = Payload(payload_path=payload_source)
    my_docx = TargetDocx(docx_filename=target_docx_filename, work_dir=my_path)
    my_docx.add_payload(payload=my_payload)
    print(my_docx.check_payload)
    my_docx.rename()


def test_unarch():
    pass


if __name__ == "__main__":
    test_file_manager()
