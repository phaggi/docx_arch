import shutil

from src.file_manager import *


def test_file_manager():
    my_path = Path.cwd() / "knimedb/"
    source = my_path / "DB Create Table As_2.knwf"

    template_path = Path.cwd() / "templates/"

    target_docx_filename = "panda.docx"
    target_docx_fullpath = my_path / target_docx_filename

    payload_source = my_path / "DB Create Table As.knwf"
    if not payload_source.exists():
        shutil.copyfile(source, payload_source)
    my_payload = Payload(payload_path=payload_source)
    my_docx = TargetDocx(docx_filename=target_docx_filename, work_dir=my_path)
    my_docx.add_payload(payload=my_payload)
    print(my_docx.check_payload)
    my_docx.rename()


if __name__ == "__main__":
    test_file_manager()
