from src.file_manager import *

if __name__ == '__main__':
    my_file = 'pandas-1.5.0.tar.gz'
    my_path = Path.cwd() / 'pandas-dep/'
    template_path = Path.cwd() / 'templates/'
    my_file_full_path = my_path / my_file
    target_docx_filename = 'panda.docx'
    target_docx_fullpath = my_path / target_docx_filename

    my_docx = TargetDocx(docx_filename=target_docx_filename, work_dir=my_path)
    my_docx.add_payload()
    print(my_docx.check_payload)
    my_docx.remove_payload()
    print(my_docx.check_payload)
