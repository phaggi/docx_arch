import json
import pathlib
import zipfile

from shutil import copyfile
from os import remove
from pathlib import Path


class Payload:
    def __init__(self, payload_path: pathlib.Path = None):
        self.target_name = "payload.xml"
        self.name: pathlib.Path = Path(self.target_name)
        self.templatedir: pathlib.Path = Path.cwd() / "templates/"
        self.payload_dir: pathlib.Path = Path.cwd() / "templates/"
        if payload_path is not None:
            self.payload_path: pathlib.Path = payload_path
            self.name: pathlib.Path = payload_path
            self.original_name: pathlib.Path = self.name
            self.prepare_payload()
        self.payload_path: pathlib.Path = self.payload_dir / self.target_name
        self.payload_ini_path = self.payload_dir / "payload_ini.xml"

    def prepare_payload(self):
        try:
            self.payload_dir: pathlib.Path = self.name.parent
            self.name.rename(self.payload_dir / self.target_name)
        except FileNotFoundError as e:
            print(e)

    @property
    def ini_file(self):
        data = {"original_name": str(self.original_name.name)}
        with open(self.payload_ini_path, "w") as payload_ini_file:
            payload_ini_file.write(json.dumps(data))
        return self.payload_ini_path


class TargetDocx:
    def __init__(
        self,
        docx_filename: pathlib.Path or str,
        work_dir,
        template=None,
        payload: Payload or None = None,
    ):
        self.current_filename = None
        if template is None:
            self.template: pathlib.Path = Path.cwd() / "templates/temp.zip"
        self.new_filename = None
        self.status = None
        self.work_dir = work_dir
        self.docx_filename: pathlib.Path = self.work_dir / docx_filename
        self.temp_arch_filename: pathlib.Path = self.work_dir / "temp.zip"
        self.prepare_docx()
        self.buffer: pathlib.Path = self.work_dir / "buffer.zip"
        if self.buffer.exists():
            remove(self.buffer)
        if payload is None:
            self.payload = Payload()
        else:
            self.payload = payload
        self.target_payload_name = "word/payload.xml"
        self.target_payload_ini = "word/payload_ini.xml"

    def prepare_docx(self):
        if not self.docx_filename.exists():
            if self.temp_arch_filename.exists():
                remove(self.temp_arch_filename)
            copyfile(self.template, self.temp_arch_filename)
            self.status = "zip"
            self.new_filename: pathlib.Path = self.temp_arch_filename
            self.current_filename: pathlib.Path = self.temp_arch_filename
        else:
            self.status = "docx"
            self.new_filename: pathlib.Path = self.docx_filename
            self.current_filename: pathlib.Path = self.docx_filename

    def rename(self):
        assert isinstance(self.status, str)
        try:
            match self.status:
                case "zip":
                    self.rename_zip2docx()
                case "docx":
                    self.rename_docx2zip()
                case "buffer":
                    self.rename_buffer2zip()
            self.rename2current()
        except Exception as e:
            print(f"Error {e} at TargetDocx.rename. Maybe python version < 3.10.")

    def rename_docx2zip(self):
        self.new_filename = self.temp_arch_filename
        self.status = "zip"

    def rename2current(self):
        self.current_filename.rename(self.new_filename)
        self.current_filename = self.new_filename

    def rename_zip2docx(self):
        self.new_filename = self.docx_filename
        self.status = "docx"

    def __repr__(self):
        return f"{self.status}: {self.current_filename}"

    @property
    def check_payload(self):
        if self.status == "docx":
            self.rename()
        with zipfile.ZipFile(self.temp_arch_filename, mode="r") as zin:
            return self.target_payload_name in zin.namelist()

    def remove_payload(self):
        if self.status == "docx":
            self.rename()
        if self.check_payload:
            with zipfile.ZipFile(self.temp_arch_filename, mode="r") as zip_in:
                with zipfile.ZipFile(self.buffer, mode="w") as zip_out:
                    if self.target_payload_name in zip_in.namelist():
                        for item in zip_in.infolist():
                            buffer = zip_in.read(item.filename)
                            if not item.filename == self.target_payload_name:
                                zip_out.writestr(item, buffer)
            self.status = "buffer"
            self.current_filename = self.buffer
            self.rename()

    def add_payload(self, payload: Payload = None):
        if payload is not None:
            self.payload = payload
        if self.buffer.exists():
            remove(self.buffer)
        with zipfile.ZipFile(self.temp_arch_filename, mode="a") as zip_file:
            zip_file.write(self.payload.payload_path, self.target_payload_name)
            zip_file.write(self.payload.ini_file, self.target_payload_ini)
        if self.check_payload:
            remove(self.payload.payload_path)
            remove(self.payload.payload_ini_path)

    def rename_buffer2zip(self):
        self.new_filename = self.temp_arch_filename
        self.status = "zip"

    def get_payload(self):
        if self.check_payload:
            pass