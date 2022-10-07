import pathlib
import zipfile
from shutil import copyfile
from os import remove
from pathlib import Path
from typing import Dict, ClassVar, get_type_hints
from zipfile import ZipFile


class Payload:
    def __init__(self, payload_path=None):
        if payload_path is None:
            self.name = "payload.xml"
            self.templatedir: pathlib.Path = Path.cwd() / "templates/"
            self.payload_path: pathlib.Path = self.templatedir / self.name
        else:
            self.payload_path: pathlib.Path = payload_path
            self.name = self.payload_path.name


class TargetDocx:
    def __init__(
        self,
        docx_filename: pathlib.Path or str,
        work_dir: pathlib.Path,
        template=None,
        payload=None,
    ):
        self.current_filename = None
        if template is None:
            self.template: pathlib.Path = Path.cwd() / "templates/temp.zip"
        self.new_filename = None
        self.status = None
        if not work_dir.exists():
            work_dir.mkdir()
        self.work_dir = work_dir
        self.docx_filename: pathlib.Path = self.work_dir / docx_filename
        self.temp_arch_filename: pathlib.Path = self.work_dir / "temp.zip"
        self.prepare_docx()
        self.buffer: pathlib.Path = self.work_dir / "buffer.zip"
        if self.buffer.exists():
            remove(self.buffer)
        if payload is None:
            self.payload = Payload()
        self.target_payload_name = "word/payload.xml"

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
        match self.status:
            case "zip":
                self.rename_zip2docx()
            case "docx":
                self.rename_docx2zip()
            case "buffer":
                self.rename_buffer2zip()
        self.rename2current()

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

    def add_payload(self):
        if self.buffer.exists():
            remove(self.buffer)
        with zipfile.ZipFile(self.temp_arch_filename, mode="a") as zip_file:
            zip_file.write(self.payload.payload_path, self.target_payload_name)

    def rename_buffer2zip(self):
        self.new_filename = self.temp_arch_filename
        self.status = "zip"

    def get_payload(self):
        if self.check_payload:
            pass
