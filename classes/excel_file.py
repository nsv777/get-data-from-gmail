import os
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook


class ExcelFile(object):
    def __init__(self, filename=None):
        self.book = Workbook()
        self.sheet = self.book.worksheets[0]
        self.filename = self.get_filename(filename)
        self.save()

    def save(self):
        self.book.save(filename=self.filename)

    def get_filename(self, filename):
        timestamp = (str(datetime.now().strftime('%Y%m%d_%H%M%S')))
        ext = "xlsx"
        if filename is not None:
            return Path.joinpath(
                Path(filename).resolve().parent,
                "{}_{}.{}".format(Path(filename).stem, timestamp, ext)
            )
        return Path.joinpath(
            self._get_dirpath(),
            "messages_{}.{}".format(timestamp, ext)
        )

    @staticmethod
    def _get_dirpath():
        dirpath = None
        if sys.platform == "win32":
            dirpath = Path('c:/', 'Users', os.getlogin(), 'Downloads')
        elif sys.platform == "linux":
            dirpath = Path('/home', os.getlogin(), 'Downloads')
        # if dirpath and dirpath.exists():
        #     shutil.rmtree(dirpath)
        #     dirpath.mkdir()
        if dirpath and not dirpath.exists():
            dirpath.mkdir()
        return dirpath

    def add_row(self, row):
        self.sheet.append(row)
