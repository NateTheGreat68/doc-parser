from typing import Optional, List
import os
import zipfile
import re
from xml_wrapper import XMLWrapper


class GenericParser:
    """
    This class contains methods for reading data from an Office Open XML (OOX)
    file. Methods specific to a particular type of file (Word Processing,
    Spreadsheets, and Presentations) are in their respective subclasses.

    Init Parameters:
      - filepath: The path to the OOX file.
    """

    def __init__(
            self,
            filepath: str,
            ):
        self.filepath = filepath
        self.filename = os.path.basename(self.filepath)
        self.filename_base, self.filename_ext = os.path.splitext(self.filename)
        self.text_node_tag = 'w:t'
        
        self._xml_files = {}
        with zipfile.ZipFile(self.filepath) as z:
            for xml_filename in [x for x in z.namelist() if x.endswith('.xml')]:
                self._xml_files[xml_filename] = None

    def xml_file(
            self,
            xml_filename: str,
            ) -> Optional[XMLWrapper]:
        """
        Fetches an XMLWrapper instance for the given XML file contained within the
        document. Will return an existing instance of one already exists, or open
        and parse a new one otherwise. The DOM will always be reset to root.

        Parameters:
          - xml_filename: The path (relative to the document root) of the XML file.
        Return: The XMLWrapper instance; None on failure.
        """
        try:
            if self._xml_files[xml_filename] is None:
                with zipfile.ZipFile(self.filepath) as z:
                    self._xml_files[xml_filename] = XMLWrapper(
                            z.read(xml_filename),
                            self.text_node_tag,
                            )
            return self._xml_files[xml_filename].reset_to_root()
        except:
            return None

    @property
    def xml_files(self) -> List[str]:
        return self._xml_files.keys()


class WordProcessingParser(GenericParser):
    def __init__(
            self,
            filepath: str,
            ):
        super().__init__(
                filepath,
                )
        self.text_node_tag = 'w:t'


class SpreadsheetParser(GenericParser):
    pass


class PresentationParser(GenericParser):
    pass
