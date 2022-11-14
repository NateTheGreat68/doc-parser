from typing import Optional, List, Dict, Type, Generator, Union
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


class ParserCollection:
    """
    Provides assistance with creating a collection of parsed documents and
    generating outputs from them.

    Init Parameters:
      - parser: The class of parser to use (GenericParser or a derivative).
      - files: If str: a directory to be traversed (including subdirs) for
        files to parse. If list: a list of filenames to be parsed.
      - output_props: Optional. A List[str], the name(s) of the parser
        named attrs which will be output by the various output methods here.
      - output_props_extended: Optional. A Dict[str, List[str]. The keys
        should be the names of attrs within the parser, each of which is a
        List[Dict[str, Any]]. This is used for the extended output
        methods here.
    """

    def __init__(
            self,
            parser: Type[GenericParser],
            files: Union[str, List[str]],
            output_props: List[str] = [],
            output_props_extended: Dict[str, List[str]] = {},
            ):
        self.parser = parser
        self.docs = {}
        self._next_key = 0
        if type(files) == str: #Walk the directory.
            for (path, dirnames, filenames) in os.walk(files):
                for filename in filenames:
                    self.add_file(os.path.join(path, filename))
        else: #Handle the list of filepaths.
            for filepath in files:
                self.add_file(filepath)
        self.output_props = output_props
        self.output_props_extended = output_props_extended

    def add_file(
            self,
            filepath: str,
            ) -> None:
        """
        Add a new file to the collection and parse it.

        Parameters:
          - filepath: The path to the file.
        """
        self.docs[self._next_key] = self.parser(filepath)
        self._next_key += 1

    def output_data(
            self,
            **kwargs,
            ) -> 'ParserCollection.OutputData':
        """
        A factory for creating ParserCollection.OutputData class.
        See that class for additional information on usage.

        Parameters:
          - **kwargs: See the init for ParseCollection.OutputData
            for information on parameters.
        Return: A built ParserCollection.OutputData class.
        """
        return ParserCollection.OutputData(
                self,
                **kwargs,
                )


    class OutputData:
        """
        Used to create output data which can be consumed by other methods to
        output csv files, SQL tables, etc.

        Init Parameters:
          - parser_collection: The instance of the outer ParserCollection class.
          - props: Optional. A list of the parser attrs to be output in each dict.
            If not provided, the output_props specified at class initialization is
            used.
          - extended_attr: Optional. The name of the parser's attribute to use when
            fetching "extended" data (data for which there are 0 or more records
            per document, rather than exactly one record per document). When None
            (default), doesn't act as extended data.
        """

        def __init__(
                self,
                parser_collection: 'ParserCollection',
                props: List[str] = None,
                extended_attr: str = None,
                ):
            self.parser_collection = parser_collection
            if not props:
                if not extended_attr:
                    self.props = parser_collection.output_props
                else:
                    self.props = parser_collection.output_props_extended[extended_attr]
            self.extended_attr = extended_attr

        @property
        def field_list(
                self,
                ) -> List[str]:
            """
            Provides a field list for the compiled output data.

            Return: A list of the fields (props) provided by this OutputData.
            """
            return self.props

        @property
        def records(
                self,
                ) -> Generator[Dict, None, None]:
            """
            Creates a dict for each doc in the collection with the selected properties.

            Return: Generator, each iteration yielding a dict with props as keys and
            the parser's related attrs as values. Each dict includes a "key" value
            (when not in extended mode) or a "doc_key" value (when in extended mode).
            """
            props = self.props
            extended_attr = self.extended_attr
            for key, doc in self.parser_collection.docs.items():
                key = getattr(doc, 'key', key) #Use the parser's key attr if it has it
                if not extended_attr: #Fetch doc attrs (non-extended mode)
                    try:
                        output = {}
                        output['key'] = key
                        for prop in [p for p in props if p != 'key']:
                            output[prop] = getattr(doc, prop, None)
                        yield output
                    except:
                        continue
                else: #Fetch 0 or more records from a single doc attr (extended mode)
                    try:
                        for record in getattr(doc, extended_attr, []): #Empty list default to prevent error
                            output = {}
                            output['doc_key'] = key
                            for prop in [p for p in props if p != 'doc_key']:
                                output[prop] = record[prop] if prop in record else None
                            yield output
                    except:
                        continue
