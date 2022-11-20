#!/usr/bin/env python3
#


from doc_parser import WordProcessingParser, ParserCollection
from typing import Optional, Generator, List, Dict
import os
import re
import csv


DOC_PATH = 'test_docs/'
CSV_DOCS = 'output_docs.csv'
CSV_REFS = 'output_refs.csv'
CSV_RESPS = 'output_resps.csv'


class SOP(WordProcessingParser):
    _re_doc_match = re.compile(r'^(?P<num>[.A-Z]+-\d+), (?P<name>.*)$')

    def __init__(
            self,
            filepath: str,
            ):
        super().__init__(
                filepath,
                )
        
        # Extract the document name and number.
        re_doc_match = SOP._re_doc_match.match(self.filename_base)
        self.doc_number = re_doc_match.group('num') if re_doc_match else None
        self.doc_name = re_doc_match.group('name') if re_doc_match else None

    @property
    def is_sop(
            self,
            ) -> bool:
        try:
            return self.issued_by is not None
        except:
            return None

    @property
    def issued_by(
            self,
            ) -> Optional[str]:
        try:
            return self.xml_file('word/header1.xml') \
                    .set_text_node_by_pattern(r'^Issued By:?\w*$') \
                    .set_ancestor_node_by_tag('w:tc') \
                    .set_nextSibling() \
                    .get_text_value()
        except:
            return None

    @property
    def has_references(
            self,
            ) -> bool:
        try:
            return self.xml_file('word/document.xml') \
                    .set_text_node_by_pattern(r'^References?:?\w*$') \
                    is not None
        except:
            return None

    @property
    def references(
            self,
            ) -> List[Dict[str, str]]:
        refs = []
        try:
            r = re.compile(r'^\t?(?P<num>[-.A-Z0-9]+)\t(?P<desc>[^\t]+)$')
            n = self.xml_file('word/document.xml') \
                    .set_text_node_by_pattern(r'^References?:?\w*$') \
                    .set_ancestor_node_by_tag('w:p')
            while n.set_nextSibling():
                re_match = r.match(n.get_text_value())
                if not re_match: #No match means it's time to stop parsing
                    break
                refs.append({
                    'doc_number': re_match.group('num'),
                    'description': re_match.group('desc'),
                    })
            return refs
        except:
            return refs

    @property
    def has_responsibilities(
            self,
            ) -> bool:
        try:
            return self.xml_file('word/document.xml') \
                    .set_text_node_by_pattern(r'^Responsibilit(y|ies):?\w*$') \
                    is not None
        except:
            return None

    @property
    def responsibilities(
            self,
            ) -> bool:
        resps = []
        try:
            r = re.compile(r'^\t?(?P<party>[A-Za-z0-9, ]+)\t(?P<resp>[^\t]+)$')
            n = self.xml_file('word/document.xml') \
                    .set_text_node_by_pattern(r'^Responsibilit(y|ies):?\w*$') \
                    .set_ancestor_node_by_tag('w:p')
            while n.set_nextSibling():
                re_match = r.match(n.get_text_value())
                if not re_match: #No match means it's time to stop parsing
                    break
                resps.append({
                    'party': re_match.group('party'),
                    'responsibilities': re_match.group('resp'),
                    })
            return resps
        except:
            return resps


if __name__ == '__main__':
    docs_collection = ParserCollection(
        SOP,
        DOC_PATH,
        [
            'key',
            'filename',
            'doc_number',
            'doc_name',
            'is_sop',
            'issued_by',
            'has_references',
            'has_responsibilities',
            ],
        {
            'references': [
                'doc_key',
                'doc_number',
                'description',
                ],
            'responsibilities': [
                'doc_key',
                'party',
                'responsibilities',
                ],
            },
        )
    outputs = [
            {
                'filename': CSV_DOCS,
                },
            {
                'filename': CSV_REFS,
                'extended_attr': 'references',
                },
            {
                'filename': CSV_RESPS,
                'extended_attr': 'responsibilities',
                },
            ]
    for o in outputs:
        docs_collection.output_csv(**o)
