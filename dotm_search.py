#!/usr/bin/env python
# -*- coding: utf-8 -*-
#importation
"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Sasha Lukas"

import os
import sys
import zipfile
import argparse
DOC_FILENAME = 'word/document.xml'



def find_text_zipfile(z, search_text, full_path):
    with z.open(DOC_FILENAME) as doc:
        xml_text = doc.read()
    xml_text = xml_text.decode('utf-8')
    text_location = xml_text.find(search_text)
    if text_location >= 0:
        print('Match found in file {}'.format(full_path))
        print('     ...' + xml_text[text_location-40:text_location] + '...')
        return True
    return False

def create_parser():
    parser = argparse.ArgumentParser(description="Find text in dotm files")
    parser.add_argument('--dir', help='directory to find dotm files', default='.')
    parser.add_argument('text', help='find text')
    return parser


def main():
    parser = create_parser()
    args = parser.parse_args()

    text_search = args.text
    dir_search = args.dir

    if not text_search:
        parser.print_usage()
        sys.exit(1)

    print('Search directory for dotm files "{}" ...'.format(dir_search))

    file_list = os.listdir(dir_search)

    match_count = 0
    search_count = 0

    for file in file_list:
        if not file.endswith('.dotm'):
            print('Not a dotm file: ' + file)
            continue
        else:
            search_count += 1

        full_path = os.path.join(dir_search, file)

        if zipfile.is_zipfile(full_path):
            with zipfile.ZipFile(full_path) as zip:
                names = zip.namelist()
                if DOC_FILENAME in names:
                    if find_text_zipfile(zip, text_search, full_path):
                        match_count += 1
        else:
            print('Not a zipfile: ' + full_path)


    print('Total dotm files searched: {}'.format(search_count))
    print('Total dotm files matched: {}'.format(match_count)
    )

    


if __name__ == '__main__':
    main()