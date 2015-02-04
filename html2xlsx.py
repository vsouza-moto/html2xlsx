#!/usr/bin/env python

"""
    A script that extracts tests executions results, from XML files to XLSX
    spreadsheets.
"""

__author__ = "vsouza@motorola.com"
__date__ = "01/26/2015"

# importing
import argparse
import re
import os
import os.path
import sys
import logging

from extract import Extractor
from export import Exporter

if __name__ == "__main__":

    parser = argparse.ArgumentParser()

    parser.description = __doc__

    parser.add_argument("input_folder",
                        nargs="*",
                        default=".",
                        help="The input folder(s) where the html files\
                        are located."
                        )

    parser.add_argument("--output_folder",
                        "-o",
                        default=".",
                        help="The folder where the output should be generated."
                        )

    parser.add_argument("--verbose",
                        action="store_true",
                        default=False,
                        help="If the program should give more information on\
                        its progress.")

    parser.add_argument("--consolidate",
                        "-c",
                        action="store_true",
                        default=False,
                        help="Whether or not to consolidate the results from\
                        all the input folders in one spreadsheet. Conflicts\
                        regarding a same test round in different folders will\
                        be resolved by considering the result of the most\
                        recent run.")

    parser.add_argument("--recursive",
                        "-r",
                        action="store_true",
                        default=False,
                        help="Whether or not the program should search\
                        recursively for html files in the given\
                        input directories.")

    parser.add_argument("--xml_filename",
                        "-f",
                        default=None,
                        help="The filename for the xml files the program\
                        will try to parse (without the trailing '.xml').\
                        Regular expressions are accepted.")

    parser.add_argument("--basename",
                        "-b",
                        default=None,
                        help="Usually the name of the module of feature being \
                        tested. This name will be used to compose the file \
                        name of the generated spreadsheet.")

    args = parser.parse_args()

    # retrieving the arguments

    input_folders = args.input_folder

    if not isinstance(input_folders, list):
        input_folders = [input_folders]

    output_folder = args.output_folder

    if args.verbose:

        # set logging level to 'INFO'
        logging.getLogger().setLevel(20)

    xml_regex = r".*\.xml"

    if args.xml_filename:

        xml_regex = r"{0}\.xml".format(args.xml_filename)

    xml_regex = re.compile(xml_regex)

    if os.path.exists(output_folder):

        if os.path.isfile(output_folder):

            logging.critical("The '{0}' output folder can't be created! A file\
            with the same name already exists.".format(output_folder))

            sys.exit(1)

    else:

        logging.info("Creating the '{0}' output folder...".format(output_folder))

        os.makedirs(output_folder)

    extracted_info = dict()

    extractor = Extractor(recursive=args.recursive, xml_regex=xml_regex)

    # run through all input folders and extract the data from the html files
    for input_folder in input_folders:

        if os.path.exists(input_folder):

            if os.path.isdir(input_folder):

                extractor.scan_folder(input_folder, extracted_info)

            else:

                logging.info("The '{0}' input folder is a file!".format(input_folder))

        else:

            logging.info("The '{0}' input folder does not exist!".format(input_folder))

    if not extracted_info:

        logging.critical("No information could be extracted.")

        sys.exit(1)

    else:

        exporter = Exporter(basename=args.basename, output_folder=output_folder)

        if args.consolidate:

            exporter.export_to_xlsx_consolidated(extracted_info)

        else:

            exporter.export_to_xlsx(extracted_info)
