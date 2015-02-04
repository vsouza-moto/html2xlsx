"""
    The 'extract' module provides functions to extract tests executions results
    from XML files.
"""

__author__ = "vsouza@motorola.com"
__date__ = "01/26/2015"

# importing
import logging
import os
import os.path

from lxml import etree


class Extractor(object):

    """
        This class implements the extracting capabilities of the 'extract'
        module.
    """

    def __init__(self, recursive=False, xml_regex=r".*\.xml"):

        self.recursive = recursive
        self.xml_regex = xml_regex

    def scan_folder(self, folder, results):

        """
            Scans a folder for xml files to extract information.
        """

        logging.info("Searching for xml files in the '{0}' folder.".format(folder))

        for child in os.listdir(folder):

            pathname = os.path.join(folder, child)

            # child is a file
            if os.path.isfile(pathname):

                # check if this is a valid tests results html file
                if self.xml_regex.match(child):

                    self.extract_info(pathname, results)

            # child is a folder
            else:

                if self.recursive:

                    self.scan_folder(pathname, results)

    def extract_info(self, filename, results):

        """
            Extracts tests executions results from the given xml file.
            The results of the extraction are put in the 'results' map.
        """

        logging.info("File to extract info: '{0}'".format(filename))

        info = {"names": [],
                "suites": [],
                "statuses": [],
                "times": [],
                "files": [],
                "functions": [],
                "descriptions": []}

        with open(filename) as file_handler:

            root = etree.parse(file_handler).getroot()

            if root.tag == "report":

                for testsuite in root.getchildren():

                    self.extract_testsuite(testsuite, info)

            elif root.tag == "testsuite":

                self.extract_testsuite(root, info)

            elif root.tag == "test_case_result":

                logging.warning("XML files using 'test_case_result' as their\
                    root are not supported yet!")

        # adds the extracted information to the main dict
        if reduce(lambda x, y: None if not x else y, info.values()):

            file_results = dict(zip(info["names"],
                                    zip(info["suites"],
                                        info["files"],
                                        info["functions"],
                                        info["descriptions"],
                                        info["statuses"],
                                        info["times"])))

            results[os.path.abspath(filename)] = file_results

    def extract_testsuite(self, testsuite, info):

        """
            Extracts the testcase elements from a testsuite.
        """

        for testcase in testsuite.getchildren():

            self.extract_testcase(testcase, info, testsuite.get("name"))

    def extract_testcase(self, testcase, info, testsuite_name):

        """
            Extracts the execution information from the testcase element.
        """

        # the test testsuite
        info["suites"].append(testsuite_name)

        try:

            # the test name
            info["names"].append(testcase.get("name"))

        except AttributeError:

            info["names"].append("")

        try:
            
            # the test description
            info["descriptions"].append(testcase.get("description"))

        except AttributeError:

            info["descriptions"].append("")

        try:

            # the test status
            info["statuses"].append(testcase.get("result"))

        except AttributeError:

            info["statuses"].append("")            

        try:

            # the test execution time
            info["times"].append(testcase.get("time"))

        except AttributeError:

            info["times"].append("")

        try:

            classname = testcase.get("classname").split(".")

            if len(classname) >= 2:

                # the test file
                info["files"].append(classname[-2])

                # the test function
                info["functions"].append(classname[-1])

            else:

                # the test file
                info["files"].append("")

                # the test function
                info["functions"].append("")

        except AttributeError:

            # the test file
            info["files"].append("")

            # the test function
            info["functions"].append("")
