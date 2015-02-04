"""
    Utilities used by the tests executions results extractor.
"""

__author__ = "vsouza@motorola.com"
__date__ = "01/26/2015"

import logging


def calculate_pass_rate(results):

    """
        Finds the pass rate (# of PASS tests / total # of tests) from the
        given map.
    """

    logging.info("Calculating pass rate...")

    passed_count = reduce(
        lambda x, y: x + 1 if y == "PASS" else x,
        [info[4] for info in results.values()],
        0
    )

    return "{0:.2f}%".format(100 * (passed_count / float(len(results))))


def calculate_total_exec_time(results):

    """
        Gives the total run time of a test round.
    """

    logging.info("Calculating total execution time...")

    total_seconds = sum([float(info[5]) for info in results.values()])

    minutes, seconds = divmod(total_seconds, 60)
    hours, minutes = divmod(minutes, 60)

    final_time = []

    if hours != 0:

        final_time.append("{0}h".format(int(round(hours))))

    if minutes != 0:

        final_time.append("{0}m".format(int(round(minutes))))

    if seconds != 0:

        final_time.append("{0}s".format(int(round(seconds))))

    return "".join(final_time)
