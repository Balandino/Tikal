"""Utilities for xlsxwriter"""

import os
import sys

import xlsxwriter
from xlsxwriter import Workbook
from xlsxwriter.workbook import FileCreateError


def close_workbook(WORKBOOK: Workbook, name=None):
    """
    Writes the workbook to disk.  If fails, will offer to retry via user input.
    If the name is passed, will attempt to open the Workbook

    Args:
        WORKBOOK: Workbook object to be saved
        name: Name opf the workbook to open
    """

    while True:
        try:
            WORKBOOK.close()
            break
        except FileCreateError:
            print("[ERROR] Unable to save Workbook, maybe it is already open...")
            retry = input("Retry? (y/n): ")

            if retry != "y":
                print("Quitting")
                sys.exit(1)

    if name is not None:
        os.startfile(os.getcwd() + "\\" + name, "open")


def set_global_font(WORKBOOK: Workbook):
    """
    Sets the global workbook font.  WARNING: methods like autosize will not work correctly
    once the default font is changed, so sizing will have to be calculated manually

    Args:
        WORKBOOK: Workbook object to have font changed
    """

    WORKBOOK.formats[0].set_font_name("Tenorite")


def create_workbook(workbook_name: str) -> Workbook:
    """
    Creates a workbook object with 1 font and returns it

    Args:
        workbook_name: The name to use for saving the workbook

    Returns:
        xlsxwriter Workbook object

    """

    workbook = xlsxwriter.Workbook(workbook_name)
    set_global_font(workbook)
    return workbook

