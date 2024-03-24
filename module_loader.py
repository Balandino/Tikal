"""Loads the modules for template_runner.py"""

from module_profile import load_profile


def add_module(start_cell: str, module: dict, formats: dict) -> str:
    """
    Writes the relevant module into the worksheet and returns the starting position for the next cell

    Args:
        name: Name of the module to be written
        module: Dictionary of data required for module
        formats: Dictionary of formats to use

    Returns:
        Cell position for next module to start at, e,g A20

    """
    match module["name"]:
        case "profile":
            return load_profile(start_cell, module, formats)

        case _:
            return start_cell

