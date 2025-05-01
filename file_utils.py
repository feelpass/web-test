import os
from pathlib import Path


def generate_unique_filename(original_path: str) -> str:
    """
    Generate a unique filename based on the original filename.
    If file already exists, append a number to make it unique.

    Args:
        original_path (str): Original file path

    Returns:
        str: Unique file path
    """
    original_path = Path(original_path)
    directory = original_path.parent
    filename = original_path.stem
    extension = original_path.suffix

    counter = 1
    new_path = original_path

    while new_path.exists():
        new_filename = f"{filename}_{counter}{extension}"
        new_path = directory / new_filename
        counter += 1

    return str(new_path)
