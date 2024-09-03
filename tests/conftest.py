# tests/test_pythonpath.py

import sys
import os

sys.path.insert(0, os.path.abspath(
    os.path.join(os.path.dirname(__file__), '../src')))


def test_pythonpath():
    print("\nPYTHONPATH content in pytest:")
    for path in sys.path:
        print(path)
    assert True  # Just a dummy assertion to make sure the test passes
