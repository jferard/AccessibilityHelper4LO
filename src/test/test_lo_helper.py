import logging
import os
import unittest
import platform
from pathlib import Path

from py4lo_commons import init_logger


class LOHelperTestCase(unittest.TestCase):
    def test_init_logger(self):
        system = platform.system()
        if system == "Windows":
            try:
                log_path = Path(os.environ["appdata"]) / "AccessibilityHelper4LO.log"
            except KeyError:
                log_path = None
        elif system == "Linux":
            log_path = Path("/var/log/AccessibilityHelper4LO.log")
        else:
            log_path = None
        init_logger(logging.getLogger(), log_path)


if __name__ == '__main__':
    unittest.main()
