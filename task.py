"""Template robot with Python."""

# +
from test import main
import os
import sys
import logging


def initialize_logging():
    stdout = logging.StreamHandler(sys.stdout)

    logging.basicConfig(
        level=logging.INFO,
        format="{%(filename)s:%(lineno)d} %(levelname)s - %(message)s",
        handlers=[stdout]
    )


if __name__ == "__main__":
    initialize_logging()
    main(os.getcwd() + "\\output\\")
# -


