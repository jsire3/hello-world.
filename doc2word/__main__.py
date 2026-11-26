import multiprocessing
import sys

from .app import main

if __name__ == "__main__":  # guard: background processes re-import this module
    multiprocessing.freeze_support()
    sys.exit(main())
