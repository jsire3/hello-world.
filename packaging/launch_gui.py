import multiprocessing
import sys

# Must run first: in the packaged .exe, background engine processes re-launch
# this same executable, and this call routes them to their job.
multiprocessing.freeze_support()

from doc2word.app import main

sys.exit(main())
