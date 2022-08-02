"""Run this file after complete compile of the formulas file."""

import os
import pprint

from eu_total_compiler import EUTotalCompiler

a = EUTotalCompiler(os.path.join(os.path.curdir, 'data_passed.json'))
a.run()
pprint.pprint(a.geos_total_dict)
