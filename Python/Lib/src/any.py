#!/usr/bin/env python3
import pickle
import pprint
import os
import subprocess
import glob
import re
import traceback


# ------------------------
# dict -> file 4
from pprint import pprint

hoge_obj = {'has': {'plants': 'yes', 'animals': 'yes',
                    'cryptonite': 'no'}, 'name': 'Earth'}

with open("hoge.txt", "w") as f:
    pprint(hoge_obj, width=40, stream=f)


# ------------------------
# dict -> file 3
# def format(d, tab=0):
#     s = ['{\n']
#     for k, v in d.items():
#         if isinstance(v, dict):
#             v = format(v, tab+1)
#         else:
#             v = repr(v)

#         s.append('%s%r: %s,\n' % ('  '*tab, k, v))
#     s.append('%s}' % ('  '*tab))
#     return ''.join(s)


# a = {'has': {'plants': 'yes', 'animals': 'yes', 'cryptonite': 'no'}, 'name': 'Earth'}
# print(format(a, 1))


# ------------------------
# dict -> file 2
# mydictionary = {'1': 123, 'A': 'xyz'}

# pp = pprint.PrettyPrinter(indent=4)
# pp.pprint(mydictionary)

# with open('myfile.txt', 'w') as f:
#     print(pprint.pformat(mydictionary, depth=2, width=40, indent=2), file=f)

# ------------------------
# dict -> file
# mydictionary = {'1': 123, 'A': 'xyz'}

# with open('myfile.txt', 'w') as f:
#     print(mydictionary, file=f)
