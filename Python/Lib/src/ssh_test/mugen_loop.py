#!/usr/bin/env python3
import time

BREAK_SEC = 120


start = time.time()

while True:
    time.sleep(1)
    print('processing...')
    if time.time() - start > BREAK_SEC:
        print('!!BREAK!!')
        break
