#!/usr/bin/env python3
import time
import sys


def main():
    args = sys.argv
    instance = ''
    if len(args) > 1:
        instance = args[1]
    filepath = "/tmp/hello.log"
    with open(filepath, 'a') as f:
        f.write(f"hello! {instance} \n")


if __name__ == '__main__':
    while True:
        main()
        time.sleep(30)
