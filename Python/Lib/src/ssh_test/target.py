#!/usr/bin/env python3
import sys
import util


def main():
    args = sys.argv

    if len(args) <= 1:
        util.print_stderr(f'argment nothing.')
        return 1

    result = int(args[1])

    if result == 0:
        return result

    util.print_stderr(
        f'Test script error! [{result}][host={util.get_hostname()}]')
    return result


if __name__ == '__main__':
    sys.exit(main())
