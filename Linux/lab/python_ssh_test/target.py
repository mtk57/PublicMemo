#!/usr/bin/env python3
import sys
import util


def main():
    try:
        result = util.run_command(['systemctl', 'restart', 'chronyd'])
        if result.returncode != 0:
            util.print_stderr(f'command failed.[{result.returncode}]')
            return result.returncode
    except Exception as e:
        util.print_stderr(f'exception catch.[{e}]')
        return 99
    else:
        return 0


if __name__ == '__main__':
    sys.exit(main())
