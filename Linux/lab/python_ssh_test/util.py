#!/usr/bin/env python3
import os
import paramiko
import socket
import subprocess
import sys


def get_hostname():
    return os.uname()[1]


def run_command(cmd: list):
    result = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL)
    return result


def print_stderr(message: str):
    print(message, file=sys.stderr)


class SshConnectError(Exception):
    def __init__(self, message: str):
        self.message = message


class SshExecCommandError(Exception):
    def __init__(self, message: str):
        self.message = message


class SshTimeoutError(Exception):
    def __init__(self, message: str):
        self.message = message


class SshCommandModel():
    def __init__(self, ip: str, user: str, password: str, connect_timeout: int, command_timeout: int, command: str):
        self.ip = ip
        self.user = user
        self.password = password
        self.connect_timeout = connect_timeout
        self.command_timeout = command_timeout
        self.command = command

    def __repr__(self):
        return f'ip={self.ip}, user={self.user}, pw={self.password}, con_to={self.connect_timeout}, cmd_to={self.command_timeout}, cmd={self.command}'


class CommandResult():
    def __init__(self, ret_code: int, stdout: list, stderr: list):
        self.ret_code = ret_code
        self.stdout = stdout
        self.stderr = stderr


def ssh_run_command(ssh_model: SshCommandModel) -> CommandResult:
    ret = None
    client = None
    try:
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.WarningPolicy())

        try:
            client.connect(ssh_model.ip, username=ssh_model.user, password=ssh_model.password,
                           timeout=ssh_model.connect_timeout)
        except paramiko.BadHostKeyException as e:
            raise SshConnectError(f'BadHostKeyException! [{e}]')
        except paramiko.AuthenticationException as e:
            raise SshConnectError(f'AuthenticationException! [{e}]')
        except paramiko.SSHException as e:
            raise SshConnectError(f'SSHException! [{e}]')
        except socket.error as e:
            # タイムアウトも含む
            raise SshConnectError(f'socket.error! [{e}]')
        try:
            stdin, stdout, stderr = client.exec_command(
                command=ssh_model.command, timeout=ssh_model.command_timeout)

            return CommandResult(
                ret_code=stdout.channel.recv_exit_status(),
                stdout=[f for f in stdout],
                stderr=[f for f in stderr])

        except paramiko.SSHException as e:
            raise SshExecCommandError(f'SSHException! [{e}]')
        except socket.timeout as e:
            raise SshTimeoutError(f'Command timeout! [{e}]')
    finally:
        if client:
            client.close()
