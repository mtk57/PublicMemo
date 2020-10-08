#!/usr/bin/env python3
import os
import subprocess
import glob
import re
import traceback

"""
gemファイルのバージョン比較
"""


def main():
    SEARCH_DIR = '/opt/td-agent/embedded/lib/ruby/gems/2.4.0/cache'
    GEM_LIST = ['td-agent-gem', 'list']

    def run_command(cmd: list):
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL)
        return result

    try:
        install_gems_tmp = [os.path.basename(f) for f in glob.glob(
            os.path.join(SEARCH_DIR, '*.gem'))]
        install_gems = [GemFileModel(f) for f in install_gems_tmp]

        result = run_command(GEM_LIST)
        gem_list = [GemListModel(f)
                    for f in result.stdout.decode().splitlines()]
        for gem in gem_list:
            for ins_gem in install_gems:
                if gem.name != ins_gem.name:
                    continue
                # print(f'gem={gem}, ins_gem={ins_gem}')
                ver1 = gem.normailze()
                ver2 = ins_gem.normailze()

                if ver1 == ver2:
                    print(f'{ver1} == {ver2}')
                elif ver1 > ver2:
                    print(f'{ver1} > {ver2}')
                else:
                    print(f'{ver1} < {ver2}')

    except Exception as e:
        print(f'Exception!! [{e}]')
        print(traceback.format_exc())


class GemModelBase():
    def __init__(self, name: str):
        self._org_name = name
        self._name = self._create_name()
        self._version = self._create_version()

    def __repr__(self):
        return f'org:{self._org_name}, name={self.name}, ver={self.version}, full={self.full_name}'

    def _create_name(self):
        pass

    def _create_version(self):
        pass

    @property
    def name(self) -> str:
        return self._name

    @property
    def version(self) -> str:
        return self._version

    @property
    def full_name(self) -> str:
        """
        @retval Ex.'bson (4.10.0)'
        """
        return f'{self.name} ({self.version})'

    def normailze(self) -> list:
        """
        @retval 410
        @note https://www.366service.com/jp/qa/2fdd2cc7c1ac56dd8b76e52aeb85d0f9
        """
        # '.0'で終わる場合は削除する(0は1つ以上)
        return [int(x) for x in re.sub(r'(\.0+)*$', '', self.version).split('.')]


class GemFileModel(GemModelBase):
    """
    Ex.'bson-4.10.0.gem'
    """

    def _create_name(self) -> str:
        """
        @retval Ex.'bson'
        """
        return self._org_name[:self._org_name.rfind('-')]

    def _create_version(self) -> str:
        """
        @retval Ex.'4.10.0'
        """
        return self._org_name[self._org_name.rfind('-')+1:].replace('.gem', '')


class GemListModel(GemModelBase):
    """
    Ex.'bson (4.10.0)'
    """

    def _create_name(self) -> str:
        """
        @retval Ex.'bson'
        """
        return self._org_name[:self._org_name.rfind('(')-1]

    def _create_version(self) -> str:
        """
        @retval Ex.'4.10.0'
        """
        return self._org_name[self._org_name.rfind('(')+1:].replace(')', '')


main()
