import os
import sys
import argparse
import time
import math
import re

import const as const
from const import Mode
from util import Util
from db_util import DB_Util
from logger import Logger
from quizer import Quizer

"""
サポート予定
・統計情報DBを利用して、正答率の低い問題を優先的に出題する。
・間違えた問題の復習モード

"""


class Main():
    MODE_QUIZ = 'quiz'
    MODE_LEARN = 'learn'

    MODE_MAP = {
        MODE_QUIZ: Mode.QUIZ,
        MODE_LEARN: Mode.LEARN
    }

    def __init__(self, logger: object):
        self._logger = logger
        self._args = None

    @property
    def info_path(self) -> str:
        return self._args.info_path

    @property
    def db_path(self) -> str:
        return self._args.db_path

    @property
    def mode(self) -> Mode:
        return Main.MODE_MAP[self._args.mode]

    @property
    def mode_str(self) -> str:
        return self._args.mode

    @property
    def num_of_questions(self) -> int:
        return self._args.num

    @property
    def pass_line(self) -> int:
        return self._args.pass_line

    @property
    def is_show_answer(self) -> bool:
        return self._args.show_answer

    @property
    def is_random(self) -> bool:
        return self._args.random

    def open(self):
        self._db_util = DB_Util(self.db_path)
        if self._db_util.open() is False:
            raise Exception('DB open failed')

    def rollback(self):
        self._db_util.rollback()

    def close(self):
        self._db_util.close()

    def clear(self):
        self._db_util.clear()

    def update(self, question_num: int, is_incorrect: bool):
        self._db_util.update(question_num, is_incorrect)

    def parse_args(self):
        fn = 'parse_args'
        self._logger.DEBUG(f'{fn} S')

        parser = argparse.ArgumentParser()
        parser.add_argument('--info_path')
        parser.add_argument('--db_path')
        parser.add_argument('--mode', choices=[Main.MODE_QUIZ, Main.MODE_LEARN])
        parser.add_argument('--show_answer', action='store_true')
        parser.add_argument('--num', type=int)
        parser.add_argument('--pass_line', type=int)
        parser.add_argument('--random', action='store_true')
        self._args = parser.parse_args()
        if not self._args.info_path:
            self._args.info_path = const.DEFAULT_EXCEL_FILE_NAME
        if not self._args.db_path:
            self._args.db_path = const.DEFAULT_DB_FILE_NAME
        if not self._args.mode:
            self._args.mode = Main.MODE_QUIZ
        if not self._args.num:
            self._args.num = 65535
        if not self._args.pass_line:
            self._args.pass_line = const.DEFAULT_PASS_LINE

        self._logger.DEBUG(f'{fn} E')

    def run(self):
        fn = 'run'
        self._logger.DEBUG(f'{fn} S')

        quizer = Quizer(logger=self._logger, info_path=self.info_path,
                        mode=self.mode)

        incorrects = []
        total_cnt = min([len(quizer.get_quiz_list()), self.num_of_questions])
        pass_line = min([100, self.pass_line])

        title = """\
***********************************************************
*                Quizerへようこそ!                         *
***********************************************************
* {0}
* mode={1}
* show answer={2}
* num={3}
* pass={4}%
* random={5}
* info_path={6}
* db_path={7}
*
* <使い方>
* [クイズモード]
*  - 回答は半角数字のみです。
*  - 複数の場合は半角SP で区切って下さい。(例：1 2 3)
* [学習モード]
*  - エンターキーで次の問題と回答を表示します。
*
* [その他]
*  - 終了する場合は、qを入力して下さい。
*  - 終了時に統計情報をクリアする場合は、cを入力して下さい。
***********************************************************
"""
        print(title.format(
                const.VERSION,
                self.mode_str,
                self.is_show_answer,
                total_cnt,
                pass_line,
                self.is_random,
                self.info_path,
                self.db_path
                ))

        is_db_clear = False
        num = 0
        correct_cnt = 0

        start = time.time()

        quizs = quizer.get_quiz_list()
        if self.is_random:
            quizs = quizer.get_random_quiz_list()

        for quiz in quizs:
            num += 1
            if num > total_cnt:
                break
            print(f'【{num}/{total_cnt}】')
            quiz.show()

            if self._db_util.is_exist():
                # 正答率を表示
                correct_rate = self._db_util.get_correct_rate(quiz.num)
                print(f'過去の正答率={correct_rate}%')

            # キー入力待ち
            input_answers = re.split(const.SPLITS, input('＞：'))
            if input_answers[0] == 'q':
                break
            if input_answers[0] == 'c':
                is_db_clear = True
                break

            if self.mode == Mode.LEARN:
                print('')
                continue

            result = Quizer.verify(quiz_info=quiz,
                                   input_answer=input_answers)

            is_incorrect = True
            if result.is_right:
                print('------------')
                print('正解です!!')
                print('------------')
                correct_cnt += 1
            else:
                print('------------')
                print('不正解です...')
                print('------------')
                if self.is_show_answer:
                    quiz.show_answer()
                incorrects.append(result)
                is_incorrect = False

            if self.update(quiz.num, is_incorrect) is False:
                raise Exception('DB update failed!')

            print('')

        # 後処理

        end = time.time()
        total_sec = end - start

        if self.mode == Mode.QUIZ:
            if len(incorrects) == 0:
                print('----------------')
                print('全問正解です!!!')
                print('----------------')
            else:
                print('-----------------------------')
                print(f'{num-1}問中, {correct_cnt}問が正解、{len(incorrects)}問が不正解でした。')
                print('-----------------------------')
                for incorrect in incorrects:
                    print(f'# {incorrect.num}')

            correct_rate = math.floor((correct_cnt / total_cnt) * 100)
            print(f'正解率={correct_rate}%')
            if int(correct_rate) >= pass_line:
                print('合格です！！')
            else:
                print('不合格です。。。orz')

            if is_db_clear:
                if self.clear() is False:
                    raise Exception('DB clear failed!')
                print('DBをクリアしました。')

        print('')
        print(f'所要時間：{Util.get_hhmmss_str_from_sec(total_sec)}')
        print('お疲れ様でした!!')
        print('')

        self._logger.DEBUG(f'{fn} E')


if __name__ == '__main__':
    ret = 0
    logger = Logger()

    # for DEBUG >>
    # sys.argv.append('--info_path')
    # sys.argv.append(const.DEFAULT_EXCEL_FILE_NAME)
    # sys.argv.append('--mode')
    # sys.argv.append(Main.MODE_LEARN)
    # for DEBUG <<

    try:
        logger.DEBUG('====================================================')
        logger.DEBUG('START')

        Util.change_current_dir()

        main = Main(logger)
        main.parse_args()

        if os.path.exists(main.info_path) is False:
            logger.ERROR(f'Input file not exist! [{main.info_path}]')
            ret = 1
        else:
            main.open()

            main.run()

    except Exception as ex:
        main.rollback()
        logger.ERROR(Util.get_exception_message(ex))
        ret = 1

    finally:
        main.close()
        logger.DEBUG('END')
        logger.DEBUG('====================================================')
        sys.exit(ret)
