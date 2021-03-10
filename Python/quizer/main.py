import os
import sys
import argparse

import const as const
from const import Mode
from util import Util
from logger import Logger
from quizer import Quizer


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
    def mode(self) -> Mode:
        return Main.MODE_MAP[self._args.mode]

    @property
    def mode_str(self) -> str:
        return self._args.mode

    @property
    def is_show_answer(self) -> bool:
        return self._args.show_answer

    def parse_args(self):
        fn = 'parse_args'
        self._logger.DEBUG(f'{fn} S')

        parser = argparse.ArgumentParser()
        parser.add_argument('--info_path')
        parser.add_argument('--mode', choices=[Main.MODE_QUIZ, Main.MODE_LEARN])
        parser.add_argument('--show_answer', action='store_true')
        self._args = parser.parse_args()
        if not self._args.info_path:
            self._args.info_path = const.DEFAULT_EXCEL_FILE_NAME
        if not self._args.mode:
            self._args.mode = Main.MODE_QUIZ

        self._logger.DEBUG(f'{fn} E')

    def run(self):
        fn = 'run'
        self._logger.DEBUG(f'{fn} S')

        title = """\
***********************************************************
*                Quizerへようこそ!                         *
***********************************************************
* {0}
* mode={1}
* show answer={2}
*
* <使い方>
* [クイズモード]
*  - 回答は半角数字のみです。
*  - 複数の場合は半角カンマで区切って下さい。(例：1,2,3)
* [学習モード]
*  - エンターキーで次の問題と回答を表示します。
*
* [その他]
*  - 途中で終了する場合は、qを入力して下さい。
***********************************************************
"""
        print(title.format(const.VERSION, self.mode_str, self.is_show_answer))

        quizer = Quizer(logger=self._logger, info_path=self.info_path,
                        mode=self.mode)

        incorrects = []
        total_cnt = len(quizer.get_random_quiz_list())
        num = 0

        for quiz in quizer.get_random_quiz_list():
            num += 1
            print(f'【{num}/{total_cnt}】')
            quiz.show()

            # キー入力待ち
            input_answers = input('＞：').split(const.MARK_COMMA)
            if input_answers[0] == 'q':
                break

            if self.mode == Mode.LEARN:
                print('')
                continue

            result = Quizer.verify(quiz_info=quiz,
                                   input_answer=input_answers)

            if result.is_right:
                print('------------')
                print('正解です!!')
                print('------------')
            else:
                print('------------')
                print('不正解です...')
                print('------------')
                if self.is_show_answer:
                    quiz.show_answer()
                incorrects.append(result)
            print('')

        if self.mode == Mode.QUIZ:
            if len(incorrects) == 0:
                print('----------------')
                print('全問正解です!!!')
                print('----------------')
            else:
                print('-----------------------------')
                print(f'{total_cnt}問中, {len(incorrects)}問が不正解です...orz')
                print('-----------------------------')
                for incorrect in incorrects:
                    print(f'# {incorrect.num}')

        print('')
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
            main.run()

    except Exception as ex:
        logger.ERROR(Util.get_exception_message(ex))
        ret = 1

    finally:
        logger.DEBUG('END')
        logger.DEBUG('====================================================')
        sys.exit(ret)
