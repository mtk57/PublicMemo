from abc import ABCMeta, abstractclassmethod

import const as const
from const import Mode
from util import Util


class CollectInfo():
    def __init__(self, info_file_path: str):
        self._info_file_path = info_file_path

    def __repr__(self):
        return f'[{self.info_file_path}]'

    @property
    def info_file_path(self) -> str:
        return self._info_file_path


class QuizInfo():
    def __init__(self, mode: Mode = Mode.QUIZ):
        self._num = None        # 番号
        self._question = None   # 問題
        self._choises = []      # 選択肢
        self._answers = []      # 答え
        self._mode = mode

    def __repr__(self):
        return f'[{self.num}]{self.question} :choices=[{self.choises}], ' \
               f'answers=[{self.answers}]'

    @property
    def num(self) -> int:
        return self._num

    @num.setter
    def num(self, num: int):
        self._num = num

    @property
    def question(self) -> str:
        return self._question

    @question.setter
    def question(self, question: str):
        self._question = question

    def add_choice(self, choice: str):
        self._choises.append(choice)

    @property
    def choices(self) -> list:
        return self._choises

    def add_answer(self, answer: str):
        self._answers.append(answer)

    @property
    def answers(self) -> list:
        return self._answers

    @property
    def answer_nums(self) -> list:
        ret = []
        index = 1
        for answer in self.answers:
            if answer == const.MARK_CORRECT:
                ret.append(index)
            index += 1
        return ret

    @property
    def answer_choices(self) -> list:
        ret = []
        for num in self.answer_nums:
            ret.append(self.choices[num - 1])
        return ret

    @property
    def answer_count(self) -> int:
        ret = 0
        for answer in self.answers:
            if answer == const.MARK_CORRECT:
                ret += 1
        return ret

    @property
    def mode(self) -> Mode:
        return self._mode

    def show(self):
        print('-----------------------------------------')
        print(f'{self.question}')
        print('-----------------------------------------')

        if self.mode == Mode.QUIZ:
            num = 1
            for choice in self.choices:
                print(f'{num}:{choice}')
                num += 1
        else:
            for choice in self.answer_choices:
                print(f'{choice}')
        print('')

    def show_answer(self):
        for choice in self.answer_choices:
            print(f'{choice}')


class ResultInfo():
    def __init__(self, num: int, is_right: bool = False, mode: Mode = Mode.QUIZ):
        self._num = num
        self._is_right = is_right
        self._mode = mode

    def __repr__(self):
        return f'{self.num}:{self.is_right}:mode=[{self.mode}]'

    @property
    def num(self) -> int:
        return self._num

    @property
    def is_right(self) -> bool:
        return self._is_right

    @property
    def mode(self) -> Mode:
        return self._mode


class CollectorBase(metaclass=ABCMeta):
    def __init__(self, logger, collect_info: CollectInfo):
        self._logger = logger
        self._collect_info = collect_info
        self._collections = []

    @abstractclassmethod
    def start_collection(self):
        pass

    def add_collection(self, data: object):
        self._collections.append(data)

    def get_collection(self) -> list:
        return self._collections

    def get_random_collection(self) -> list:
        return Util.get_random_list(self._collections)

    @property
    def logger(self) -> object:
        return self._logger

    @property
    def collect_info(self) -> CollectInfo:
        return self._collect_info
