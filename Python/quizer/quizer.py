import const as const
from const import Offset, Mode
from util import Util
from data_models import CollectorBase, CollectInfo, QuizInfo, ResultInfo


class QuizCollector(CollectorBase):
    def __init__(self, logger, collect_info: CollectInfo,
                 mode: Mode = Mode.QUIZ):
        CollectorBase.__init__(self, logger, collect_info)
        self._book = None
        self._sheets = {}
        self._mode = mode

    def start_collection(self):
        fn = 'start_collection'
        self.logger.DEBUG(f'{fn} S')

        self._collect()

        self.logger.DEBUG(f'{fn} E')

    def _collect(self):
        fn = '_collect'
        self.logger.DEBUG(f'{fn} S')

        self._book = Util.read_excel(path=self.collect_info.info_file_path)

        for sheet_name in const.REQUIRED_SHEETS:
            if sheet_name not in self._book:
                raise Exception(f'必須シートがありません! expect=[{sheet_name}]')
            self._sheets[sheet_name] = self._book[sheet_name]

        self._verify_version()

        self._collect_quiz_from_excel_sheet()

        self.logger.DEBUG(f'{fn} E')

    def _verify_version(self):
        fn = '_verify_version'
        self.logger.DEBUG(f'{fn} S')

        ver = self._sheets[const.SHEET_COMMON][const.VER_POS]
        if ver is None or ver.value is None or ver.value != const.VERSION:
            raise Exception(f'バージョンが合っていません! expect=[{const.VERSION}], '
                            f'pos=[{const.VER_POS}]')

        self.logger.DEBUG(f'{fn} E')

    def _collect_quiz_from_excel_sheet(self):
        fn = '_collect_quiz_from_excel_sheet'
        self.logger.DEBUG(f'{fn} S')

        sheet = self._sheets[const.SHEET_QUIZ_ADMIN]
        q = QuizInfo(mode=self._mode)
        for row in sheet.iter_rows(min_row=const.OFFSET_ADMIN):
            for cell in row:
                v = cell.value
                if v is None:
                    continue

                if cell.column == Offset.NUM:
                    if q.num != v:
                        q = QuizInfo(mode=self._mode)
                        q.num = v
                        self.add_collection(q)
                elif cell.column == Offset.QUESTION:
                    q.question = v
                elif cell.column == Offset.CHOICE:
                    q.add_choice(v)
                elif cell.column == Offset.ANSWER:
                    q.add_answer(v)
                else:
                    continue

        self.logger.DEBUG(f'{fn} E')


class Quizer():
    def __init__(self, logger, info_path: str, mode: Mode = Mode.QUIZ):
        self._logger = logger
        self._info_path = info_path
        self._mode = mode
        self._quiz_collector = None

        self._collect_quiz()

    def get_random_quiz_list(self) -> list:
        return self._quiz_collector.get_random_collection()

    @classmethod
    def verify(cls, quiz_info: QuizInfo,
               input_answer: list) -> ResultInfo:
        num = quiz_info.num
        input_len = len(input_answer)
        NG = ResultInfo(num=num)

        if Quizer._verify_answer_range(input_len) is False:
            return NG

        if quiz_info.answer_count != input_len:
            return NG

        answers = []
        for answer in input_answer:
            if answer.isdecimal() is False:
                return NG
            answer_int = int(answer)
            if Quizer._verify_answer_range(answer_int) is False:
                return NG
            answers.append(answer_int)
        sort_answers = sorted(answers)

        if quiz_info.answer_nums != sort_answers:
            return NG

        return ResultInfo(num=num, is_right=True)

    @classmethod
    def _verify_answer_range(cls, answer: int) -> bool:
        if answer < const.MIN_ANSWER or \
           answer > const.MAX_ANSWER:
            return False
        return True

    @property
    def mode(self) -> Mode:
        return self._mode

    def _collect_quiz(self):
        ci = CollectInfo(info_file_path=self._info_path)
        self._quiz_collector = QuizCollector(logger=self._logger,
                                             collect_info=ci, mode=self._mode)
        self._quiz_collector.start_collection()