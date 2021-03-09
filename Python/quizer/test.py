from util import Util


class quiz():
    def __init__(self):
        self.num = None
        self.q = None
        self.choice = []
        self.a = []

    def __repr__(self):
        return f'{self.num}, {self.q}, {self.choice}, {self.a}'

quizes = []


PATH = r'sample.xlsx'
SHEET = r'QUIZ'

book = Util.read_excel(PATH)
sheet = book[SHEET]

q = quiz()
for row in sheet.iter_rows(min_row=4):
    for cell in row:
        v = cell.value
        if v is None:
            continue

        if cell.column == 1:
            if q.num != v:
                q = quiz()
                q.num = v
                quizes.append(q)

        elif cell.column == 2:
            q.q = v
        elif cell.column == 3:
            q.choice.append(v)
        elif cell.column == 4:
            q.a.append(v)
        else:
            continue

print(quizes)

pass