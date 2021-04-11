import numpy as np
import matplotlib.pyplot as plt

from db_util import DBUtil


class GraphUtil():
    def __init__(self, results: dict):
        self._results = results

    def show(self):
        left = np.array(self._results[DBUtil.CLM_DATETIME])
        height = np.array(self._results[DBUtil.CLM_CORRECT_RATE])
        plt.plot(left, height)
        plt.gcf().autofmt_xdate()
        plt.hlines([65], left[0], left[-1], "blue", linestyles='dashed')
        plt.show()
