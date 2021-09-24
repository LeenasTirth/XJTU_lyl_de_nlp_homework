import numpy as np
import pandas as pd
import xlwt
import xlrd
from xlutils.copy import copy


class Elem:
    def __init__(self, ntimes, index, parentindex):
        self.mnTimes = ntimes  # value
        self.index = index
        self.parent = parentindex  # the index (-1,-1) means this obj is origin.

    def GetIndex(self):
        return self.index

    def GetParent(self):
        return self.parent

    def GetValues(self):
        return self.mnTimes


class EditDistance:
    def __init__(self, Astr='', Bstr=''):
        self.msAstr = '#' + Astr
        self.msBstr = '#' + Bstr

        # Bstr is longer.
        if len(self.msAstr) > len(self.msBstr):
            self.msAstr, self.msBstr = self.msBstr, self.msAstr
        self.mElemMat = np.zeros((len(self.msAstr), len(self.msBstr))).tolist()
        self.rows = len(self.msAstr)
        self.cols = len(self.msBstr)
        self.mResultMat = None
        self.mMinPath = []
        self.insertcost = 1
        self.changecost = 2
        self.deletcost = 1

    def GetStr(self, Astr, Bstr):
        self.msAstr = '#' + Astr
        self.msBstr = '#' + Bstr
        # Bstr is longer.
        if len(self.msAstr) > len(self.msBstr):
            self.msAstr, self.msBstr = self.msBstr, self.msAstr
        self.mElemMat = np.zeros((len(self.msAstr), len(self.msBstr))).tolist()
        self.rows = len(self.msAstr)
        self.cols = len(self.msBstr)

    def __Min(self, i, j):
        parent = (-1, -1)
        values = 0
        # which = -1  # 0->left,1->diag,2->down
        flag = self.msAstr[i] != self.msBstr[j]

        leftvalues = self.mElemMat[i][j - 1].GetValues() + self.insertcost
        downvalues = self.mElemMat[i - 1][j].GetValues() + self.deletcost
        diagvalues = self.mElemMat[i - 1][j - 1].GetValues() + self.changecost * flag

        if diagvalues <= leftvalues and diagvalues <= downvalues:
            values = diagvalues
            parent = (i - 1, j - 1)
        elif leftvalues <= diagvalues and leftvalues <= downvalues:
            values = leftvalues
            parent = (i, j - 1)
        else:
            values = downvalues
            parent = (i - 1, j)
        return values, parent

    def __Initialize(self):
        self.mElemMat[0][0] = Elem(0, (0, 0), (-1, -1))
        for i in range(1, self.rows):
            self.mElemMat[i][0] = Elem(i, (i, 0), (i - 1, 0))
        for j in range(1, self.cols):
            self.mElemMat[0][j] = Elem(j, (0, j), (0, j - 1))

    def __UseDP(self):
        for i in range(1, self.rows):
            for j in range(1, self.cols):
                index = (i, j)
                values, parent = self.__Min(i, j)
                self.mElemMat[i][j] = Elem(values, index, parent)

    def ShowMinPath(self):
        print('Min Path:')
        for i in range(len(self.mMinPath)):
            print('index:', self.mMinPath[i][0], '   value:', self.mMinPath[i][1])

    def __GetMinPath(self):
        index = self.mElemMat[self.rows - 1][self.cols - 1].GetIndex()
        value = self.mElemMat[self.rows - 1][self.cols - 1].GetValues()
        elem = self.mElemMat[self.rows - 1][self.cols - 1]
        while index != (-1, -1):
            self.mMinPath.append((index, value))
            index = elem.GetParent()
            if index != (-1, -1):
                elem = self.mElemMat[index[0]][index[1]]
                value = elem.GetValues()
        self.mMinPath = self.mMinPath[::-1]

    def __GetResultMat(self):
        self.mResultMat = np.zeros((self.rows, self.cols))
        for i in range(self.rows):
            for j in range(self.cols):
                self.mResultMat[i][j] = self.mElemMat[i][j].GetValues()

    def WriteToExcel(self, filepath):
        out_df = pd.DataFrame(self.mResultMat)
        out_df.index = list(self.msAstr)
        out_df.columns = list(self.msBstr)
        out_df.to_excel(filepath)

    def Process(self):
        self.__Initialize()
        self.__UseDP()
        self.__GetMinPath()
        self.__GetResultMat()


if __name__ == '__main__':
    ED = EditDistance()
    ED.GetStr('execution', 'intention')
    # ED.GetStr('sdasfawfadasda','skadjkwjdlmfewk')
    ED.Process()
    ED.WriteToExcel(r'Result.xlsx')
    ED.ShowMinPath()
    # ED.ColorExcel(r'Result.xlsx')
