import pandas as pd
import numpy as np
import gurobipy as gp


class impData:

    def __init__(self, file, fProd):
        """
        read in and clean data

        :param file: data spreadsheet to be read in
        :param fProd: list of strings for products in system
        """
        self.fileName = file
        # read in from spreadsheet
        with pd.ExcelFile(self.fileName) as x:
            self.fiberList = pd.read_excel(x, 'Fiber', usecols="B:G", skiprows=1, nrows=21)
            self.fiber2pulpYield = pd.read_excel(x, 'Fiber', usecols="I:O", skiprows=1, nrows=21)

            self.pulpList = pd.read_excel(x, 'Pulp', usecols="B:E", skiprows=1, nrows=14)
            self.minPulp = pd.read_excel(x, 'Pulp', usecols="H:M", skiprows=1, nrows=14)
            self.maxrPulp = {}
            self.maxvPulp = {}
            for i in range(len(fProd)):
                self.maxrPulp[fProd[i]] = pd.read_excel(x, 'Pulp', usecols="H:N", skiprows=19 + i * 8, nrows=5)
                self.maxvPulp[fProd[i]] = pd.read_excel(x, 'Pulp', usecols="Q:W", skiprows=19 + i * 8, nrows=5)

            self.oldDemand = pd.read_excel(x, 'Demand', usecols="F:K", skiprows=1, nrows=5)
            self.rLevel = pd.read_excel(x, 'Demand', usecols="F:K", skiprows=16, nrows=5)
            self.minDemand = pd.read_excel(x, 'Demand', usecols="N:S", skiprows=1, nrows=5)
            self.maxDemand = pd.read_excel(x, 'Demand', usecols="N:S", skiprows=10, nrows=5)

            self.channelYield = pd.read_excel(x, 'Recovery', usecols="C:F", skiprows=1, nrows=1)
            self.productUse = pd.read_excel(x, "Recovery", usecols="C:G", skiprows=5, nrows=6)
            self.colOldRate = pd.read_excel(x, "Recovery", usecols="C:G", skiprows=24, nrows=6)
            self.colMaxRate = pd.read_excel(x, "Recovery", usecols="C:G", skiprows=35, nrows=6)
            self.colCost = pd.read_excel(x, "Recovery", usecols="C:F", skiprows=44, nrows=6)
            self.colByProd = {}
            for i in range(len(fProd)):
                cols = [y + 4 * i for y in range(14, 18)]
                self.colByProd[fProd[i]] = pd.read_excel(x, 'Recovery', usecols=cols, skiprows=31,
                                                         nrows=16, header=None,
                                                         names=[x for x in self.channelYield.keys()])
            self.rConsExp = pd.read_excel(x, 'Recovery', usecols="U:V", skiprows=5, nrows=16)

            self.recipeYield = pd.read_excel(x, "Recipe", usecols="L:R", skiprows=1, nrows=21)
            self.recipeMin = pd.read_excel(x, "Recipe", usecols="T:Z", skiprows=1, nrows=16)
            self.recipeMax = pd.read_excel(x, "Recipe", usecols="T:Z", skiprows=20, nrows=16)

            self.nonFiberPct = pd.read_excel(x, "nonFiber", usecols="A:B,E:J", skiprows=2, nrows=3, index_col=0)

            self.f2pVolOld = pd.read_excel(x, 'OldData', usecols="A:H", skiprows=1, nrows=21, index_col=0)
            self.exportOld = pd.read_excel(x, 'OldData', usecols="E:F", skiprows=31, nrows=16, index_col=0)
            self.pbpVolOld = pd.read_excel(x, 'OldData', usecols="K:Q", skiprows=1, nrows=14, index_col=0)
            self.consCollOld = pd.read_excel(x, 'OldData', usecols="K:Q", skiprows=29, nrows=3, index_col=0)

        # break up into lists, tuplelists, and tupledicts
        self.rFiber, self.rCat, self.rFCost, self.rExCost = gp.multidict({k: [l, m, n]
                                                                          for k, l, m, n in
                                                                          zip((x for x in self.fiberList.iloc[:16, 1]),
                                                                              # Fiber Code
                                                                              (x for x in self.fiberList.iloc[:16, 2]),
                                                                              # Fiber Category
                                                                              (x for x in self.fiberList.iloc[:16, 4]),
                                                                              # Domestic Price
                                                                              (x for x in self.fiberList.iloc[:16,
                                                                                          5]))})  # Export Price
        self.rPulp, self.rPCost, self.rPYield, self.rPCap = gp.multidict({k: [l, m, n]
                                                                          for k, l, m, n in
                                                                          zip((x for x in self.pulpList.iloc[:7, 0]),
                                                                              # Pulp Name
                                                                              (x for x in self.pulpList.iloc[:7, 1]),
                                                                              # Price ($/short ton)
                                                                              (x for x in self.pulpList.iloc[:7, 2]),
                                                                              # Yield
                                                                              (x for x in
                                                                               self.pulpList.iloc[:7, 3]))})  # Capacity
        self.vFiber, self.vFCost = gp.multidict({k: l
                                                 for k, l in zip((x for x in self.fiberList.iloc[16:, 1]),  # Fiber Code
                                                                 (x for x in
                                                                  self.fiberList.iloc[16:, 4]))})  # Domestic Price
        self.vPulp, self.vPCost, self.vPYield, self.vPCap = gp.multidict({k: [l, m, n]
                                                                          for k, l, m, n in
                                                                          zip((x for x in self.pulpList.iloc[7:, 0]),
                                                                              # Pulp Name
                                                                              (x for x in self.pulpList.iloc[7:, 1]),
                                                                              # Price ($/short ton)
                                                                              (x for x in self.pulpList.iloc[7:, 2]),
                                                                              # Yield
                                                                              (x for x in
                                                                               self.pulpList.iloc[7:, 3]))})  # Capacity

        # reset df indexes for consistent labeling
        for i in fProd:
            self.colByProd[i].index = self.rFiber
        self.rConsExp.index = self.rFiber
        self.recipeMin.index = self.rFiber
        self.recipeMax.index = self.rFiber

        self.recipeYield.index = self.rFiber + self.vFiber
        self.fiber2pulpYield.index = self.rFiber + self.vFiber

        self.minPulp.index = self.rPulp + self.vPulp

        self.productUse.index = fProd
        self.colOldRate.index = fProd
        self.colMaxRate.index = fProd
        self.colCost.index = fProd

        # reset df columns for consistent labeling (& remove duplicate exception in column names)
        for i in fProd:
            self.maxrPulp[i].columns = self.rPulp
            self.maxvPulp[i].columns = self.vPulp
        self.recipeMax.columns = self.rPulp
        self.recipeMin.columns = self.rPulp

        self.minDemand.columns = fProd
        self.maxDemand.columns = fProd

        self.recipeYield.columns = fProd + ['Market']

        self.pbpVolOld.columns = fProd
        self.consCollOld.columns = fProd

        # pricing & setting up variables for ease of use
        self.rF2PYield = self.fiber2pulpYield.iloc[:16, :]
        self.rF2PYield.columns = self.rPulp
        self.vF2PYield = self.fiber2pulpYield.iloc[16:, :]
        self.vF2PYield.columns = self.vPulp

        self.rRecipeYield = self.recipeYield.iloc[:16, :]
        self.rRecipeYield.columns = self.rPulp
        self.vRecipeYield = self.recipeYield.iloc[16:, :]
        self.vRecipeYield.columns = self.vPulp

        self.rExpOld = {i: self.rConsExp.loc[i, 'Export'] for i in self.rFiber}

        self.wsource = [x for x in self.channelYield.keys()]
        self.prodConsumed = {(t, s): self.productUse.loc[t, s] * self.productUse.loc[t, 'Domestic Consumption']
                             for t in self.productUse.index for s in self.wsource}
        self.sCollectCost = {(t, s): self.colCost.loc[t, s] for t in fProd for s in self.wsource}
        self.sCollectOld = {(t, s): self.colOldRate.loc[t, s] for t in fProd for s in self.wsource}
        self.sCollectMax = {(t, s): self.colMaxRate.loc[t, s] for t in fProd for s in self.wsource}
        self.wPaperYld = {(i, t, s): self.colByProd[t].loc[i, s] * self.channelYield[s].values[0] for i in self.rFiber
                          for t in fProd for s in self.wsource}

        self.nfYld = 0.99

        # rLevel for each product's different recycled content level (df to dict)
        self.rLevel = {t: list(self.rLevel[t][np.isfinite(self.rLevel[t])].values) for t in fProd}

        self.oldDemand = {t: self.oldDemand[t] for t in fProd}
        self.minDemand = {t: self.minDemand[t] for t in fProd}
        self.maxDemand = {t: self.maxDemand[t] for t in fProd}

        for t in fProd:
            self.maxrPulp[t] = self.maxrPulp[t].iloc[:len(self.rLevel[t]), :]
            self.maxvPulp[t] = self.maxvPulp[t].iloc[:len(self.rLevel[t]), :]
            self.oldDemand[t] = self.oldDemand[t].iloc[:len(self.rLevel[t])]
            self.minDemand[t] = self.minDemand[t].iloc[:len(self.rLevel[t])]
            self.maxDemand[t] = self.maxDemand[t].iloc[:len(self.rLevel[t])]

            self.maxrPulp[t].index = self.rLevel[t]
            self.maxvPulp[t].index = self.rLevel[t]
            self.oldDemand[t].index = self.rLevel[t]
            self.minDemand[t].index = self.rLevel[t]
            self.maxDemand[t].index = self.rLevel[t]
