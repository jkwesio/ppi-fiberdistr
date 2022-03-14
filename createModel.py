import gurobipy as gp
from gurobipy import GRB
import sys
import pandas as pd
import numpy as np
import streamlit as st
import emissionsCalcV4 as em

class createModel:

    def __init__(self, name, data, fProd):
        """
        create Gurobi model as decision variables, objective, and base constraints
        :param name: model name str, formatted as "FiberDistrModel_[name]"
        :param data: object from impData class with variables read in from spreadsheet
        :param fProd: products in system
        """
        self.fileName = data.fileName

        # bring in all the relevant data
        self.fProd = fProd
        self.rFiber = data.rFiber
        self.vFiber = data.vFiber
        self.rCat = data.rCat
        self.rPulp = data.rPulp
        self.vPulp = data.vPulp
        self.rLevel = data.rLevel

        self.wsource = data.wsource
        self.channelYield = data.channelYield
        self.wPaperYld = data.wPaperYld
        self.sCollectCost = data.sCollectCost
        self.sCollectOld = data.sCollectOld
        self.sCollectMax = data.sCollectMax
        self.productUse = data.productUse
        self.prodConsumed = data.prodConsumed

        self.rFCost = data.rFCost
        self.vFCost = data.vFCost
        self.rPCost = data.rPCost
        self.vPCost = data.vPCost
        self.rExCost = data.rExCost
        self.rExpOld = data.rExpOld

        self.rF2PYield = data.rF2PYield
        self.vF2PYield = data.vF2PYield
        self.f2pYld = data.fiber2pulpYield

        self.minPulp = data.minPulp
        self.rPCap = data.rPCap
        self.vPCap = data.vPCap
        self.maxrPulp = data.maxrPulp
        self.maxvPulp = data.maxvPulp

        self.recipeMin = data.recipeMin
        self.recipeMax = data.recipeMax
        self.rRecipeYield = data.rRecipeYield
        self.vRecipeYield = data.vRecipeYield
        self.rPYield = data.rPYield
        self.vPYield = data.vPYield

        self.minDemand = data.minDemand
        self.maxDemand = data.maxDemand
        self.oldDemand = data.oldDemand
        self.nonFiberPct = data.nonFiberPct
        self.nfYld = data.nfYld

        self.f2pVolOld = data.f2pVolOld
        self.pbpVolOld = data.pbpVolOld
        self.consCollOld = data.consCollOld
        self.exportOld = data.exportOld

        # flag if results exist
        self.solved = False
        self.failed = False

        # create model
        self.m = gp.Model(f"FiberDistrModel_{name}")

        # define decision variables
        arc_rp2p = [(p, r, t) for p in self.rPulp for t in self.fProd for r in self.rLevel[t]]
        arc_vp2p = [(q, r, t) for q in self.vPulp for t in self.fProd for r in self.rLevel[t]]
        arc_dem = [(r, t) for t in self.fProd for r in self.rLevel[t]]

        self.rfiber2pulp = self.m.addVars(self.rFiber, self.rPulp, lb=0, name='rFiber2Pulp')
        self.vfiber2pulp = self.m.addVars(self.vFiber, self.vPulp, lb=0, name='vFiber2Pulp')

        self.rpulpProd = self.m.addVars(self.rPulp, lb=0, name='rpulpProd')
        self.vpulpProd = self.m.addVars(self.vPulp, lb=0, name='vpulpProd')

        self.rpulp2prod = self.m.addVars(arc_rp2p, lb=0, name='rPulp2Prod')
        self.vpulp2prod = self.m.addVars(arc_vp2p, lb=0, name='vPulp2Prod')

        self.rExpNew = self.m.addVars(self.rFiber, lb=0, name='rExp')
        self.rResidue = self.m.addVars(self.rFiber, lb=0, name='rResidue')

        self.prodDemand = self.m.addVars(arc_dem, lb=0, name='prodDemand')
        self.prodFiberW = self.m.addVars(arc_dem, lb=0, name='prodFiberWeight')
        self.prodNonFW = self.m.addVars(arc_dem, lb=0, name='prodNonFiberWeight')

        self.sCollectNew = self.m.addVars(self.fProd, self.wsource, lb=0, name='sCollectNew')
        self.sCollectDelta = self.m.addVars(self.fProd, self.wsource, lb=0, name='sCollectDelta')

        # set objective function
        self.m.setObjective( # TODO: does the weighting still make sense?
            0.01 * gp.quicksum(
                self.rFCost[i] * gp.quicksum(self.rfiber2pulp[i, p] * 1000 for p in self.rPulp if 'Deinked' not in p)
                for i in self.rFiber)
            + 0.01 * gp.quicksum(
                self.vFCost[j] * gp.quicksum(self.vfiber2pulp[j, q] * 1000 for q in self.vPulp if 'Market' not in q)
                for j in self.vFiber)
            + 0.01 * gp.quicksum(
                self.rPCost[p] * gp.quicksum(
                    gp.quicksum(self.rpulp2prod[p, r, t] * 1000 for r in self.rLevel[t]) for t in self.fProd) for p in
                self.rPulp)
            + 0.01 * gp.quicksum(
                self.vPCost[q] * gp.quicksum(
                    gp.quicksum(self.vpulp2prod[q, r, t] * 1000 for r in self.rLevel[t]) for t in self.fProd) for q in
                self.vPulp)
            + 0.01 * gp.quicksum(
                self.sCollectCost[t, s] * self.sCollectDelta[t, s] * self.prodConsumed[t, s] for t in self.fProd for s in
                self.wsource)
            + 0.01 * gp.quicksum(
                (self.rExCost[i] - self.rFCost[i]) * (self.rExpOld[i] - self.rExpNew[i]) * 1000 for i in self.rFiber),
            GRB.MINIMIZE)

        # set base constraints
        # Mass balance
        self.m.addConstrs(
            (self.rF2PYield[p].dot(self.rfiber2pulp.select('*', p)) == self.rpulpProd[p] for p in self.rPulp), "rPulpinFlow")
        self.m.addConstrs(
            (self.vF2PYield[q].dot(self.vfiber2pulp.select('*', q)) == self.vpulpProd[q] for q in self.vPulp),
            "vPulpinFlow")

        self.m.addConstrs(
            (
            gp.quicksum(gp.quicksum(self.rpulp2prod[p, r, t] for r in self.rLevel[t]) for t in self.fProd) == self.rpulpProd[
                p]
            for p in self.rPulp), "rPulpoutFlow")
        self.m.addConstrs(
            (
            gp.quicksum(gp.quicksum(self.vpulp2prod[q, r, t] for r in self.rLevel[t]) for t in self.fProd) == self.vpulpProd[
                q]
            for q in self.vPulp), "vPulpoutFlow")

        # Market pulp consumption capacity (minimum)
        self.m.addConstrs(
            (gp.quicksum(self.rpulp2prod[p, r, t] for r in self.rLevel[t]) >= self.minPulp.loc[p, t]
             for p in self.rPulp for t in self.fProd if 'Deinked' in p), "dinkMin")
        self.m.addConstrs(
            (gp.quicksum(self.vpulp2prod[q, r, t] for r in self.rLevel[t]) >= self.minPulp.loc[q, t]
             for q in self.vPulp for t in self.fProd if 'Market' in q), "mpMin")

        # Pulp production capacity (maximum)
        self.m.addConstrs((self.rpulpProd[p] <= self.rPCap[p] for p in self.rPulp if 'Deinked' in p), "rPulpCap")
        self.m.addConstrs((self.vpulpProd[q] <= self.vPCap[q] for q in self.vPulp if 'Market' not in q),
                                 "vPulpCap")

        # Pulp production recycled content capacity (maximum)
        self.m.addConstrs(
            (self.rpulp2prod[p, r, t] <= self.maxrPulp[t].loc[r, p] for p in self.rPulp for t in self.fProd for r in
             self.rLevel[t]), "rpMaxUse")
        self.m.addConstrs(
            (self.vpulp2prod[q, r, t] <= self.maxvPulp[t].loc[r, q] for q in self.vPulp for t in self.fProd for r in
             self.rLevel[t]), "vpMaxUse")

        # Recipe
        self.m.addConstrs(
            (self.rF2PYield.loc[i, p] * self.rfiber2pulp[i, p] >= self.recipeMin.loc[i, p] * self.rpulpProd[p] for p in
             self.rPulp for i in
             self.rFiber),
            "rRecipeMin")
        self.m.addConstrs(
            (self.rF2PYield.loc[i, p] * self.rfiber2pulp[i, p] <= self.recipeMax.loc[i, p] * self.rpulpProd[p] for p in
             self.rPulp for i in
             self.rFiber),
            "rRecipeMax")
        self.m.addConstrs(
            (self.rF2PYield.loc[i, p] * self.rfiber2pulp[i, p] - self.rRecipeYield.loc[i, p] * self.rpulpProd[p] == 0
             for p in self.rPulp for i in self.rFiber
             if 'Deinked' in p or 'P&W' in p or 'News' in p), 'rRecipe')
        self.m.addConstrs(
            (self.vF2PYield.loc[j, q] * self.vfiber2pulp[j, q] == self.vRecipeYield.loc[j, q] * self.vpulpProd[q]
             for q in self.vPulp for j in
             self.vFiber),
            "vRecipe")

        # Wastepaper recovery
        self.m.addConstrs((self.sCollectNew[t, s] <= self.sCollectMax[t, s] for t in self.fProd for s in self.wsource),
                               "collMax")
        self.m.addConstrs((self.sCollectNew[t, s] >= self.sCollectOld[t, s] for t in self.fProd for s in self.wsource),
                               "collMin")
        self.m.addConstrs(
            (self.sCollectNew[t, s] - self.sCollectDelta[t, s] == self.sCollectOld[t, s] for t in self.fProd for s in
             self.wsource), "collDel")

        # Export
        self.m.addConstrs((self.rExpNew[i] <= self.rExpOld[i] for i in self.rFiber), "expUpper")

        # Recovery balance
        self.m.addConstrs((self.rfiber2pulp.sum(i, '*') + self.rExpNew[i] + self.rResidue[i] -
                                 gp.quicksum(gp.quicksum(
                                     self.wPaperYld[i, t, s] * self.sCollectNew[t, s] * self.prodConsumed[t, s] for s in
                                     self.wsource)
                                             for t in self.fProd) == 0 for i in self.rFiber), "rFiberAvail")

        self.m.addConstrs(
            (self.prodDemand.sum('*', t) == self.productUse.loc[t, 'Domestic Consumption'] for t in self.fProd),
            "constantDemand")
        self.m.addConstrs(
            (self.prodDemand[r, t] >= self.minDemand[t].loc[r] for t in self.fProd for r in self.rLevel[t]), "minDemand")

        self.m.addConstrs(
            (self.prodNonFW[r, t] == self.nonFiberPct[t].sum() * self.prodDemand[r, t] for t in self.fProd for r in
             self.rLevel[t]), "nfMass")
        self.m.addConstrs(
            (self.prodFiberW[r, t] == (1 - self.nfYld * self.nonFiberPct[t].sum()) * self.prodDemand[r, t]
             for t in self.fProd for r in self.rLevel[t]), "fPlusNF")

        self.m.addConstrs((gp.quicksum(self.rPYield[p] * self.rpulp2prod[p, r, t] for p in self.rPulp)
                                 + gp.quicksum(self.vPYield[q] * self.vpulp2prod[q, r, t] for q in self.vPulp) ==
                                 self.prodFiberW[r, t]
                                 for t in self.fProd for r in self.rLevel[t]), "bothBalance")
        self.m.addConstrs((gp.quicksum(self.rPYield[p] * self.rpulp2prod[p, r, t] for p in self.rPulp) - r *
                                 self.prodFiberW[r, t] == 0
                                 for t in self.fProd for r in self.rLevel[t]), "recBalance")
        self.m.addConstrs(
            (gp.quicksum(self.vPYield[q] * self.vpulp2prod[q, r, t] for q in self.vPulp) - (1 - r) * self.prodFiberW[
                r, t] == 0
             for t in self.fProd for r in self.rLevel[t]), "virBalance")
        # baseline -- no scenario
        self.rAvg = {t: self.oldDemand[t].dot(self.rLevel[t]) / self.oldDemand[t].sum() for t in self.fProd}

        self.m.addConstrs(
            (self.prodDemand[r, t] <= self.maxDemand[t].loc[r] for t in self.fProd for r in self.rLevel[t]
             if t not in self.fProd[-2:]), "maxDemand")
        self.m.addConstrs(
            (gp.quicksum(r * self.prodDemand[r, t] for r in self.rLevel[t]) == self.rAvg[t] * self.oldDemand[t].sum()
             for t in self.fProd if t in self.fProd[-2:]), "recFix")
        self.m.addConstrs(
            (gp.quicksum(r * self.prodDemand[r, t] for r in self.rLevel[t]) <= self.rAvg[t] * self.oldDemand[t].sum()
             for t in self.fProd if t not in self.fProd[-2:]), "recMax")

        self.m.update()

    def runModel(self):
        """
        update and run Gurobi optimization
        """
        self.m.update()
        self.m.optimize()

        if self.m.status == GRB.OPTIMAL:
            self.solved = True
            self.failed = False
        elif self.m.status != GRB.OPTIMAL:
            self.solved = False
            self.failed = True
            st.write('\nThe model is infeasible. Try changing the scenario settings.')
            # st.write('The model is infeasible; relaxing the constraints')
            # orignumvars = self.m.NumVars
            # self.m.feasRelaxS(0, False, False, True)
            # self.m.optimize()
            # status = self.m.status
            # if status in (GRB.INF_OR_UNBD, GRB.INFEASIBLE, GRB.UNBOUNDED):
            #     st.write('The relaxed model cannot be solved \
            #            because it is infeasible or unbounded')
            #     sys.exit(1)
            # if status != GRB.OPTIMAL:
            #     st.write('Optimization was stopped with status %d' % status)
            #     sys.exit(1)
            # st.write('\nSlack values:')
            # slacks = self.m.getVars()[orignumvars:]
            # for sv in slacks:
            #     if sv.X > 1e-6:
            #         st.write('%s = %g' % (sv.VarName, sv.X))

    def getResults(self):
        """
        save Gurobi model decision variables, convert from MM short ton to short ton, and calculate emissions data
        """
        fSetPulp = self.fProd + ['Market']
        fSetFiber = self.rFiber + self.vFiber
        self.f2pVolNew = pd.DataFrame(0, index=fSetFiber, columns=fSetPulp)

        fAllPulp = self.rPulp + self.vPulp
        self.pbpVolNew = pd.DataFrame(0, index=fAllPulp, columns=self.fProd)

        self.exportNew = pd.DataFrame(0, index=self.rFiber, columns=['exportNew'])
        self.rsdlNew = pd.Series(0, index=self.rFiber, name='residuals')
        self.demandNew = {}
        self.nfNew = {}
        self.fiberWNew = {}
        self.collDelta = pd.DataFrame(0, index=self.fProd, columns=self.wsource)
        self.collNew = pd.DataFrame(0, index=self.fProd, columns=self.wsource)
        self.addlRec = pd.Series(0, index=self.rFiber)

        for i in fSetFiber:
            if i in self.rFiber:
                self.f2pVolNew.loc[i] = self.m.getAttr('x', self.rfiber2pulp.select(i, '*'))
            else:
                self.f2pVolNew.loc[i] = self.m.getAttr('x', self.vfiber2pulp.select(i, '*'))

        for p in fAllPulp:
            if p in self.rPulp:
                self.pbpVolNew.loc[p] = [sum(self.m.getAttr('x', self.rpulp2prod.select(p, '*', t))) for t in self.fProd]
            else:
                self.pbpVolNew.loc[p] = [sum(self.m.getAttr('x', self.vpulp2prod.select(p, '*', t))) for t in self.fProd]

        for i in self.rFiber:
            self.exportNew.loc[i] = self.m.getAttr('x', self.rExpNew.select(i))

        for i in self.rFiber:
            self.rsdlNew[i] = np.unwrap(self.m.getAttr('x', self.rResidue.select(i)))

        for t in self.fProd:
            self.demandNew[t] = pd.Series(self.m.getAttr('x', self.prodDemand.select('*', t)), index=self.rLevel[t], name=t)
            self.nfNew[t] = pd.Series(self.m.getAttr('x', self.prodNonFW.select('*', t)), index=self.rLevel[t], name=t)
            self.fiberWNew[t] = pd.Series(self.m.getAttr('x', self.prodFiberW.select('*', t)), index=self.rLevel[t], name=t)

        for t in self.fProd:
            self.collDelta.loc[t] = self.m.getAttr('x', self.sCollectDelta.select(t, '*'))

        for t in self.fProd:
            self.collNew.loc[t] = self.m.getAttr('x', self.sCollectNew.select(t, '*'))

        for i in self.rFiber:
            self.addlRec.loc[i] = sum(
                [self.wPaperYld[i, t, s] * self.prodConsumed[t, s] * np.unwrap(self.m.getAttr('x', self.sCollectDelta.select(t, s)))
                 for t in self.fProd for s in self.wsource])

        ccIndex = self.consCollOld.index
        self.consCollNew = pd.DataFrame(0, index=ccIndex, columns=self.fProd)
        for t in self.fProd:
            self.consCollNew.loc[ccIndex[0],t] = self.productUse.loc[t,ccIndex[0]]
            self.consCollNew.loc[ccIndex[1], t] = self.productUse.loc[t, ccIndex[0]] * sum(
                self.productUse.loc[t][:-1].multiply(self.collNew.loc[t].values))
            self.consCollNew.loc[ccIndex[2], t] = self.productUse.loc[t, ccIndex[0]] * sum(self.productUse.loc[t][:-1].multiply(
                self.collNew.loc[t].values * self.channelYield.loc[0]))

        # convert to US tons (short tons) from MM short tons
        self.f2pVolNew = self.f2pVolNew * 1000;
        self.pbpVolNew = self.pbpVolNew * 1000;
        self.f2pVolOld = self.f2pVolOld * 1000;
        self.pbpVolOld = self.pbpVolOld * 1000
        self.exportNew = self.exportNew * 1000;
        self.exportOld = self.exportOld * 1000;
        self.rsdlNew = self.rsdlNew * 1000;
        self.addlRec = self.addlRec * 1000
        self.consCollOld = self.consCollOld * 1000
        self.consCollNew = self.consCollNew * 1000
        for t in self.demandNew.keys():
            self.demandNew[t] = self.demandNew[t] * 1000
            self.nfNew[t] = self.nfNew[t] * 1000
            self.fiberWNew[t] = self.fiberWNew[t] * 1000
            self.oldDemand[t] = self.oldDemand[t] * 1000

        # energy & emissions calculations
        p2pYld = self.rPYield.copy(); p2pYld.update(self.vPYield)
        p2pYld = pd.DataFrame.from_dict(p2pYld, orient='index', columns=['pYield'])
        em1 = em.en_emissions(self.fileName, self.fProd, self.rLevel, self.f2pYld, p2pYld, self.f2pVolNew, self.pbpVolNew, self.consCollNew, self.exportNew, self.demandNew)
        self.emissions = em1.calculateEmissions()

    def runSlack(self):
        """
        Gurobi workforce solution 3 to relax constraints & solve model
        :return: print relaxations to console
        """
        print('The model is infeasible; relaxing the constraints')
        orignumvars = self.m.NumVars
        self.m.feasRelaxS(0, False, False, True)
        self.m.optimize()
        status = self.m.status
        if status in (GRB.INF_OR_UNBD, GRB.INFEASIBLE, GRB.UNBOUNDED):
            print('The relaxed model cannot be solved \
                   because it is infeasible or unbounded')
            sys.exit(1)
        if status != GRB.OPTIMAL:
            print('Optimization was stopped with status %d' % status)
            sys.exit(1)
        print('\nSlack values:')
        slacks = self.m.getVars()[orignumvars:]
        for sv in slacks:
            if sv.X > 1e-6:
                print('%s = %g' % (sv.VarName, sv.X))
