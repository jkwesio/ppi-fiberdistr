import pandas as pd
import numpy as np
import gurobipy as gp
from gurobipy import GRB


rF2PYield2 = []

def inModel(model, constrName):
    if isinstance(constrName, (list, tuple, set, np.ndarray)):
        return [i for i in model.m.getConstrs() for x in constrName if x in i.getAttr(GRB.Attr.ConstrName)]
    else:
        return [i for i in model.m.getConstrs() if constrName in i.getAttr(GRB.Attr.ConstrName)]


def setfAvg(model, target='Containerboard', incr=5):
    """
    Increase recycled content for one or more product sectors by given percentage point

    :param model: Gurobi model to be modified
    :param target: product name(s) as str
    :param incr: percent of increase as float
    :return: new constraints to update model
    """
    targetProd = pd.DataFrame(pd.Series(False, index=model.fProd)).T
    fRecFix = dict.fromkeys(model.fProd, False)
    fTarget = dict.fromkeys(model.fProd, False)
    rAvg = {t: model.oldDemand[t].dot(model.rLevel[t]) / model.oldDemand[t].sum() for t in model.fProd}

    if not isinstance(target, (list, tuple, set, np.ndarray)): target = [target]
    if not isinstance(incr, (list, tuple, set, np.ndarray)): incr = [incr]

    incr = [i / 100 for i in incr]
    targetProd[target] = True
    fRecFix.update({i: True for i in model.fProd if 'P&W' in i or 'News' in i})
    fTarget.update({i: True for i in model.fProd if i in target})
    targInc = [sum(x) for x in zip(incr, [rAvg[i] for i in target])]
    rAvg.update({i: j for i, j in zip(target, targInc)})

    model.m.remove(inModel(model, ['maxDemand', 'recFix', 'recMax', 'prodTargetRec']))

    # additional constraints
    model.m.addConstrs((model.prodDemand[r, t] <= model.maxDemand[t].loc[r] for t in model.fProd
                        for r in model.rLevel[t] if not fRecFix[t]), "maxDemand")
    model.m.addConstrs(
        (gp.quicksum(r * model.prodDemand[r, t] for r in model.rLevel[t]) == rAvg[t] * model.oldDemand[t].sum() for t in
         model.fProd
         if fRecFix[t]), "recFix")
    model.m.addConstrs(
        (gp.quicksum(r * model.prodDemand[r, t] for r in model.rLevel[t]) <= rAvg[t] * model.oldDemand[t].sum() for t in
         model.fProd
         if not fRecFix[t] and not fTarget[t]), "recMax")
    model.m.addConstrs((gp.quicksum(r * model.prodDemand[r, t] for r in model.rLevel[t]) ==
                        rAvg[t] * model.productUse.loc[t, 'Domestic Consumption'] for t in model.fProd
                        if fTarget[t]), "prodTargetRec")
    model.m.update()

    return model


def setmpYield(model, perc=5):
    global rF2PYield2
    rF2PYield2 = model.rF2PYield.copy()

    mixedList = [key for key in model.rFiber if 'MIXED' in model.rCat[key]]
    rF2PYield2.loc[mixedList] = rF2PYield2.loc[mixedList] * (1 - perc / 100)

    model.m.remove(inModel(model, ['rRecipeMin', 'rRecipeMax', 'rRecipe', 'rPulpinFlow']))
    model.m.update()

    model.m.addConstrs(
        (rF2PYield2.loc[i, p] * model.rfiber2pulp[i, p] >= model.recipeMin.loc[i, p] * model.rpulpProd[p]
         for p in model.rPulp for i in model.rFiber), "rRecipeMin")
    model.m.addConstrs(
        (rF2PYield2.loc[i, p] * model.rfiber2pulp[i, p] <= model.recipeMax.loc[i, p] * model.rpulpProd[p]
         for p in model.rPulp for i in model.rFiber), "rRecipeMax")
    model.m.addConstrs(
        (rF2PYield2.loc[i, p] * model.rfiber2pulp[i, p] - model.rRecipeYield.loc[i, p] * model.rpulpProd[p] == 0
         for p in model.rPulp for i in model.rFiber if 'Deinked' in p or 'P&W' in p or 'News' in p), 'rRecipe')
    model.m.addConstrs(
        (rF2PYield2.loc[:, p].dot(model.rfiber2pulp.select('*', p)) == model.rpulpProd[p] for p in model.rPulp),
        "rPulpinFlow")

    model.m.update()
    return model


def setDemand(model, prod='Containerboard', perc=5):
    model.m.remove(inModel(model, ['non-constantDemand', 'constantDemand']));
    model.m.update()

    if not isinstance(prod, (list, tuple, set, np.ndarray)): prod = [prod]
    if not isinstance(perc, (list, tuple, set, np.ndarray)): perc = [perc]

    newD = {t: 0 for t in model.fProd}
    newD.update({t: x for t, x in zip(prod, perc)})

    model.m.addConstrs(
        (model.prodDemand.sum('*', t) == model.productUse.loc[t, 'Domestic Consumption'] * (1 + newD[t] / 100)
         for t in model.fProd), "non-constantDemand")

    model.m.update()
    return model


def setContam(model, recovery='mix', custom_cm=0):
    affChan = ['Residential', 'Retail']  # affected channels, most susceptible to changes in recycling
    cRate = [0.272, 0.17, 0.19, 0.24]  # contamination rates from Eureka Recycling, The Recycling Partnership, and WM
    rng = np.random.default_rng()
    x = rng.normal(np.mean(cRate), np.std(cRate), 1)

    wPaperYld2 = model.wPaperYld.copy()

    try:  # might be using this wrong
        if custom_cm > 0.001:
            for t in model.fProd:
                wPaperYld2.update(
                    {(i, t, s): wPaperYld2[i, t, s] * (1 - custom_cm) for s in model.wsource for i in model.rFiber if
                     s in affChan})
        else:
            if recovery.lower() == 'mix':
                pass  # do nothing, mix is default

            elif recovery.lower() == 'ssr':
                for t in model.fProd:
                    wPaperYld2.update(
                        {(i, t, s): wPaperYld2[i, t, s] * (1 - x) for s in model.wsource for i in model.rFiber if
                         s in affChan})

            elif recovery.lower() == 'dual':
                for t in model.fProd:
                    wPaperYld2.update(
                        {(i, t, s): wPaperYld2[i, t, s] * (1 + x) for s in model.wsource for i in model.rFiber if
                         s in affChan})
                wPaperYld2.update({(i, t, s): 1 for s in model.wsource for i in model.rFiber for t in model.fProd if
                                        wPaperYld2[i, t, s] > 1})

        for i in wPaperYld2.keys():
            if isinstance(wPaperYld2[i], (list, tuple, set, np.ndarray)):
                wPaperYld2[i] = wPaperYld2[i][0]

        model.m.remove(inModel(model, 'rFiberAvail'));
        model.m.update()

        model.m.addConstrs((model.rfiber2pulp.sum(i, '*') + model.rExpNew[i] + model.rResidue[i] -
                            gp.quicksum(gp.quicksum(
                                wPaperYld2[i, t, s] * model.sCollectNew[t, s] * model.prodConsumed[t, s] for
                                s in model.wsource)
                                        for t in model.fProd) == 0 for i in model.rFiber), "rFiberAvail")

        model.m.update()

    except ValueError:
        print("Not a valid scenario. Try 'Mix', 'SSR', or 'Dual'.")


def fiberDeg(model, custom_fb=0):
    rng = np.random.default_rng()
    switch = {
        'SUBS': rng.random() / 100 + 0.125,  # 12.5 - 13.5%, about 1 in 8
        'HIGH': rng.random() * 1.5 / 100 + 0.135,  # 13.5 - 15%
        'CORRUGATED': rng.random() * 2 / 100 + 0.15,  # 15 - 17%
        'MIXED': rng.random() * 3 / 100 + 0.17,  # 17 - 20%
        'NEWS': rng.random() * 5 / 100 + 0.2  # 20 - 25%, about 1 in 4
    }

    global rF2PYield2
    if not isinstance(rF2PYield2, list):
        pass  # means MP yield scenario has been set, so just use that version
    else:
        rF2PYield2 = model.rF2PYield.copy()

    if custom_fb > 0.001:
        for i in model.rFiber:  # decrease for each grade in dataframe
            rF2PYield2.loc[i] = rF2PYield2.loc[i] * (1 - custom_fb)
    else:
        for i in model.rFiber:  # decrease for each grade in dataframe
            rF2PYield2.loc[i] = rF2PYield2.loc[i] * (1 - switch.get(model.rCat[i], 'N/A'))

    # remove old constraints
    model.m.remove([inModel(model, ['rRecipeMin', 'rRecipeMax', 'rRecipe', 'rPulpinFlow'])])

    # add new ones
    model.m.addConstrs(
        (rF2PYield2[p].dot(model.rfiber2pulp.select('*', p)) == model.rpulpProd[p] for p in model.rPulp),
        "rPulpinFlow")
    model.m.addConstrs(
        (rF2PYield2.loc[i, p] * model.rfiber2pulp[i, p] >= model.recipeMin.loc[i, p] * model.rpulpProd[p] for p in
         model.rPulp for i in
         model.rFiber),
        "rRecipeMin")
    model.m.addConstrs(
        (rF2PYield2.loc[i, p] * model.rfiber2pulp[i, p] <= model.recipeMax.loc[i, p] * model.rpulpProd[p] for p in
         model.rPulp for i in
         model.rFiber),
        "rRecipeMax")
    model.m.addConstrs(
        (rF2PYield2.loc[i, p] * model.rfiber2pulp[i, p] - model.rRecipeYield.loc[i, p] * model.rpulpProd[p] == 0
         for p in model.rPulp for i in model.rFiber
         if 'Deinked' in p or 'P&W' in p or 'News' in p), 'rRecipe')

    model.m.update()


def resetScenarios(model):
    model.m.remove([  # TODO: re-add constraints from base case 791 -> 797?
        inModel(model,
                ['maxDemand', 'recFix', 'recMax', 'prodTargetRec', 'rRecipeMin', 'rRecipeMax', 'rRecipe', 'rPulpinFlow',
                 'rFiberAvail', 'constantDemand', 'non-constantDemand']),
    ])
    # re-add base constraints
    model.m.addConstrs(
        (model.rF2PYield[p].dot(model.rfiber2pulp.select('*', p)) == model.rpulpProd[p] for p in model.rPulp),
        "rPulpinFlow")
    model.m.addConstrs(
        (model.rF2PYield.loc[i, p] * model.rfiber2pulp[i, p] >= model.recipeMin.loc[i, p] * model.rpulpProd[p] for p in
         model.rPulp for i in
         model.rFiber),
        "rRecipeMin")
    model.m.addConstrs(
        (model.rF2PYield.loc[i, p] * model.rfiber2pulp[i, p] <= model.recipeMax.loc[i, p] * model.rpulpProd[p] for p in
         model.rPulp for i in
         model.rFiber),
        "rRecipeMax")
    model.m.addConstrs(
        (model.rF2PYield.loc[i, p] * model.rfiber2pulp[i, p] - model.rRecipeYield.loc[i, p] * model.rpulpProd[p] == 0
         for p in model.rPulp for i
         in model.rFiber
         if 'Deinked' in p or 'P&W' in p or 'News'), 'rRecipe')
    model.m.addConstrs((model.rfiber2pulp.sum(i, '*') + model.rExpNew[i] + model.rResidue[i] -
                        gp.quicksum(gp.quicksum(
                            model.wPaperYld[i, t, s] * model.sCollectNew[t, s] * model.prodConsumed[t, s] for s in
                            model.wsource)
                                    for t in model.fProd) == 0 for i in model.rFiber), "rFiberAvail")
    model.m.addConstrs(
        (model.prodDemand.sum('*', t) == model.productUse.loc[t, 'Domestic Consumption'] for t in model.fProd),
        "constantDemand")
    model.m.addConstrs(
        (model.prodDemand[r, t] <= model.maxDemand[t].loc[r] for t in model.fProd for r in model.rLevel[t]
         if t not in model.fProd[-2:]), "maxDemand")

    rAvg = {t: model.oldDemand[t].dot(model.rLevel[t]) / model.oldDemand[t].sum() for t in model.fProd}
    model.m.addConstrs(
        (gp.quicksum(r * model.prodDemand[r, t] for r in model.rLevel[t]) == rAvg[t] * model.oldDemand[t].sum()
         for t in model.fProd if t in model.fProd[-2:]), "recFix")
    model.m.addConstrs(
        (gp.quicksum(r * model.prodDemand[r, t] for r in model.rLevel[t]) <= rAvg[t] * model.oldDemand[t].sum()
         for t in model.fProd if t not in model.fProd[-2:]), "recMax")

    model.m.update()

    # reset global
    global rF2PYield2
    rF2PYield2 = model.rF2PYield.copy()

    return model
