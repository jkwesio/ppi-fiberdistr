# -*- coding: utf-8 -*-
"""
Created on Wed Aug 4 2021, last edited 27 Oct 2021

Fiber flow emissions calculations module - class version

Inputs:
    Excel file with old PPI market & emissions data ('FiberModelAll_Python_v3-yields.xlsx')
    
Outputs:
    Dict of keys 'old','new','forest','trade' with emissions calcs
    
(*testing inputs*
     x = 'FiberModelAll_Python_v2.xlsx'
     f2pVolOld = pd.read_excel(x, 'OldData', usecols="A:I", skiprows=1, nrows=21, index_col=0)
     pbpVolOld = pd.read_excel(x, 'OldData', usecols="K:R", skiprows=1, nrows=14, index_col=0)
     pbpVolOld.columns = [x[:-2] for x in pbpVolOld.columns]
     consCollOld = pd.read_excel(x, 'OldData', usecols="K:Q", skiprows=34, nrows=3, index_col=0)
     
     rLevel = pd.read_excel(x, 'Demand', usecols="F:K", skiprows=16, nrows=5)
     rLevel = {t: list(rLevel[t][np.isfinite(rLevel[t])].values) for t in fProd}
     
     fProd = [t for t in f2pVolOld.iloc[:,:6].columns]
     fProdM = [t for t in f2pVolOld.iloc[:,:7].columns]
     rFiber = f2pVolOld.index[:16]
     vFiber = f2pVolOld.index[16:]
     rPulp = [p for p in pbpVolOld.index if 'Rec' in p]
     vPulp = [q for q in pbpVolOld.index if 'Vir' in q]
     fPulp = [f for f in pbpVolOld.index]
     
     import numpy as np
     f2pYld = pd.read_excel(x, 'Fiber', usecols="I:O", skiprows=1, nrows=21)
     f2pYld.index = np.concatenate([rFiber.values, vFiber.values], axis=0)
     pulpYld = pd.read_excel(x, 'Pulp', usecols="D", skiprows=1, nrows=14)
     pulpYld.index = rPulp + vPulp
     
     transPct = pd.read_excel(x, 'EmTables', usecols="L:P", skiprows=32, nrows=11, index_col=0)
     transKM = pd.read_excel(x, 'EmTables', usecols="L:P", skiprows=46, nrows=11, index_col=0)
     transUMI = pd.read_excel(x, 'EmTables', usecols="L:P", skiprows=59, nrows=1, index_col=0)
     
     rsdlModes = pd.read_excel(x, 'EmTables', usecols="A:G", skiprows=32, nrows=6, index_col=0)
     rsdlbio = pd.read_excel(x, 'EmTables', usecols="A:H", skiprows=41, nrows=4, index_col=0)
     rsdlbio = rsdlbio.fillna(0)
     rsdlfos = pd.read_excel(x, 'EmTables', usecols="A:H", skiprows=48, nrows=4, index_col=0)
     rsdlfos = rsdlfos.fillna(0)
     
     woodint = pd.read_excel(x, 'EmTables', usecols="A:H", skiprows=58, nrows=1, index_col=0)
     wtotalGHGb0 = pd.read_excel(x, 'EmTables', usecols="A:K", skiprows=62, nrows=6, index_col=0)
     wtotalGHGb1 = pd.read_excel(x, 'EmTables', usecols="A:K", skiprows=71, nrows=6, index_col=0)
     wbioGHGb0 = pd.read_excel(x, 'EmTables', usecols="A:K", skiprows=80, nrows=6, index_col=0)
     wbioGHGb1 = pd.read_excel(x, 'EmTables', usecols="A:K", skiprows=89, nrows=6, index_col=0)
     wfosGHGb0 = pd.read_excel(x, 'EmTables', usecols="A:K", skiprows=98, nrows=6, index_col=0)
     wfosGHGb1 = pd.read_excel(x, 'EmTables', usecols="A:K", skiprows=107, nrows=6, index_col=0)
     
     exportOld = pd.read_excel(x, 'OldData', usecols="E:G", skiprows=31, nrows=16, index_col=0)
     exportOld.iloc[:,:-1] = exportOld.iloc[:,:-1]
     exportNew = exportOld.iloc[:,:-1] * 1.5
     exportNew.columns = ['exportNew']
     exportNew = exportNew.assign(TransCode=exportOld['TransCode'].values)
     fiberType = pd.read_excel(x, 'OldData', usecols="A:B", skiprows=31, nrows=20, index_col=0)
     chinaVals = pd.read_excel(x, 'EmTables', usecols="L:M", skiprows=66, nrows=3, index_col=0)
     chinaCons = pd.read_excel(x, 'EmTables', usecols="L:M", skiprows=72, nrows=6, index_col=0)
     fYield = pd.read_excel(x, 'EmTables', usecols="L:N", skiprows=81, nrows=5, index_col=0)
     
 )

@author: Jacqueline Baidoo
"""
import pandas as pd
import numpy as np

class en_emissions(): # energy & emissions
    
    def __init__(cls,xls,fProd,rLevel,f2pYld,pulpYld,f2pVolNew,pbpVolNew,consCollNew,exportNew,demandNew):
        # xls (str) - name of Excel spreadsheet to pull data from
        # fProd (list) - list of products in current scenario
        # rLevel (df) - recycled content level by product
        # f2pYld (df) - fiber to pulp yield by pulp product; indexed by fiber
        # pulpYld (df) - pulp to product yield; pulp as index
        # f2pVolNew (df) - fiber to pulp volume (in short tons); indexed by pulp name
        # pbpVolNew (df) - pulp by product volume; indexed by pulp name
        # consCollNew (df) - domestic consumption, collection, and recovery by product
        # demandNew (df) - new demand by product; indexed by rec level
        uC = 0.907185 # unit conversion of MM US ton to Mg/metric ton
        
        cls.fProd = fProd
        cls.fProdM = fProd + ['Market']
        cls.rLevel = rLevel
        cls.f2pYld = f2pYld
        cls.pulpYld = pulpYld
        cls.f2pVolNew = f2pVolNew * uC
        cls.pbpVolNew = pbpVolNew * uC
        cls.consCollNew = consCollNew * uC
        cls.exportNew = exportNew * uC
        cls.demandNew = {t: demandNew[t] * uC for t in demandNew.keys()}
        
        with pd.ExcelFile(xls) as x:
            # Old data   
            cls.f2pVolOld = pd.read_excel(x, 'OldData', usecols="A:I", skiprows=1, nrows=21, index_col=0)
            cls.f2pVolOld.iloc[:,:-1] = cls.f2pVolOld.iloc[:,:-1] * uC * 1000

            cls.f2pVolNew = cls.f2pVolNew.assign(TransCode=cls.f2pVolOld['TransCode'].values)
            
            cls.pbpVolOld = pd.read_excel(x, 'OldData', usecols="K:R", skiprows=1, nrows=14, index_col=0)
            cls.pbpVolOld.columns = [x[:-2] for x in cls.pbpVolOld.columns] # has .1 after column names for pandas duplicate
            cls.pbpVolOld.iloc[:,:-1] = cls.pbpVolOld.iloc[:,:-1] * uC * 1000

            cls.pbpVolNew = cls.pbpVolNew.assign(TransCode=cls.pbpVolOld['TransCode'].values)
            
            cls.prodLD = pd.read_excel(x, 'OldData', usecols="K:Q", skiprows=19, nrows=5, index_col=0) * uC * 1000
            cls.prodDemand = pd.read_excel(x, 'OldData', usecols="A:G", skiprows=26, nrows=1, index_col=0) * uC * 1000
            
            cls.consCollOld = pd.read_excel(x, 'OldData', usecols="K:Q", skiprows=29, nrows=3, index_col=0) * uC * 1000
            
            cls.exportOld = pd.read_excel(x, 'OldData', usecols="E:G", skiprows=31, nrows=16, index_col=0)
            cls.exportOld.iloc[:,:-1] = cls.exportOld.iloc[:,:-1] * uC * 1000
            cls.exportNew = cls.exportNew.assign(TransCode=cls.exportOld['TransCode'].values)
            cls.fiberType = pd.read_excel(x, 'OldData', usecols="A:B", skiprows=31, nrows=20, index_col=0)
            
            cls.rFiber = cls.f2pVolOld.index[:16]
            cls.vFiber = cls.f2pVolOld.index[16:]
            cls.rPulp = [p for p in cls.pbpVolOld.index if 'Rec' in p]
            cls.vPulp = [q for q in cls.pbpVolOld.index if 'Vir' in q]
            cls.fPulp = [f for f in cls.pbpVolOld.index]
            
            # Emissions Info
            cls.chemicals = pd.read_excel(x, 'nonFiber', usecols="A:B,E:L", skiprows=2, nrows=42, index_col=0)
            cls.eolEmissions = pd.read_excel(x, 'EmTables', usecols="A:G", skiprows=2, nrows=3, index_col=0)
            
            cls.bfEI = pd.read_excel(x, 'EmTables', usecols="J:P", skiprows=2, nrows=3, index_col=0)
            cls.bfEI.columns = [x[:-2] for x in cls.bfEI.columns] # has .1 after column names for some reason
            cls.bioPct = pd.read_excel(x, 'EmTables', usecols="J:P", skiprows=8, nrows=2, index_col=0)
            cls.pwpEI = pd.read_excel(x, 'EmTables', usecols="O:P", skiprows=14, nrows=5, index_col=0)
            cls.bfCO2 = pd.read_excel(x, 'EmTables', usecols="A:G", skiprows=9, nrows=2, index_col=0)
            
            cls.fuelTable = pd.read_excel(x, 'EmTables', usecols="A:M", skiprows=15, nrows=13, index_col=0)
            cls.fuelTable = cls.fuelTable.fillna(0)
            
            cls.rsdlModes = pd.read_excel(x, 'EmTables', usecols="A:G", skiprows=32, nrows=6, index_col=0)
            cls.rsdlbio = pd.read_excel(x, 'EmTables', usecols="A:H", skiprows=41, nrows=4, index_col=0)
            cls.rsdlbio = cls.rsdlbio.fillna(0)
            cls.rsdlfos = pd.read_excel(x, 'EmTables', usecols="A:H", skiprows=48, nrows=4, index_col=0)
            cls.rsdlfos = cls.rsdlfos.fillna(0)
            
            cls.transPct = pd.read_excel(x, 'EmTables', usecols="L:P", skiprows=32, nrows=11, index_col=0)
            cls.transKM = pd.read_excel(x, 'EmTables', usecols="L:P", skiprows=46, nrows=11, index_col=0)
            cls.transUMI = pd.read_excel(x, 'EmTables', usecols="L:P", skiprows=59, nrows=1, index_col=0)
            
            cls.woodint = pd.read_excel(x, 'EmTables', usecols="A:H", skiprows=58, nrows=1, index_col=0)
            cls.wtotalGHGb0 = pd.read_excel(x, 'EmTables', usecols="A:K", skiprows=62, nrows=6, index_col=0)
            cls.wtotalGHGb1 = pd.read_excel(x, 'EmTables', usecols="A:K", skiprows=71, nrows=6, index_col=0)
            cls.wbioGHGb0 = pd.read_excel(x, 'EmTables', usecols="A:K", skiprows=80, nrows=6, index_col=0)
            cls.wbioGHGb1 = pd.read_excel(x, 'EmTables', usecols="A:K", skiprows=89, nrows=6, index_col=0)
            cls.wfosGHGb0 = pd.read_excel(x, 'EmTables', usecols="A:K", skiprows=98, nrows=6, index_col=0)
            cls.wfosGHGb1 = pd.read_excel(x, 'EmTables', usecols="A:K", skiprows=107, nrows=6, index_col=0)
            
            cls.chinaVals = pd.read_excel(x, 'EmTables', usecols="L:M", skiprows=66, nrows=3, index_col=0)
            cls.chinaCons = pd.read_excel(x, 'EmTables', usecols="L:M", skiprows=72, nrows=6, index_col=0)
            cls.fYield = pd.read_excel(x, 'EmTables', usecols="L:N", skiprows=81, nrows=5, index_col=0)
    
    def calculateTrans(cls,transVol):
        # transVol [df] - item, volume (in Mg) by product, TransCode; indexed by fiberCode or other label
        # transPct [df] - % traversed for transMode by transCode; indexed by transCode
        # transKM [df] - distance traversed for transMode by transCode; indexed by transCode
        # transUMI [s] - unit impact by mode (truck, train, boat); indexed by "transUMI"
        transImpact = pd.Series(0, index = cls.fProd)
        
        tC = transVol['TransCode']
        tC = tC[(tC != 0) & (tC != 1)] # index non-zero/non-NaN elements only
        transVol = transVol.loc[tC.index]
        for t in cls.fProd:
            for m in cls.transUMI.columns:
                transImpact[t] += sum(transVol[t] * cls.transPct.loc[tC,m].values * cls.transKM.loc[tC,m].values * cls.transUMI[m].values * 1)
                
        return transImpact
    
    def calculateChem(cls,chemicals,prodDemand):
        # chemicals [df] - nonfiber name, % use by product, transCode, impact factor; indexed by number
        # prodDemand [df] - total demand; indexed by product
        chemImpact = pd.Series(0, index = cls.fProd, name = 'chemImp')
        chemVol = pd.DataFrame(0, index = chemicals.index, columns = cls.fProd)
        
        for t in cls.fProd:
            chemImpact[t] = sum(prodDemand[t].values * chemicals[t] * chemicals['Impact Factor'])
            chemVol[t] = chemicals[t] * prodDemand[t].values
        chemVol = chemVol.join(chemicals['TransCode'])
        chemTrans = pd.Series(cls.calculateTrans(chemVol), name = 'chemTrans')
        
        chemImpact = pd.DataFrame(chemImpact)
        return pd.concat([chemImpact, chemTrans], axis=1)
    
    def calculateEoL(cls,eolEmissions,consColl):
        # eolEmissions [df] - biogenic and fossil CO2 emission factors & transportation code by product; indexed by bio/fosCO2
        # consColl [df] - domestic consumption, collection, and recovery by product; indexed by name        
        prod2landfill = pd.Series(consColl.loc['Domestic Consumption'] - consColl.loc['Recovery Volume'],
                                  index = cls.fProd, name = 'prod2landfill')
        mrf2landfill = pd.Series(consColl.loc['Collection Volume'] - consColl.loc['Recovery Volume'],
                                 index = cls.fProd, name = 'mrf2landfill')
        
        bioEoL = pd.Series(prod2landfill * eolEmissions.loc['bioCO2'], index = cls.fProd, name = 'bioEoL')
        
        mrf2landfill = pd.DataFrame(mrf2landfill) # works b/c all prods have same TransCode
        transEoL = pd.Series(cls.calculateTrans(mrf2landfill.T.assign(TransCode=eolEmissions.loc['TransCode'].values[0])),
                             index = cls.fProd, name = 'eolTrans')
        
        fesTransEoL = pd.Series(prod2landfill * eolEmissions.loc['fossilCO2'] + transEoL, index = cls.fProd,
                                name = 'fesTransEoL')
        bftEoL = pd.Series(bioEoL + fesTransEoL, name = 'bftEoL')
        
        return pd.concat([bioEoL, fesTransEoL, bftEoL, transEoL], axis=1)
    
    def getEnergyYldCoeff(cls,f2pVol,pbpVol):
        # f2pVol [df] - recycled fiber to pulp (in Mg); indexed by fiber code
        # pbpVol [df] - pulp by product (in Mg); indexed by pulp name
        #
        # PYCoeff [s] - pulp yield coeffient; indexed by pulp
        f2pByPulp = pd.Series(0, index = pbpVol.index, name = 'fiber2pulp')
        
        for p in cls.rPulp:
            f2pByPulp[p] = sum([f2pVol.loc[cls.rFiber,t].sum() for t in cls.fProdM 
                               if cls.fProdM.index(t) == cls.rPulp.index(p)])
        for q in cls.vPulp:
            f2pByPulp[q] = sum([f2pVol.loc[cls.vFiber,t].sum() for t in cls.fProdM 
                               if cls.fProdM.index(t) == cls.vPulp.index(q)])
            
        pulpProd = pd.Series([pbpVol.loc[i].sum() for i in pbpVol.index], index = pbpVol.index, name = 'pulpProd')
        PYCoeff = (pd.Series(f2pByPulp / pulpProd, name = 'pulpYldCoeff'))
        PYCoeff.replace([np.inf, -np.inf], np.nan, inplace=True)
        PYCoeff = PYCoeff.fillna(0)
        
        return PYCoeff
    
    def getEnergyPulpPct(cls,pbpVol):
        # pbpVol [df] - pulp by product (in Mg); indexed by pulp name
        #
        # pulpPct [df] - % of rec/vir pulp used in product; indexed by pulp name
        pulpPct = pbpVol.copy().drop(['TransCode'], axis=1)

        for t in pulpPct.columns:
            rTotalPulp = pulpPct.loc[cls.rPulp,t].sum()
            vTotalPulp = pulpPct.loc[cls.vPulp,t].sum()

            pulpPct.loc[cls.rPulp,t] = pulpPct.loc[cls.rPulp,t] / rTotalPulp
            pulpPct.loc[cls.vPulp,t] = pulpPct.loc[cls.vPulp,t] / vTotalPulp
        
        return pulpPct.fillna(0)
    
    def getEnergyMultiProd(cls,PYMult,pulpPct):
        # PYMult [s] - pulp yield multiplier; indexed by pulp name
        # pulpPct [df] - % of rec/vir pulp used in product; indexed by pulp name
        #
        # (return) [df] -  rec/vir yield multiprod by product; index by r/vYldMultiProd
        rYldMultiProd = pd.Series([sum(pulpPct.loc[cls.rPulp,t] * PYMult[cls.rPulp]) for t in cls.fProd],
                                   index = cls.fProd, name = 'rYldMultiProd')
        vYldMultiProd = pd.Series([sum(pulpPct.loc[cls.vPulp,t] * PYMult[cls.vPulp]) for t in cls.fProd],
                                   index = cls.fProd, name = 'vYldMultiProd')
        
        rYldMultiProd.replace([np.inf, -np.inf], np.nan, inplace=True)
        vYldMultiProd.replace([np.inf, -np.inf], np.nan, inplace=True)
        
        return pd.concat([rYldMultiProd.fillna(0), vYldMultiProd.fillna(0)], axis=1)
    
    def calculateEnergy(cls,pbpVol,prodLD,multiProd,pwpEI,paperEI):
        # prodLD (df) - demand by product; indexed by % recycled content level
        # bfEI (df) - bio & fes energy intensity fitting parameters by product; indexed by name
        # bioPct (df) - bio fitting parameter for PWP; indexed by name
        # pwpEI (df) - energy intensity of PWP pulp; indexed by pulp name
        # paperEI (df) - paper production energy intensity; indexed by 'PPE'
        # pbpVol (df) - pulp by product (in Mg); indexed by pulp name
        # multiProd (df) - rec/vir yield multiprod by product; indexed by product
        bioEnergy = pd.Series(0, index = cls.fProd, name = "bioEnergy")
        fesEnergy = pd.Series(0, index = cls.fProd, name = 'fesEnergy')
        totalEnergy = pd.Series(0, index = cls.fProd, name = 'totalEnergy')
        
        for t in cls.fProd:
            bioEnergy[t] = sum(prodLD[t].values[:len(cls.rLevel[t])] *
                               sum([r * cls.bfEI.loc['bioEI b1',t] + cls.bfEI.loc['bioEI b0',t] for r in cls.rLevel[t]]))
            fesEnergy[t] = sum(prodLD[t].values[:len(cls.rLevel[t])] *
                               cls.bfEI.loc['fesEI',t] * multiProd.loc[t,'rYldMultiProd'])
            
            if 'P&W' or 'News' in t:
                avgrecPct = sum(prodLD[t].values[:len(cls.rLevel[t])] * cls.rLevel[t]) / prodLD[t].sum()
                bioPctPW = avgrecPct * cls.bioPct.loc['bioPct b1',t] + cls.bioPct.loc['bioPct b0',t]
                
                pulpProdEnergy = sum([pbpVol.loc[p,t] * pwpEI.loc[p].values[0] for p in pwpEI.index])
                ppEnergy = pulpProdEnergy + prodLD[t].sum() * paperEI.values[0]
                
                bioEnergy[t] = bioPctPW * ppEnergy
                fesEnergy[t] = (1 - bioPctPW) * ppEnergy * multiProd.loc[t,'rYldMultiProd']
                
            totalEnergy[t] = bioEnergy[t] + fesEnergy[t]
        
        return pd.concat([bioEnergy, fesEnergy, totalEnergy], axis=1)
    
    def calculateProduction(cls,calcEnergy):
        # calcEnergy (df) - bio, fes, and total energy from calculateEnergy; indexed by product
        # bfCO2 (df) - bio & fes CO2 fitting parameters; indexed by product
        bioCO2 = pd.Series(0, index = cls.fProd, name = 'bioCO2')
        fesCO2 = pd.Series(0, index = cls.fProd, name = 'fesCO2')
        totalCO2 = pd.Series(0, index = cls.fProd, name = 'totalCO2')
        for t in cls.fProd:
            bioCO2[t] = calcEnergy.loc[t,'bioEnergy'] * cls.bfCO2.loc['bioCO2 b1',t]
            fesCO2[t] = calcEnergy.loc[t,'fesEnergy'] * cls.bfCO2.loc['fesCO2 b1',t]
            totalCO2[t] = bioCO2[t] + fesCO2[t]
            
        return pd.concat([bioCO2, fesCO2, totalCO2], axis=1)
    
    def calculateFuel(cls,calcEnergy):
        # calcEnergy (df) - bio, fes, and total energy from calculateEnergy; indexed by product
        # fuelTable (df) - fuel impact by product; indexed by fuel type
        fuels = cls.fuelTable.index
        
        bioFI = pd.Series(0, index = cls.fProd, name = 'bioFuelImp')
        fesFI = pd.Series(0, index = cls.fProd, name = 'fesFuelImp')
        fuelImp = pd.Series(0, index = cls.fProd, name = 'fuelImp')
        
        for t in cls.fProd:
            bioFI[t] = calcEnergy.loc[t,'bioEnergy'] * sum([cls.fuelTable.loc[f,t] * cls.fuelTable.loc[f,'Upstream Impact Factor'] 
                                                             for f in fuels if cls.fuelTable.loc[f,'Fuel Type'] == 1])
            fesFI[t] = calcEnergy.loc[t,'fesEnergy'] * sum([cls.fuelTable.loc[f,t] * cls.fuelTable.loc[f,'Upstream Impact Factor'] 
                                                             for f in fuels if cls.fuelTable.loc[f,'Fuel Type'] == 2])
            fuelImp[t] = bioFI[t] + fesFI[t]
        
        fuelTransVol = cls.fuelTable.copy()
        fuel1 = [f for f in fuels if cls.fuelTable.loc[f,'Fuel Type'] == 1]
        fuel2 = [f for f in fuels if cls.fuelTable.loc[f,'Fuel Type'] == 2]
        for t in cls.fProd:
            fuelTransVol.loc[fuel1,t] = [calcEnergy.loc[t,'bioEnergy'] * cls.fuelTable.loc[f,t] * cls.fuelTable.loc[f,'FU/GJ'] 
                               for f in fuel1]
            fuelTransVol.loc[fuel2,t] = [calcEnergy.loc[t,'fesEnergy'] * cls.fuelTable.loc[f,t] * cls.fuelTable.loc[f,'FU/GJ'] 
                               for f in fuel2]
        
        fuelTrans = pd.Series(cls.calculateTrans(fuelTransVol), name = 'fuelTrans')
        
        return pd.concat([bioFI, fesFI, fuelImp, fuelTrans], axis=1)
    
    def calculateResidual(cls,pbpVol,f2pVol):
        # pbpVol [df] - pulp by product (in Mg); indexed by pulp name
        # f2pVol [df] - recycled fiber to pulp (in Mg); indexed by fiber code
        # f2pYld [df] - fiber to pulp yield by pulp product; indexed by fiber
        # pulpYld [df] - pulp to product yield; indexed by pulp
        # rsdlModes [df] - residual treatments modes; indexed by residual type
        # rsdlbio [df] - transport and biogenic emissions factors; indexed by residual treatment mode
        # rsdlfos [df] - transport and fossil emissions factors; indexed by residual treatment mode
        pulpProd = pd.Series(0, index = cls.rPulp + cls.vPulp, name = 'pulpProduced')
        fiberRes = pd.Series(0, index = cls.rPulp + cls.vPulp, name = 'fiberResidue')
        
        for p in cls.rPulp: # order of fPulp must match order of r/vPulp
            pulpProd[p] = sum([(f2pVol.loc[cls.rFiber,t].mul(cls.f2pYld.loc[cls.rFiber,t])).sum() for t in cls.fProdM 
                    if cls.fProdM.index(t) == cls.rPulp.index(p)])
            fiberRes[p] = sum([(f2pVol.loc[cls.rFiber,t].mul(1 - cls.f2pYld.loc[cls.rFiber,t])).sum() for t in cls.fProdM 
                    if cls.fProdM.index(t) == cls.rPulp.index(p)])
        for q in cls.vPulp:
            pulpProd[q] = sum([(f2pVol.loc[cls.vFiber,t].mul(cls.f2pYld.loc[cls.vFiber,t])).sum() for t in cls.fProdM 
                    if cls.fProdM.index(t) == cls.vPulp.index(q)])
            fiberRes[q] = sum([(f2pVol.loc[cls.vFiber,t].mul(1 - cls.f2pYld.loc[cls.vFiber,t])).sum() for t in cls.fProdM 
                    if cls.fProdM.index(t) == cls.vPulp.index(q)])
        
        pulpUP = pbpVol.iloc[:,:-1].div(pulpProd, axis=0).fillna(0) # pulpUsePct
        
        rFiberRsd = pd.Series((pulpUP.loc[cls.rPulp].mul(fiberRes[cls.rPulp], axis=0)).sum(), index = cls.fProd, name = 'rFiberRsd')
        rPulpRsd = pd.Series((pulpUP.loc[cls.rPulp].mul(1 - cls.pulpYld.iloc[:,0].loc[cls.rPulp], axis=0)).sum(), index = cls.fProd, name = 'rPulpRsd')
        rTotalRsd = rFiberRsd + rPulpRsd
        
        vFiberRsd = pd.Series((pulpUP.loc[cls.vPulp].mul(fiberRes[cls.vPulp], axis=0)).sum(), index = cls.fProd, name = 'vFiberRsd')
        vPulpRsd = pd.Series((pulpUP.loc[cls.vPulp].mul(1 - cls.pulpYld.iloc[:,0].loc[cls.vPulp], axis=0)).sum(), index = cls.fProd, name = 'vPulpRsd')
        vTotalRsd = vFiberRsd + vPulpRsd
        
        rsdlType = cls.rsdlModes.index
        rsdlQuantity = pd.DataFrame(0, index = rsdlType, columns = cls.fProd)
        for rt in rsdlType:
            if cls.rsdlModes.loc[rt,'Input Base'] ==  1:
                rsdlQuantity.loc[rt,:] = rTotalRsd * cls.rsdlModes.loc[rt,'Intensity']
            if cls.rsdlModes.loc[rt,'Input Base'] == 2:
                rsdlQuantity.loc[rt,:] = vTotalRsd * cls.rsdlModes.loc[rt,'Intensity'] 
        
        rsdlMode = cls.rsdlModes.columns[:-2]
        rsdlModeVol = {rM: pd.DataFrame(0, index = rsdlType, columns = cls.fProd)
                          for rM in rsdlMode}
        for rM in rsdlMode:
            rsdlModeVol[rM] = rsdlQuantity.mul(cls.rsdlModes[rM], axis=0)
            rsdlModeVol[rM] = rsdlModeVol[rM].assign(TransCode=cls.rsdlbio.loc[rM,'TransCode'] * np.ones(len(rsdlType)))

            rsdlModeVol[rM].replace([np.inf, -np.inf], np.nan, inplace=True) # TODO: what happens to make this inf?
            rsdlModeVol[rM].fillna(0)

        bioImp = pd.Series(0, index = cls.fProd, name = 'bioImp')
        fosImp = pd.Series(0, index = cls.fProd, name = 'fossilImp')
        
        for t in cls.fProd:
            bioImp[t] = sum([rsdlModeVol[rM][t].sum() * cls.rsdlbio.loc[rM,t] for rM in rsdlMode])
            fosImp[t] = sum([rsdlModeVol[rM][t].sum() * cls.rsdlfos.loc[rM,t] for rM in rsdlMode])
        
        biofosImp = pd.Series(bioImp + fosImp, name = 'bio+fos')
        
        rsdlTrans = pd.Series(0, index = cls.fProd, name = 'rsdlTrans')
        for rM in rsdlMode:
            rsdlTrans += cls.calculateTrans(rsdlModeVol[rM])
        
        return pd.concat([bioImp, fosImp, biofosImp, rsdlTrans], axis=1)
    
    def getExportTrans(cls,transVol):
        transImpact = pd.Series(0, index = transVol.columns[:-1])
        
        tC = transVol['TransCode']
        tC = tC[(tC != 0) & (tC != 1)] # index non-zero/non-NaN elements only
        transVol = transVol.loc[tC.index]
        for n in transVol.columns[:-1]:
            for m in cls.transUMI.columns:
                transImpact[n] += sum(transVol[n] * cls.transPct.loc[tC,m].values * cls.transKM.loc[tC,m].values * cls.transUMI[m].values)
        
        return transImpact.values
    
    def calculateExport(cls,exportOld,exportNew):
        # exportOld [df] old export from US; indexed by rec fiber
        # exportNew [df] new export from US; indexed by rec fiber
        impChange = pd.Series(0, index = cls.fYield.index, name = 'impChangeByGroup')
        sumChange = pd.Series(0, index = cls.fYield.index, name = 'sumNetChange')
        for r in impChange.index:
            typeMask = cls.fiberType[cls.fiberType['fiberType'] == r].index
            # impChange[r] = (exportOld.loc[typeMask, 'exportOld'] - exportNew.loc[typeMask, 'exportNew']).sum()
            impChange[r] = (exportNew.loc[typeMask, 'exportNew'] - exportOld.loc[typeMask, 'exportOld']).sum()
            sumChange[r] = impChange[r] * (1 - cls.fYield.loc[r,'US'] / cls.fYield.loc[r,'China'])
            
        beta = sumChange.sum() / (cls.chinaCons.loc['totalVir'].values + cls.chinaCons.loc['domesticRec'].values +
                            cls.chinaCons.loc['importRec-US'].values + cls.chinaCons.loc['importRec-nonUS'].values)
        
        # chinaTrans = cls.getExportTrans(exportOld) - cls.getExportTrans(exportNew)
        chinaTrans = cls.getExportTrans(exportNew) - cls.getExportTrans(exportOld)
        
        return cls.chinaVals.loc['Production'] * cls.chinaVals.loc['Energy Intensity'] * cls.chinaVals.loc['Emission Factor'] * beta + chinaTrans
    
    def getForestVirginGHG(cls,virCons,woodint,slope,intercept):
        # virCons [df] change in virgin consumption; products as columns
        # woodint [df] intervals of virgin wood consumption
        # slope [s] b1 value for GHG emissions
        # intercept[s] b0 value for GHG emissions
        for n in range(1,len(woodint.columns)):
            if (woodint[n].values <= virCons) & (virCons < woodint[n+1].values):
                return virCons * slope[n] + intercept[n]
        return 0 # catch values outside of interval
            
    def calculateForest(cls,virCons,forYear):
        # virCons [float] change in virgin consumption, sum of all products
        # forYear [int] forest year length for cumulative emissions calcs; 10-90 by ten        
        deltaTotalGHG = pd.Series(cls.getForestVirginGHG(virCons / 1e6, cls.woodint, cls.wtotalGHGb1[forYear], cls.wtotalGHGb0[forYear]),
                                  name = 'totalGHG') * 1e6
        deltabioGHG = pd.Series(cls.getForestVirginGHG(virCons / 1e6, cls.woodint, cls.wbioGHGb1[forYear], cls.wbioGHGb0[forYear]),
                                  name = 'bioGHG') * 1e6
        deltafosGHG = pd.Series(cls.getForestVirginGHG(virCons / 1e6, cls.woodint, cls.wfosGHGb1[forYear], cls.wfosGHGb0[forYear]),
                                  name = 'fosGHG') * 1e6
        
        return pd.concat([deltaTotalGHG, deltabioGHG, deltafosGHG], axis=1)
    
    def calculateEmissions(cls):
        # xls [df] - name of Excel spreadsheet to pull data from
        # fProd [df] - list of products in current scenario
        # rL [dict] - recycled content level by product
        # f2pYld [df] - fiber to pulp yield by pulp product; indexed by fiber
        # pulpYld [df] - pulp to product yield; indexed by pulp
        # f2pVolNew [df] - fiber to pulp volume (in Mg); indexed by fiber code
        # pbpVolNew [df] - pulp by product volume; indexed by pulp name
        # consCollNew [df] - domestic consumption, collection, and recovery by product                
        pulpNames = cls.rPulp + cls.vPulp
        mvO = [cls.pbpVolOld.loc[p] for p in pulpNames if 'Deinked' in p or 'Market' in p]
        marketVolOld = pd.concat([mvO[0],mvO[1]], axis=1).T
        mvN = [cls.pbpVolNew.loc[p] for p in pulpNames if 'Deinked' in p or 'Market' in p]
        marketVolNew = pd.concat([mvN[0],mvN[1]], axis=1).T
        
        # Chemical
        chemImp = cls.calculateChem(cls.chemicals, cls.prodDemand)
        
        # EoL
        oldEoL = cls.calculateEoL(cls.eolEmissions, cls.consCollOld)
        newEoL = cls.calculateEoL(cls.eolEmissions, cls.consCollNew)
        
        # Energy
        oldPulpPct = cls.getEnergyPulpPct(cls.pbpVolOld)
        newPulpPct = cls.getEnergyPulpPct(cls.pbpVolNew)
        
        oldPYCoeff = cls.getEnergyYldCoeff(cls.f2pVolOld, cls.pbpVolOld)
        newPYCoeff = cls.getEnergyYldCoeff(cls.f2pVolNew, cls.pbpVolNew)
        
        oldYldMultiplier = (oldPYCoeff / oldPYCoeff).fillna(0)
        newYldMultiplier = (newPYCoeff / oldPYCoeff).fillna(0)
        
        oldMP = cls.getEnergyMultiProd(oldYldMultiplier, oldPulpPct)
        newMP = cls.getEnergyMultiProd(newYldMultiplier, newPulpPct)
        
        oldEnergy = cls.calculateEnergy(cls.pbpVolOld, cls.prodLD, oldMP, cls.pwpEI.iloc[:-1], cls.pwpEI.iloc[-1])
        newEnergy = cls.calculateEnergy(cls.pbpVolNew, cls.demandNew, newMP, cls.pwpEI.iloc[:-1], cls.pwpEI.iloc[-1])
        
        # Production
        oldProd = cls.calculateProduction(oldEnergy)
        newProd = cls.calculateProduction(newEnergy)
        
        # Fuel
        oldFuel = cls.calculateFuel(oldEnergy)
        newFuel = cls.calculateFuel(newEnergy)
        
        # Residual
        oldRsdl = cls.calculateResidual(cls.pbpVolOld, cls.f2pVolOld)
        newRsdl = cls.calculateResidual(cls.pbpVolNew, cls.f2pVolNew)
        
        # Transportation
        oldFiberTrans = pd.Series(cls.calculateTrans(cls.f2pVolOld), name = 'fiberTrans')
        oldMarketTrans = pd.Series(cls.calculateTrans(marketVolOld), name = 'marketTrans')
        
        oldTrans = pd.concat([oldFiberTrans, oldMarketTrans, chemImp['chemTrans'], oldFuel['fuelTrans'],
                            oldRsdl['rsdlTrans'], oldEoL['eolTrans']], axis=1)
        
        newFiberTrans = pd.Series(cls.calculateTrans(cls.f2pVolNew), name = 'fiberTrans')
        newMarketTrans = pd.Series(cls.calculateTrans(marketVolNew), name = 'marketTrans')
        
        newTrans = pd.concat([newFiberTrans, newMarketTrans, chemImp['chemTrans'], newFuel['fuelTrans'],
                            newRsdl['rsdlTrans'], newEoL['eolTrans']], axis=1)
        
        # Export
        exportImp = cls.calculateExport(cls.exportOld,cls.exportNew)
        
        # FASOM/LURA
        forestGHG = cls.calculateForest(cls.f2pVolNew.iloc[:,:-1].loc[cls.vFiber].sum().sum() - 
                                     cls.f2pVolOld.iloc[:,:-1].loc[cls.vFiber].sum().sum(), 90)
        
        # Summary calcs for plotting
        oldSums = pd.concat([pd.Series(chemImp['chemImp'], name='chemImp'),
                             pd.Series(oldFuel['bioFuelImp'], name='fuelbio'),
                             pd.Series(oldFuel['fesFuelImp'], name='fuelfos'),
                             pd.Series(oldProd['totalCO2'], name='prodImp'),
                             pd.Series(oldProd['bioCO2'], name='prodbio'),
                             pd.Series(oldProd['fesCO2'], name='prodfos'),
                             pd.Series(oldEnergy['totalEnergy'], name='energy'),
                             pd.Series(oldEnergy['bioEnergy'], name='energybio'),
                             pd.Series(oldEnergy['fesEnergy'], name='energyfos'),
                             pd.Series(oldRsdl['bio+fos'], name='residImp'),
                             pd.Series(oldRsdl['bioImp'], name='residbio'),
                             pd.Series(oldRsdl['fossilImp'], name='residfos'),
                             pd.Series(oldEoL['bftEoL'], name='eolImp'),
                             pd.Series(oldEoL['bioEoL'], name='eolbio'),
                             pd.Series(oldEoL['fesTransEoL'], name='eolfos'),
                             pd.Series(oldProd['bioCO2'] + oldRsdl['bioImp'] + oldEoL['bioEoL'], name='bioCO2'),
                             pd.Series(oldTrans.sum(axis=1) + chemImp['chemImp'] + oldFuel['fuelImp'] + 
                                       oldProd['fesCO2'] + oldRsdl['fossilImp'] + oldEoL['fesTransEoL'], name='fossilCO2'),
                             pd.Series(oldProd['bioCO2'] + oldRsdl['bioImp'], name='g2gbio'),
                             pd.Series(oldProd['fesCO2'] + oldRsdl['fossilImp'] + oldTrans.sum(axis=1), name='g2gfos')], axis=1)
        oldSums = pd.concat([oldSums, pd.Series(oldSums['bioCO2'] + oldSums['fossilCO2'], name='totalImp')], axis=1)
        oldSums = pd.concat([oldSums, pd.Series(oldSums['totalImp'] / cls.prodLD.sum(), name='unitImp')], axis=1, sort=True)
        
        newSums = pd.concat([pd.Series(chemImp['chemImp'], name='chemImp'),
                             pd.Series(newFuel['bioFuelImp'], name='fuelbio'),
                             pd.Series(newFuel['fesFuelImp'], name='fuelfos'),
                             pd.Series(newProd['totalCO2'], name='prodImp'),
                             pd.Series(newProd['bioCO2'], name='prodbio'),
                             pd.Series(newProd['fesCO2'], name='prodfos'),
                             pd.Series(newEnergy['totalEnergy'], name='energy'),
                             pd.Series(newEnergy['bioEnergy'], name='energybio'),
                             pd.Series(newEnergy['fesEnergy'], name='energyfos'),
                             pd.Series(newRsdl['bio+fos'], name='residImp'),
                             pd.Series(newRsdl['bioImp'], name='residbio'),
                             pd.Series(newRsdl['fossilImp'], name='residfos'),
                             pd.Series(newEoL['bftEoL'], name='eolImp'),
                             pd.Series(newEoL['bioEoL'], name='eolbio'),
                             pd.Series(newEoL['fesTransEoL'], name='eolfos'),
                             pd.Series(newProd['bioCO2'] + newRsdl['bioImp'] + newEoL['bioEoL'], name='bioCO2'),
                             pd.Series(newTrans.sum(axis=1) + chemImp['chemImp'] + newFuel['fuelImp'] + 
                                       newProd['fesCO2'] + newRsdl['fossilImp'] + newEoL['fesTransEoL'], name='fossilCO2'),
                             pd.Series(newProd['bioCO2'] + newRsdl['bioImp'], name='g2gbio'),
                             pd.Series(newProd['fesCO2'] + newRsdl['fossilImp'] + newTrans.sum(axis=1), name='g2gfos')],axis=1)
        newSums = pd.concat([newSums, pd.Series(newSums['bioCO2'] + newSums['fossilCO2'], name='totalImp')], axis=1)
        newSums = pd.concat([newSums, pd.Series(newSums['totalImp'] / cls.prodLD.sum(), name='unitImp')], axis=1, sort=True)
        
        return {k: v for k,v in zip(['old','new','forest','trade','oldenergy','newenergy'],
                                    [oldSums,newSums,forestGHG,exportImp,oldEnergy,newEnergy])}