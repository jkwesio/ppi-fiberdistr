import matplotlib.pyplot as plt
import pandas as pd

def fUseChange(model, ax):
    nonPWP = [t for t in model.f2pVolNew.columns if 'P&W' not in t and 'Market' not in t and 'News' not in t]
    colors = ['#c55911']

    recData = model.f2pVolNew.loc[model.rFiber][nonPWP].sum().sum() - model.f2pVolOld.loc[model.rFiber][nonPWP].sum().sum()
    virData = model.f2pVolNew.loc[model.vFiber][nonPWP].sum().sum() - model.f2pVolOld.loc[model.vFiber][nonPWP].sum().sum()

    ax.bar(['Recovered Fiber', 'Virgin Pulpwood'], [recData / 1e6 * 0.907185, virData / 1e6 * 0.907185], color=colors)
    ax.set_ylabel('Total Fiber Use Change from Baseline (million Mg)')
    ax.set_title('Total Fiber Use Changes')

def fUseByProd(model, ax):
    labels = ['Recovered Fiber', 'Virgin Pulpwood']
    colors = ['#5b9bd5', '#ed7d31', '#a5a5a5', '#ffc000', '#002eff', '#ce57ff']

    recData = [model.f2pVolNew.loc[model.rFiber].sum()[t] - model.f2pVolOld.loc[model.rFiber].sum()[t] for t in model.fProd]
    virData = [model.f2pVolNew.loc[model.vFiber].sum()[t] - model.f2pVolOld.loc[model.vFiber].sum()[t] for t in model.fProd]

    # df = pd.DataFrame({t: [r, v] for t, r, v in zip(model.fProd[:-2], recData[:-2], virData[:-2])}, index=labels) / 1e6 * 0.907185
    # df.plot(ax=ax, kind='bar', title='Fiber Use by Product', rot=0, color=colors[:-2])
    df = pd.DataFrame({t: [r, v] for t, r, v in zip(model.fProd, recData, virData)}, index=labels) / 1e6 * 0.907185
    df.plot(ax=ax, kind='bar', title='Fiber Use by Product', rot=0, color=colors)

    ax.set_ylabel('Fiber Use Change from Baseline (million Mg)')

def fSources(model, ax):
    # fig, ax = plt.subplots()

    colors = ['#ffc000']

    addlData = sum([model.addlRec[i] for i in model.rFiber])
    expData = model.exportOld.sum().sum() - model.exportNew.sum().sum()

    # ax.bar(['From Additional Collection','From Export Market'], [addlData, expData], color='gold')
    ax.bar(['From Additional Collection', 'From Export Market'], [addlData / 1e6 * 0.907185, expData / 1e6 * 0.907185], color=colors)
    ax.set_ylabel('Recovered Fiber (million Mg)')
    ax.set_title('Sources of Recovered Fiber')

def mxpShift(model, ax):
    colors = ['#5b9bd5', '#ed7d31', '#a5a5a5', '#ffc000', '#002eff', '#ce57ff']

    mixedList = [key for key in model.rFiber if 'MIXED' in model.rCat[key]]
    mxpData = [model.f2pVolNew.loc[mixedList, t].sum() - model.f2pVolOld.loc[mixedList, t].sum() for t in model.fProd]

    df = pd.DataFrame({t: r for t, r in zip(model.fProd, mxpData)}, index=['MIXED Paper Shift']) / 1e6 * 0.907185
    df.plot(ax=ax, kind='bar', title='MIXED Paper Shift from Baseline', rot=0, color=colors)
    ax.set_ylabel('Fiber Use Change from Baseline (million Mg)')

def fShiftByCat(model, ax):
    labels = []
    for key in model.rCat.keys():
        if model.rCat[key] not in labels:
            labels.append(model.rCat[key])

    mixedData = {}
    newsData = {}
    corrData = {}
    subsData = {}
    highData = {}
    for i in model.rFiber:
        if model.rCat[i] in labels[0]:
            mixedData.update({i: [model.f2pVolNew.loc[i, t] - model.f2pVolOld.loc[i, t] for t in model.fProd]})
        elif model.rCat[i] in labels[1]:
            newsData.update({i: [model.f2pVolNew.loc[i, t] - model.f2pVolOld.loc[i, t] for t in model.fProd]})
        elif model.rCat[i] in labels[2]:
            corrData.update({i: [model.f2pVolNew.loc[i, t] - model.f2pVolOld.loc[i, t] for t in model.fProd]})
        elif model.rCat[i] in labels[3]:
            subsData.update({i: [model.f2pVolNew.loc[i, t] - model.f2pVolOld.loc[i, t] for t in model.fProd]})
        elif model.rCat[i] in labels[4]:
            highData.update({i: [model.f2pVolNew.loc[i, t] - model.f2pVolOld.loc[i, t] for t in model.fProd]})
        else:
            print(f'Fiber grade {i} was missed in calculating shift from baseline.')

    colors = ['#5b9bd5', '#ed7d31', '#a5a5a5', '#ffc000', '#002eff', '#ce57ff']

    df1 = pd.DataFrame.from_dict(mixedData, orient='index', columns=model.fProd) / 1e6 * 0.907185
    df1.plot(ax=ax[0], kind='bar', title='MIXED Grade Recovered Paper Shift', rot=0, color=colors)
    ax[0].set_ylabel('Fiber Use Change from Baseline (million Mg)')

    df2 = pd.DataFrame.from_dict(newsData, orient='index', columns=model.fProd) / 1e6 * 0.907185
    df2.plot(ax=ax[1], kind='bar', title='NEWS Grade Recovered Paper Shift', rot=0, color=colors)
    ax[1].set_ylabel('Fiber Use Change from Baseline (million Mg)')

    df3 = pd.DataFrame.from_dict(corrData, orient='index', columns=model.fProd) / 1e6 * 0.907185
    df3.plot(ax=ax[2], kind='bar', title='CORRUGATED Grade Recovered Paper Shift', rot=0, color=colors)
    ax[2].set_ylabel('Fiber Use Change from Baseline (million Mg)')

    df4 = pd.DataFrame.from_dict(subsData, orient='index', columns=model.fProd) / 1e6 * 0.907185
    df4.plot(ax=ax[3], kind='bar', title='PULP SUBS Grade Recovered Paper Shift', rot=0, color=colors)
    ax[3].set_ylabel('Fiber Use Change from Baseline (million Mg)')

    df5 = pd.DataFrame.from_dict(highData, orient='index', columns=model.fProd) / 1e6 * 0.907185
    df5.plot(ax=ax[4], kind='bar', title='HIGH Grade Recovered Paper Shift', rot=0, color=colors)
    ax[4].set_ylabel('Fiber Use Change from Baseline (million Mg)')


def dirEnergyCons(emissions, ax):
    # fig, ax = plt.subplots()
    colors = ['#44546a']

    bioData = emissions['new']['energybio'].sum() - emissions['old']['energybio'].sum()
    fosData = emissions['new']['energyfos'].sum() - emissions['old']['energyfos'].sum()

    # ax.bar(['Biomass Energy','Fossil Energy'], [bioData, fosData], color='slategray')
    ax.bar(['Biomass Energy', 'Fossil Energy'], [bioData, fosData], color=colors)
    ax.set_ylabel('Direct Energy Consumption Change from Baseline (GJ)')
    ax.set_title('Direct Energy Consumption')

def ghg_notForEoL(emissions, ax):
    # in order of upstream - G2G - trade
    bioData = [emissions['new']['fuelbio'].sum() - emissions['old']['fuelbio'].sum(),
               emissions['new']['g2gbio'].sum() - emissions['old']['g2gbio'].sum(), 0]
    fosData = [emissions['new']['fuelfos'].sum() - emissions['old']['fuelfos'].sum(),
               emissions['new']['g2gfos'].sum() - emissions['old']['g2gfos'].sum(), emissions['trade'].sum()]

    # stacked bar
    # fig, ax = plt.subplots()
    labels = ['Upstream Except Forest', 'Gate-to-Gate', 'International Trade']
    colors = ['#375da1', '#4472c4', '#a7b5db']

    i = 0
    for x in zip(bioData, fosData):
        ax.bar(['Biogenic GHG', 'Fossil GHG'], x, color=colors[i])
        i += 1

    ax.legend(labels)
    ax.set_ylabel('GHG Emission Change from Baseline (Mg CO2 eq)')
    ax.set_title('GHG Emissions Except Forest & EoL')

def ghg_wForEoL(emissions, ax):
    # in order of forest - eol
    bioData = [emissions['forest']['bioGHG'].sum(), emissions['new']['eolbio'].sum() - emissions['old']['eolbio'].sum()]
    fosData = [emissions['forest']['fosGHG'].sum(), emissions['new']['eolfos'].sum() - emissions['old']['eolfos'].sum()]

    # bars from pandas
    labels = ['Biogenic GHG', 'Fossil GHG']
    colors = ['#70ad47', '#ffc000']

    df = pd.DataFrame({k: [l, m] for k, l, m in zip(['Forest', 'EoL'], bioData, fosData)}, index=labels)
    # ax = df.plot(kind='bar', title='GHG Emissions Change in Forest & Agriculture & EoL', rot=0, color=colors)
    df.plot(ax=ax, kind='bar', title='GHG Emissions Change in Forest & Agriculture & EoL', rot=0, color=colors)
    ax.set_ylabel('Average annual emission over 90 years (Mg CO2 eq/year)')

def emBreakdown(emissions, ax):
    # in order of forest - upstream - G2G - eol - trade
    bioData = [emissions['forest']['bioGHG'].sum(),
               emissions['new']['fuelbio'].sum() - emissions['old']['fuelbio'].sum(),
               emissions['new']['g2gbio'].sum() - emissions['old']['g2gbio'].sum(),
               emissions['new']['eolbio'].sum() - emissions['old']['eolbio'].sum(), 0]
    fosData = [emissions['forest']['fosGHG'].sum(),
               emissions['new']['fuelfos'].sum() - emissions['old']['fuelfos'].sum(),
               emissions['new']['g2gfos'].sum() - emissions['old']['g2gfos'].sum(),
               emissions['new']['eolfos'].sum() - emissions['old']['eolfos'].sum(), emissions['trade'].sum()]

    # bars from pandas
    labels = ['Forest', 'Upstream', 'Gate-to-Gate', 'EoL', 'International Trade']
    colors = ['#70ad47', '#4472c4']

    df = pd.DataFrame({'Biogenic GHG': bioData, 'Fossil GHG': fosData}, index=labels)
    # ax = df.plot(kind='bar', title='Emission Results Breakdown', rot=0, color=colors)
    df.plot(ax=ax, kind='bar', title='Emission Results Breakdown', rot=0, color=colors)
    ax.set_ylabel('GHG Emission Change (metric tons of CO2 eq/yr)')

def emNet(emissions, ax):
    # fig, ax = plt.subplots()

    colors = ['#5b9bd5']

    bioData = sum(
        [emissions['forest']['bioGHG'].sum(), 0, emissions['new']['g2gbio'].sum() - emissions['old']['g2gbio'].sum(),
         emissions['new']['eolbio'].sum() - emissions['old']['eolbio'].sum(), 0])
    fosData = sum(
        [emissions['forest']['fosGHG'].sum(), 0, emissions['new']['g2gfos'].sum() - emissions['old']['g2gfos'].sum(),
         emissions['new']['eolfos'].sum() - emissions['old']['eolfos'].sum(), emissions['trade'].sum()])

    ax.bar(['Biogenic GHG', 'Fossil GHG'], [bioData, fosData], color=colors)
    # ax.set_title('')
    ax.set_ylabel('Net GHG Emissions Change from Baseline (tons CO2 eq/yr')
