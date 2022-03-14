import streamlit as st # need 0.89.0, 1.7.0 too late
import numpy as np
import pandas as pd # need at least 1.4.1
import matplotlib.pyplot as plt
from fromTerminal import *
from impData import *
from createModel import *
from scenarios import *
from plots import *
from gurobipy import GRB


## BUTTON DEFINITIONS
def load_data(file,fProd):
    if 'data' not in st.session_state:
        st.session_state.data = impData(file, fProd)

def create_model(data,fProd):
    # if 'model' not in st.session_state:
    st.session_state.model = createModel('2019data', data, fProd)

def apply_settings(scene, scene_tog):
    if 'model' in st.session_state:
        if scene == 'Increase average recycled content':
            setfAvg(st.session_state.model, target, incr)

        elif scene == 'Decrease mixed paper yield':
            setmpYield(st.session_state.model, perc)
        elif scene == 'Change demand':
            setDemand(st.session_state.model, prod, perc)

        if 'Contamination by recovery' in scene_tog:
            setContam(st.session_state.model, recovery, contam_num)

        if 'Fiber degradation' in scene_tog:
            fiberDeg(st.session_state.model, fibdeg_num)

# def reset_settings(model):
#     st.session_state.model = resetScenarios(model)
#     # scene_tog.clear()

def run_optimization(model, plc1):
    with st_stdout("info", plc1):
        model.runModel()
    if model.solved:
        model.getResults()

# def clear_model(model):
#     model.m.reset(0) # reset model to unsolved state
#     reset_settings(model)


## FRONT END STARTS HERE
st.title('Fiber Distribution Model')

fProd = ['Containerboard', 'Paperboard', 'Tissue-Away', 'Tissue-Home', 'P&W', 'Newsprint']
fBox = st.multiselect('Select products to be included in system: ', options=fProd, default=fProd, key='fProd')

# load data
uploaded_file = st.file_uploader("Select spreadsheet to be analyzed: ", type=['.xls', '.xlsx', '.csv'], key='upload')
if uploaded_file is not None:
    ld_col1, ld_col2 = st.columns(2)
    with ld_col1:
        ld_but = st.button("Load data", key='ld_but', on_click=load_data, args=(uploaded_file, fBox, ))
    with ld_col2:
        if ld_but:
            st.success('Data loaded')
else:
    st.text("Please select spreadsheet.")

if 'data' in st.session_state:
    # create Gurobi model
    # model_created = False
    cm_col1, cm_col2 = st.columns(2)
    with cm_col1:
        cm_but = st.button("Create Gurobi optimization model", key='cm_but', on_click=create_model, args=(st.session_state.data, fBox, ))
    with cm_col2:
        if cm_but:
            # model_created = True
            st.success('Model created')

    if 'model' in st.session_state:
        # offer scenarios to run
        sc_all = st.write("## Scenarios")

        scene = st.radio(
            "Which scenario would you like to run?",
            ('Increase average recycled content', 'Decrease mixed paper yield', 'Change demand'), key='scene')

        sc_screen1 = st.empty()
        sc_screen2 = st.empty()
        sc_screen3 = st.empty()
        sc_screen4 = st.empty()


        if scene == 'Increase average recycled content':
            target = sc_screen1.selectbox('Which product sector will be affected?', options=fBox)
            incr = sc_screen2.slider('By how much should the average recycled content be increased?', 0.0, 15.0, 2.0, 0.1,
                              format='%.1f%%')
            sc_screen3.text(f"Increasing average recycled content in {target} from {st.session_state.model.rAvg[target] * 100}% to {st.session_state.model.rAvg[target] * 100 + incr}%\t\nof the total domestic consumption of {target}.")
        elif scene == 'Decrease mixed paper yield':
            perc = sc_screen1.slider('By how much should the mixed paper yield be decreased?', -20.0, 0.0, 0.0, 0.1, format='%.1f%%')
        elif scene == 'Change demand':
            prod = sc_screen1.selectbox('Which product sector will be affected?', options=fBox)
            perc = sc_screen2.slider('By how much should the demand be changed?', -20.0, 20.0, 2.0, 0.1, format='%.1f%%')
        else:
            st.write('Please select a scenario to run.')

        scene_tog = st.multiselect("More scenario settings to consider:",
                                   options=['Contamination by recovery', 'Fiber degradation'], key='scene_tog')
        if 'Contamination by recovery' in scene_tog:
            recovery = sc_screen4.selectbox('Which recovery system will be used?', options=['ssr', 'dual', 'mix'])
            st.text("The contamination rate is defined here as the fraction of recovered paper that is \nunusable and "
                    "must be discarded due to contaminants introduced through consumption \nand collection in the "
                    "Residential and Retail consumption channels. This value may \nbe specified below or left to a "
                    "stochastic fraction bounded by literature values. \nFor example, a contamination rate of 24.6% "
                    "is 0.246.")
            contam_num = st.number_input(label="(optional) Specify the contamination rate of recovered paper collected "
                                               f'through a {recovery.upper()} recovery system',
                                         min_value=0.001, max_value=1.0, key='contam_num')

        if 'Fiber degradation' in scene_tog:
            st.text("Fiber degradaton is defined here as the fraction of fiber from a particular \nrecovered paper "
                    "product too short to be reused in papermaking. Fiber yields will be \nreduced by some stochastic "
                    "percent between 12.5% and 25% by fiber grade category \nunless otherwise specified.")
            fibdeg_num = st.number_input(label="(optional) Specify the rate of fiber degradation of recovered paper",
                                         min_value=0.001, max_value=1.0, key='fibdeg_num')

        st.text('Once all settings have been set, click "Apply scenarios settings" to save to model.')


        sc_col1, sc_col2 = st.columns(2)
        with sc_col1:
            sc_set_but = st.button("Apply scenarios settings", on_click=apply_settings, args=(st.session_state.scene, st.session_state.scene_tog, ))

            allConstr = {}
            for i in [i for i in st.session_state.model.m.getConstrs()]:
                name = i.getAttr(GRB.Attr.ConstrName).split("[", 1)
                if name[0] in allConstr.keys():
                    allConstr[name[0]].append(i)
                else:
                    allConstr.update({name[0]: [i]})

            # st.write(allConstr)

        # with sc_col2: # reset settings to default
        #     sc_reset_but = st.button("Reset settings", on_click=reset_settings, args=(st.session_state.model, ))
        #
        #     allCLen = {'TOTAL': 0}
        #     for key in allConstr.keys():
        #         allCLen.update({key: len(allConstr[key])})
        #         allCLen.update(({'TOTAL': allCLen['TOTAL'] + len(allConstr[key])}))

            # st.write(allCLen)

        # optimize and get results
        opt_all = st.write('## Fiber Allocation')

        sc_screen5 = st.empty()
        opt_col1, opt_col2 = st.columns(2)
        with opt_col1:
            om_but = st.button("Solve fiber distribution model", key='om_but', on_click=run_optimization, args=(st.session_state.model, sc_screen5, ))

        # with opt_col2: # prepare to create new model
        #     rm_but = st.button("Clear model", key='rm_but', on_click=clear_model, args=(st.session_state.model, ))

        st.session_state.fail = ""
        if st.session_state.model.solved:  # only display if optimal solutions have been obtained
            st.session_state.fail = ""
            # plot data
            plot_all = st.expander('Plot results')
            with plot_all:
                plt_col1, plt_col2 = st.columns(2)
                with plt_col1:
                    st.write('### Fiber Distribution')
                    fig1, ax1 = plt.subplots()
                    fUseChange(st.session_state.model, ax1)
                    st.pyplot(fig1)

                    fig2, ax2 = plt.subplots()
                    fUseByProd(st.session_state.model, ax2)
                    st.pyplot(fig2)

                    fig3, ax3 = plt.subplots()
                    fSources(st.session_state.model, ax3)
                    st.pyplot(fig3)

                    fig4, ax4 = plt.subplots()
                    mxpShift(st.session_state.model, ax4)
                    st.pyplot(fig4)

                    # fig5, ax5 = plt.subplots()
                    # fig6, ax6 = plt.subplots()
                    # fig7, ax7 = plt.subplots()
                    # fig8, ax8 = plt.subplots()
                    # fig9, ax9 = plt.subplots()
                    # fShiftByCat(st.session_state.model, [ax5, ax6, ax7, ax8, ax9])
                    # st.pyplot(fig5)
                    # st.pyplot(fig6)
                    # st.pyplot(fig7)
                    # st.pyplot(fig8)
                    # st.pyplot(fig9)

                with plt_col2:
                    st.write('### Energy & Emissions')
                    fig10, ax10 = plt.subplots()
                    dirEnergyCons(st.session_state.model.emissions, ax10)
                    st.pyplot(fig10)

                    fig11, ax11 = plt.subplots()
                    ghg_notForEoL(st.session_state.model.emissions, ax11)
                    st.pyplot(fig11)

                    fig12, ax12 = plt.subplots()
                    ghg_wForEoL(st.session_state.model.emissions, ax12)
                    st.pyplot(fig12)

                    fig13, ax13 = plt.subplots()
                    emBreakdown(st.session_state.model.emissions, ax13)
                    st.pyplot(fig13)

                    fig14, ax14 = plt.subplots()
                    emNet(st.session_state.model.emissions, ax14)
                    st.pyplot(fig14)


            # make data available to save
            data_all = st.expander('Save results')
            with data_all: # TODO: write to new worksheet on Excel file & would it print right if fProd changes?
                # save to Excel file
                sv_col1, sv_col2 = st.columns(2)
                with sv_col1:
                    sv_but = st.button("Save results to spreadsheet", key="sv_but")
                    st.session_state.save = sv_but
                    if sv_but:
                        with pd.ExcelWriter(uploaded_file.name, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:

                            st.session_state.model.f2pVolNew.to_excel(writer, sheet_name='Results-FiberPulp', startrow=2)
                            st.session_state.model.pbpVolNew.to_excel(writer, sheet_name='Results-FiberPulp', startrow=2, startcol=10)

                            pd.DataFrame.from_dict(st.session_state.model.demandNew).to_excel(writer, sheet_name='Results-Demand', startrow=2)
                            st.session_state.model.collNew.to_excel(writer, sheet_name='Results-Demand', startrow=13)
                            st.session_state.model.exportNew.to_excel(writer, sheet_name='Results-Demand', startrow=22)
                            st.session_state.model.rsdlNew.to_excel(writer, sheet_name='Results-Demand', startrow=22, startcol=5)

                            st.session_state.model.emissions['old']['bioCO2'].to_excel(writer, sheet_name='Results-GHG',
                                                                                        startrow=2)
                            st.session_state.model.emissions['old']['fossilCO2'].to_excel(writer, sheet_name='Results-GHG',
                                                                                       startrow=11)
                            st.session_state.model.emissions['old']['g2gbio'].to_excel(writer,
                                                                                          sheet_name='Results-GHG',
                                                                                          startrow=20)
                            st.session_state.model.emissions['old']['g2gfos'].to_excel(writer,
                                                                                          sheet_name='Results-GHG',
                                                                                          startrow=29)
                            st.session_state.model.emissions['old']['totalImp'].to_excel(writer,
                                                                                       sheet_name='Results-GHG',
                                                                                       startrow=38)
                            st.session_state.model.emissions['old']['unitImp'].to_excel(writer,
                                                                                         sheet_name='Results-GHG',
                                                                                         startrow=46)

                            st.session_state.model.emissions['new']['bioCO2'].to_excel(writer, sheet_name='Results-GHG',
                                                                                       startrow=2, startcol=3)
                            st.session_state.model.emissions['new']['fossilCO2'].to_excel(writer,
                                                                                          sheet_name='Results-GHG',
                                                                                          startrow=11, startcol=3)
                            st.session_state.model.emissions['new']['g2gbio'].to_excel(writer,
                                                                                       sheet_name='Results-GHG',
                                                                                       startrow=20, startcol=3)
                            st.session_state.model.emissions['new']['g2gfos'].to_excel(writer,
                                                                                       sheet_name='Results-GHG',
                                                                                       startrow=29, startcol=3)
                            st.session_state.model.emissions['new']['totalImp'].to_excel(writer,
                                                                                         sheet_name='Results-GHG',
                                                                                         startrow=38, startcol=3)
                            st.session_state.model.emissions['new']['unitImp'].to_excel(writer,
                                                                                        sheet_name='Results-GHG',
                                                                                        startrow=46, startcol=3)

                            st.session_state.model.emissions['forest']['bioGHG'].to_excel(writer,
                                                                                        sheet_name='Results-GHG',
                                                                                        startrow=57, startcol=0)
                            st.session_state.model.emissions['forest']['fosGHG'].to_excel(writer,
                                                                                          sheet_name='Results-GHG',
                                                                                          startrow=57, startcol=3)

                            st.session_state.model.emissions['trade'].to_excel(writer,
                                                                               sheet_name='Results-GHG',
                                                                               startrow=61, startcol=0)

                            st.session_state.model.emissions['oldenergy'].to_excel(writer,
                                                                               sheet_name='Results-GHG',
                                                                               startrow=65, startcol=0)
                            st.session_state.model.emissions['newenergy'].to_excel(writer,
                                                                                   sheet_name='Results-GHG',
                                                                                   startrow=65, startcol=6)

                with sv_col2:
                    if st.session_state.save:
                        st.success('Results saved!')

                # display data
                st.write('New fiber-to-pulp distribution (short tons)')
                st.dataframe(st.session_state.model.f2pVolNew.style.format('{:,.0f}', precision=0))
                # st.dataframe(st.session_state.model.f2pVolNew.iloc[:16,:].sum() - st.session_state.model.f2pVolOld.iloc[:16,:].sum())
                # st.dataframe(st.session_state.model.f2pVolNew.iloc[16:,:].sum() - st.session_state.model.f2pVolOld.iloc[16:,:].sum())
                # st.write(f"Total virgin pulpwood (new) = {sum(st.session_state.model.f2pVolNew.iloc[16:,:4].sum()):,.0f}")
                # st.write(f"Total virgin pulpwood (new) = {sum(st.session_state.model.f2pVolNew.iloc[16:, :4].sum()):,.0f}")
                st.write('Old fiber-to-pulp distribution (short tons)')
                st.dataframe(st.session_state.model.f2pVolOld.style.format('{:,.0f}', precision=0))
                # st.write(f"Total virgin pulpwood (old) = {sum(st.session_state.model.f2pVolOld.iloc[16:, :4].sum()):,.0f}")
                # st.write(f"Total virgin pulpwood (old) = {sum(st.session_state.model.f2pVolOld.iloc[16:, :4].sum()):,.0f}")

                st.write('New pulp distribution by product (short tons)')
                st.dataframe(st.session_state.model.pbpVolNew.style.format('{:,.0f}', precision=0))
                # st.write(f"Total deinked market pulp = {st.session_state.model.pbpVolNew.loc['RecPulp_Deinked'].sum():,.0f}")
                # st.write(f"Total virgin market pulp = {st.session_state.model.pbpVolNew.loc['VirPulp_Market'].sum():,.0f}")
                st.write('Old pulp distribution by product (short tons)')
                st.dataframe(st.session_state.model.pbpVolOld.style.format('{:,.0f}', precision=0))
                # st.write(f"Total deinked market pulp = {st.session_state.model.pbpVolOld.loc['RecPulp_Deinked'].sum():,.0f}")
                # st.write(f"Total virgin market pulp = {st.session_state.model.pbpVolOld.loc['VirPulp_Market'].sum():,.0f}")

                st.write('New demand by product (short tons)')
                # model.demandNew # dict
                st.dataframe(pd.DataFrame.from_dict(st.session_state.model.demandNew).fillna(0).style.format('{:,.0f}', precision=0))
                # st.write(f"Total demand = {pd.DataFrame.from_dict(st.session_state.model.demandNew).fillna(0).sum()}")
                st.write('Old demand by product (short tons)')
                st.dataframe(pd.DataFrame.from_dict(st.session_state.model.oldDemand).fillna(0).style.format('{:,.0f}', precision=0))

                st.write('New collection from consumption channels')
                st.dataframe(st.session_state.model.collNew.style.format('{:.1%}'))
                st.write('Change in collection from consumption channels')
                st.dataframe(st.session_state.model.collDelta.style.format('{:.1%}'))
                st.write('Additional collection from consumption channels')
                st.dataframe(st.session_state.model.addlRec)
                st.write(f"Total additional collection from consumption channels (short tons): {st.session_state.model.addlRec.sum()}")
                # st.write(pd.DataFrame.from_dict(st.session_state.data.colCost))

                # st.write('Old collection from consumption channels')
                # model.sCollectOld

                st.write('New export volume by fiber grade (short tons)')
                st.dataframe(st.session_state.model.exportNew.style.format('{:,.0f}', precision=0))
                st.write('Change in export volume by fiber grade (short tons)')
                st.dataframe(pd.DataFrame(st.session_state.model.exportNew.values - st.session_state.model.exportOld.values,
                                          index=st.session_state.model.exportNew.index).style.format('{:,.0f}', precision=0))
                st.write(f"Total shift in export (short tons): {st.session_state.model.exportNew.values.sum() - st.session_state.model.exportOld.values.sum()}")

                st.write('New residuals volume by fiber grade (short tons)')
                st.dataframe(pd.DataFrame(st.session_state.model.rsdlNew).style.format('{:,.0f}', precision=0))

                st.write('New collection and recovery volume (short tons)')
                st.dataframe(st.session_state.model.consCollNew.style.format('{:,.0f}', precision=0))

                st.write('Direct energy consumption at mills (GJ)')
                st.dataframe(st.session_state.model.emissions['new']['energybio'] - st.session_state.model.emissions['old']['energybio'])
                st.dataframe(st.session_state.model.emissions['new']['energyfos'] - st.session_state.model.emissions['old']['energyfos'])

                st.write('Net GHG emissions (tons CO2 eq/yr)')
                st.dataframe(st.session_state.model.emissions['forest']['bioGHG'])
                bioEm = st.session_state.model.emissions['new']['g2gbio'] - st.session_state.model.emissions['old']['g2gbio'] + \
                        st.session_state.model.emissions['new']['eolbio'] - st.session_state.model.emissions['old']['eolbio']
                st.dataframe(bioEm)
                st.write(f"Total bioEm is {bioEm.sum()}")
                st.dataframe(st.session_state.model.emissions['forest']['bioGHG'] + bioEm.sum())
                st.dataframe(st.session_state.model.emissions['forest']['fosGHG'])
                fosEm = st.session_state.model.emissions['new']['g2gfos'] - st.session_state.model.emissions['old']['g2gfos'] + \
                        st.session_state.model.emissions['new']['eolfos'] - st.session_state.model.emissions['old']['eolfos']
                st.dataframe(fosEm)
                st.write(f"Total fosEm is {fosEm.sum()}")
                st.write(f"fosEm from trade is {st.session_state.model.emissions['trade'].sum()}")
                st.dataframe(st.session_state.model.emissions['forest']['fosGHG'] + fosEm.sum() + st.session_state.model.emissions['trade'].sum())


        else:
            if 'model' in st.session_state and st.session_state.model.failed:
                # st.write('No solution found. Please adjust the settings and try again.')
                st.session_state.fail = "No solution found. Please adjust the settings and try again."

            st.write(st.session_state.fail)

