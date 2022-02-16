import pandas as pd
import numpy as np
from scipy.stats import norm

import random
import statistics
import pathlib

def emissions_calculate(
    year_len,
    index_year2018,
    index_year2060,
    Belite_share,Belite_fue_save,Belite_pro_save,
    BYF_share,BYF_fue_save,BYF_pro_save,
    CCSC_share,CCSC_fue_save,CCSC_pro_save,CCSC_CO2_credit,
    CSAB_share,CSAB_fue_save,CSAB_pro_save,
    Celite_share,Celite_fue_save,Celite_pro_save,
    MOMS_share,MOMS_fue_save,MOMS_pro_save,
    Pro_Em_factor,
    Them_eff,
    Coal_share_cem,Coal_CO2_cem,
    Oil_share_cem,Oil_CO2_cem,
    NaGa_share_cem,NaGa_CO2_cem,
    WaFu_share_cem,WaFu_CO2_cem,
    Biom_share_cem,Biom_CO2_cem,
    Ele_eff,
    Coal_share_ele,Coal_CO2_ele,
    Oil_share_ele,Oil_CO2_ele,
    Naga_share_ele,Naga_CO2_ele,
    Hydr_share_ele,Hydr_CO2_ele,
    Geot_share_ele,Geot_CO2_ele,
    Sol_share_ele,Sol_CO2_ele,
    Wind_share_ele,Wind_CO2_ele,
    Tide_share_ele,Tide_CO2_ele,
    Nuc_share_ele,Nuc_CO2_ele,
    Bio_share_ele,Bio_CO2_ele,
    Was_share_ele,Was_CO2_ele,
    SoTh_share_ele,SoTh_CO2_ele,
    clinker,
    Cla_share,Ene_Cla,
    CCS_Oxy_percent,Eff_Oxy,Ene_Oxy,
    CCS_Pos_percent,Eff_Pos,Ene_Pos,
    Cement_demand,
    CO2_tran_cem,
    Concrete_demand,
    CO2_prod_agg,CO2_tran_agg,
    Concrete_demolish,
    recRate,CO2_prod_rec,CO2_tran_rec,
    burRate,CO2_tran_bur,
    CO2_mix_con,CO2_tran_con,
    CO2_onsite_con,
    CO2_curing_emi,
    CO2_mine_EOL_emi,
    CO2_mine_slag_emi,CO2_mine_slag_sto,
    CO2_mine_ash_emi,CO2_mine_ash_sto,
    CO2_mine_lime_emi,CO2_mine_lime_sto,
    CO2_mine_red_emi,CO2_mine_red_sto,
    CO2_mine_slag_share,
    CO2_mine_ash_share,
    CO2_mine_lime_share,
    CO2_mine_red_share,
    Timber_demand,
    CO2_tim_pen,CO2_tim_sto,
    Timber_demolish,
    EoL_tim_land,Land_tim_deg,
    Deg_tim_CH4,Deg_tim_CO2,
    CH4_to_CO2,CH4_recover,
    EoL_tim_inci,
    scenario_index,
    storyline,
    filepath
):
    ##############################################################################
    #simulate cement CO2 emissions; 2018-2060; 43 yrs
    ##############################################################################
    Chem_fue_save = Belite_share*Belite_fue_save + BYF_share*BYF_fue_save +\
        CCSC_share*CCSC_fue_save + CSAB_share*CSAB_fue_save +\
            Celite_share*Celite_fue_save + MOMS_share*MOMS_fue_save
    Chem_fue_save = Chem_fue_save[index_year2018:index_year2060]

    Chem_pro_save = Belite_share*Belite_pro_save + BYF_share*BYF_pro_save +\
        CCSC_share*CCSC_pro_save + CSAB_share*CSAB_pro_save +\
            Celite_share*Celite_pro_save + MOMS_share*MOMS_pro_save
    Chem_pro_save = Chem_pro_save[index_year2018:index_year2060]

    Chem_CO2_storage = CCSC_share*CCSC_CO2_credit
    Chem_CO2_storage = Chem_CO2_storage[index_year2018:index_year2060]

    #Process emission when green cement chemistries are implemented
    Pro_Em_factor = Pro_Em_factor*(1-Chem_pro_save)

    #CO2 emissions from biomass are not accounted following the BECCS accounting principle
    #kg CO2/t clinker
    Fue_Em_factor = Them_eff*(Coal_share_cem*Coal_CO2_cem + Oil_share_cem*Oil_CO2_cem + NaGa_share_cem*NaGa_CO2_cem +\
        WaFu_share_cem*WaFu_CO2_cem)*(1-Chem_fue_save)
    Biom_Em_factor = Them_eff*Biom_share_cem*Biom_CO2_cem*(1-Chem_fue_save)

    #kg CO2/t cement
    Ele_Em_factor = Ele_eff*(Coal_share_ele*Coal_CO2_ele+Oil_share_ele*Oil_CO2_ele+Naga_share_ele*Naga_CO2_ele+\
        Hydr_share_ele*Hydr_CO2_ele+Geot_share_ele*Geot_CO2_ele+Sol_share_ele*Sol_CO2_ele+\
            Wind_share_ele*Wind_CO2_ele+Tide_share_ele*Tide_CO2_ele+Nuc_share_ele*Nuc_CO2_ele+\
                Bio_share_ele*Bio_CO2_ele+Was_share_ele*Was_CO2_ele+SoTh_share_ele*SoTh_CO2_ele)/1000

    #Total CO2 emissions; including energy use of producing calcined clay; including CO2 storage of CCSC
    #kg CO2/t cement
    Tot_Em_factor = (Pro_Em_factor+Fue_Em_factor-Chem_CO2_storage)*clinker[index_year2018:index_year2060] +\
        Ele_Em_factor + Cla_share*Ene_Cla*Fue_Em_factor

    #CO2 emissions due to energy penalty
    CO2_Oxy = CCS_Oxy_percent*Eff_Oxy*Ene_Oxy*(Coal_share_cem*Coal_CO2_cem + Oil_share_cem*Oil_CO2_cem +\
        NaGa_share_cem*NaGa_CO2_cem + WaFu_share_cem*WaFu_CO2_cem)*(Pro_Em_factor+Fue_Em_factor)*clinker[index_year2018:index_year2060]/1000

    CO2_Pos = CCS_Pos_percent*Eff_Pos*Ene_Pos*(Coal_share_cem*Coal_CO2_cem + Oil_share_cem*Oil_CO2_cem +\
        NaGa_share_cem*NaGa_CO2_cem + WaFu_share_cem*WaFu_CO2_cem)*(Pro_Em_factor+Fue_Em_factor)*clinker[index_year2018:index_year2060]/1000

    #Total CO2 emissions with CCS
    Tot_Em_factor_CCS = Tot_Em_factor+CO2_Oxy+CO2_Pos-\
        (Pro_Em_factor+Fue_Em_factor+Biom_Em_factor)*(CCS_Oxy_percent*Eff_Oxy+CCS_Pos_percent*Eff_Pos)*clinker[index_year2018:index_year2060]

    Em_factor_list = pd.concat([Tot_Em_factor_CCS, Tot_Em_factor,\
                                CO2_Oxy, CO2_Pos,\
                                Pro_Em_factor, Fue_Em_factor, Biom_Em_factor],
                    axis = 1, ignore_index = True)
    Em_factor_list.columns = ['Tot_Em_factor_CCS','Tot_Em_factor',\
                                'CO2_Oxy','CO2_Pos',\
                                    'Pro_Em_factor','Fue_Em_factor','Biom_Em_factor']

    CO2_emissions_cement_prod   = Cement_demand[index_year2018:index_year2060] * Tot_Em_factor_CCS/1000
    CO2_emissions_cement_tran   = Cement_demand[index_year2018:index_year2060] * CO2_tran_cem[index_year2018:index_year2060]
    CO2_emissions_aggreg_prod   = (Concrete_demand[index_year2018:index_year2060] - Cement_demand[index_year2018:index_year2060] -\
        Concrete_demolish[index_year2018:index_year2060] * recRate[index_year2018:index_year2060] -\
            Concrete_demand[index_year2018:index_year2060] * CO2_mine_slag_share[index_year2018:index_year2060] -\
                Concrete_demand[index_year2018:index_year2060] * CO2_mine_ash_share[index_year2018:index_year2060] -\
                    Concrete_demand[index_year2018:index_year2060] * CO2_mine_lime_share[index_year2018:index_year2060] -\
                        Concrete_demand[index_year2018:index_year2060] * CO2_mine_red_share[index_year2018:index_year2060])\
                            * CO2_prod_agg[index_year2018:index_year2060]
    CO2_emissions_aggreg_tran   = (Concrete_demand[index_year2018:index_year2060] - Cement_demand[index_year2018:index_year2060] -\
        Concrete_demolish[index_year2018:index_year2060] * recRate[index_year2018:index_year2060]) * CO2_tran_agg[index_year2018:index_year2060]
    CO2_emissions_recagg_prod   = Concrete_demolish[index_year2018:index_year2060] * recRate[index_year2018:index_year2060] * CO2_prod_rec[index_year2018:index_year2060]
    CO2_emissions_recagg_tran   = Concrete_demolish[index_year2018:index_year2060] * recRate[index_year2018:index_year2060] * CO2_tran_rec[index_year2018:index_year2060]
    CO2_emissions_buragg_tran   = Concrete_demolish[index_year2018:index_year2060] * burRate[index_year2018:index_year2060] * CO2_tran_bur[index_year2018:index_year2060]
    CO2_emissions_concrete_mix    = Concrete_demand[index_year2018:index_year2060] * CO2_mix_con[index_year2018:index_year2060]
    CO2_emissions_concrete_tran   = Concrete_demand[index_year2018:index_year2060] * CO2_tran_con[index_year2018:index_year2060]
    CO2_emissions_concrete_onsite = Concrete_demand[index_year2018:index_year2060] * CO2_onsite_con[index_year2018:index_year2060]

    CO2_emissions_CO2_curing    = Concrete_demand[index_year2018:index_year2060] * CO2_curing_emi[index_year2018:index_year2060]
    CO2_emissions_CO2_mine_EOL  = Concrete_demolish[index_year2018:index_year2060] * recRate[index_year2018:index_year2060] * CO2_mine_EOL_emi[index_year2018:index_year2060]
    CO2_emissions_CO2_mine_slag = Concrete_demand[index_year2018:index_year2060] * CO2_mine_slag_emi[index_year2018:index_year2060]
    CO2_emissions_CO2_mine_ash  = Concrete_demand[index_year2018:index_year2060] * CO2_mine_ash_emi[index_year2018:index_year2060]
    CO2_emissions_CO2_mine_lime = Concrete_demand[index_year2018:index_year2060] * CO2_mine_lime_emi[index_year2018:index_year2060]
    CO2_emissions_CO2_mine_red  = Concrete_demand[index_year2018:index_year2060] * CO2_mine_red_emi[index_year2018:index_year2060]

    CO2_emissions_timber = Timber_demand[index_year2018:index_year2060] * CO2_tim_pen[index_year2018:index_year2060]

    CO2_emissions = CO2_emissions_cement_prod + CO2_emissions_cement_tran +\
                    CO2_emissions_aggreg_prod + CO2_emissions_aggreg_tran +\
                    CO2_emissions_recagg_prod + CO2_emissions_recagg_tran +\
                    CO2_emissions_buragg_tran + CO2_emissions_concrete_mix+\
                    CO2_emissions_concrete_tran+CO2_emissions_concrete_onsite+\
                    CO2_emissions_CO2_curing  + CO2_emissions_CO2_mine_EOL +\
                    CO2_emissions_CO2_mine_slag + CO2_emissions_CO2_mine_ash +\
                    CO2_emissions_CO2_mine_lime + CO2_emissions_CO2_mine_red +\
                    CO2_emissions_timber

    CO2_emissions_list = pd.concat([CO2_emissions_cement_prod, CO2_emissions_cement_tran,\
                    CO2_emissions_aggreg_prod, CO2_emissions_aggreg_tran,\
                    CO2_emissions_recagg_prod, CO2_emissions_recagg_tran,\
                    CO2_emissions_buragg_tran, CO2_emissions_concrete_mix,\
                    CO2_emissions_concrete_tran, CO2_emissions_concrete_onsite,\
                    CO2_emissions_CO2_curing, CO2_emissions_CO2_mine_EOL,\
                    CO2_emissions_CO2_mine_slag, CO2_emissions_CO2_mine_ash,\
                    CO2_emissions_CO2_mine_lime, CO2_emissions_CO2_mine_red,\
                    CO2_emissions_timber],
                    axis = 1, ignore_index = True)
    CO2_emissions_list.columns = ['cement_prod', 'cement_tran',\
                    'aggreg_prod', 'aggreg_tran',\
                    'recagg_prod', 'recagg_tran',\
                    'buragg_tran', 'concrete_mix',\
                    'concrete_tran', 'concrete_onsite',\
                    'CO2_curing', 'CO2_mine_EOL',\
                    'CO2_mine_slag', 'CO2_mine_ash',\
                    'CO2_mine_lime', 'CO2_mine_red',\
                    'timber']

    CO2_storage_CO2_mine_slag = Concrete_demand[index_year2018:index_year2060] * CO2_mine_slag_sto[index_year2018:index_year2060]
    CO2_storage_CO2_mine_ash  = Concrete_demand[index_year2018:index_year2060] * CO2_mine_ash_sto[index_year2018:index_year2060]
    CO2_storage_CO2_mine_lime = Concrete_demand[index_year2018:index_year2060] * CO2_mine_lime_sto[index_year2018:index_year2060]
    CO2_storage_CO2_mine_red  = Concrete_demand[index_year2018:index_year2060] * CO2_mine_red_sto[index_year2018:index_year2060]

    CO2_storage_timber = Timber_demand[index_year2018:index_year2060] * CO2_tim_sto[index_year2018:index_year2060] -\
        Timber_demolish[index_year2018:index_year2060]*CO2_tim_sto[index_year2018:index_year2060]*12/44*\
            (EoL_tim_land[index_year2018:index_year2060]*Land_tim_deg[index_year2018:index_year2060]*Deg_tim_CH4[index_year2018:index_year2060]*16/12*CH4_to_CO2[index_year2018:index_year2060]*(1-CH4_recover[index_year2018:index_year2060])+\
                EoL_tim_land[index_year2018:index_year2060]*Land_tim_deg[index_year2018:index_year2060]*Deg_tim_CO2[index_year2018:index_year2060]*44/12+\
                    EoL_tim_inci[index_year2018:index_year2060]*44/12)   

    CO2_storage = CO2_storage_CO2_mine_slag + CO2_storage_CO2_mine_ash +\
                CO2_storage_CO2_mine_lime + CO2_storage_CO2_mine_red +\
                CO2_storage_timber

    CO2_storage_list = pd.concat([CO2_storage_CO2_mine_slag, CO2_storage_CO2_mine_ash,\
                CO2_storage_CO2_mine_lime, CO2_storage_CO2_mine_red,\
                CO2_storage_timber],
                axis = 1, ignore_index = True)
    CO2_storage_list.columns = ['CO2_mine_slag', 'CO2_mine_ash',\
                'CO2_mine_lime', 'CO2_mine_red',\
                'CO2_storage_timber']

    pd.DataFrame(Em_factor_list).to_excel(
        'C:\\Users\\caozh\\Documents\\2020-ClimateWork\\Carbonation\\' + filepath + \
                '\\Em_factor_list_' + str(scenario_index+"_"+storyline) + '.xlsx')

    pd.DataFrame(CO2_emissions).to_excel(
        'C:\\Users\\caozh\\Documents\\2020-ClimateWork\\Carbonation\\' + filepath + \
                '\\CO2_emissions_' + str(scenario_index+"_"+storyline) + '.xlsx')

    pd.DataFrame(CO2_emissions_list).to_excel(
        'C:\\Users\\caozh\\Documents\\2020-ClimateWork\\Carbonation\\' + filepath + \
                '\\CO2_emissions_list_' + str(scenario_index+"_"+storyline) + '.xlsx')

    pd.DataFrame(CO2_storage).to_excel(
        'C:\\Users\\caozh\\Documents\\2020-ClimateWork\\Carbonation\\' + filepath + \
                '\\CO2_storage_' + str(scenario_index+"_"+storyline) + '.xlsx')

    pd.DataFrame(CO2_storage_list).to_excel(
        'C:\\Users\\caozh\\Documents\\2020-ClimateWork\\Carbonation\\' + filepath + \
                '\\CO2_storage_list_' + str(scenario_index+"_"+storyline) + '.xlsx')
