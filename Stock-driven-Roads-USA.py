import pandas as pd
import numpy as np
from scipy.stats import norm

import random
import statistics
import pathlib

import matplotlib.pyplot as plt

def movingaverage (values, window):
    weights = np.repeat(1.0, window)/window
    sma = np.convolve(values, weights, 'valid')
    return sma

##############################################################################
############################stock-driven model################################
##############################################################################
#read new floor area 1931-2017 from "Data input.xlsx"
scenario_frozen_inputs = pd.read_excel (io = r'C:\Users\caozh\Box\ERSAL Cement 2018-2019\Zhi Cao\Raw data\Data inputs.xlsx', 
header = None, sheet_name="USA-Roads", usecols= "B:GX", nrows = 161, skiprows = 4)

scenario_steps_inputs = pd.read_excel (io = r'C:\Users\caozh\Box\ERSAL Cement 2018-2019\Zhi Cao\Raw data\Data inputs.xlsx', 
header = None, sheet_name="USA-Roads-STEPS", usecols= "B:GX", nrows = 161, skiprows = 4)

scenario_production_inputs = pd.read_excel (io = r'C:\Users\caozh\Box\ERSAL Cement 2018-2019\Zhi Cao\Raw data\Data inputs.xlsx', 
header = None, sheet_name="USA-Roads-Production", usecols= "B:GX", nrows = 161, skiprows = 4)

scenario_demand_inputs = pd.read_excel (io = r'C:\Users\caozh\Box\ERSAL Cement 2018-2019\Zhi Cao\Raw data\Data inputs.xlsx', 
header = None, sheet_name="USA-Roads-Demand", usecols= "B:GX", nrows = 161, skiprows = 4)

lever_list = {
     'Base': [],
     'L1a': [88],
     'L1b': [113],
     'L1c': [89,90,91,92,93],
     'L3':  [99,100,101,102,103,104,105],
     'L4':  [107,108,109,110],
     'L2':  [138,139,140,141,142,143],
     'L5a': [163],
     'L5f': [167,171,174,177,180],
     'L6a': [183],
     'L6b': [185],
     'L6c': [189],
     'L6d': [16,17,18,19],
     'L6e': [36],
     'L7a': [191],
     'L7b': [192],
     'L7c': [204],
     'L8a': [114,115,116,117,118,119,120,121,122,123,124,125],
     'L8b': [194,195,196,197,198,199,200,201,202]
    }

storylines = ['STEPS','Production','Demand']
#data_inputs = scenario_frozen_inputs
#storylines = ['Production','Demand']

for storyline in storylines:
    print("storyline is " + storyline)

    if storyline == 'STEPS':
        lever_list_story = {
            'Base':lever_list['Base'],
            'L1a': lever_list['Base']+lever_list['L1a'],
            'L1b': lever_list['Base']+lever_list['L1a']+lever_list['L1b'],
            'L1c': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c'],
            'L3':  lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3'],
            'L4':  lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L4']
        }
        data_inputs = scenario_frozen_inputs.copy()
        scenario_inputs = scenario_steps_inputs.copy()

    if storyline == 'Production':
        lever_list_story = {
            'Base':lever_list['Base'],
            'L1a': lever_list['Base']+lever_list['L1a'],
            'L1b': lever_list['Base']+lever_list['L1a']+lever_list['L1b'],
            'L1c': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c'],
            'L3':  lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3'],
            'L2':  lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2'],
            'L5a': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a'],
            'L5f': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f'],
            'L4':  lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f']+lever_list['L4'],
            'L8a': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f']+lever_list['L4']+lever_list['L8a'],
            'L8b': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f']+lever_list['L4']+lever_list['L8a']+lever_list['L8b']
        }
        data_inputs = scenario_steps_inputs.copy()
        scenario_inputs = scenario_production_inputs.copy()

    if storyline == 'Demand':
        lever_list_story = {
            'Base':lever_list['Base'],
            'L1a': lever_list['Base']+lever_list['L1a'],
            'L1b': lever_list['Base']+lever_list['L1a']+lever_list['L1b'],
            'L1c': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c'],
            'L3':  lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3'],
            'L2':  lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2'],
            'L5a': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a'],
            'L5f': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f'],
            'L4':  lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f']+lever_list['L4'],
            'L6a': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f']+lever_list['L4']+lever_list['L6a'],
            'L6b': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f']+lever_list['L4']+lever_list['L6a']+lever_list['L6b'],
            'L6c': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f']+lever_list['L4']+lever_list['L6a']+lever_list['L6b']+lever_list['L6c'],
            'L6d': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f']+lever_list['L4']+lever_list['L6a']+lever_list['L6b']+lever_list['L6c']+\
                    lever_list['L6d'],
            'L6e': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f']+lever_list['L4']+lever_list['L6a']+lever_list['L6b']+lever_list['L6c']+\
                    lever_list['L6d']+lever_list['L6e'],
            'L7a': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f']+lever_list['L4']+lever_list['L6a']+lever_list['L6b']+lever_list['L6c']+\
                    lever_list['L6d']+lever_list['L6e']+lever_list['L7a'],
            'L7b': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f']+lever_list['L4']+lever_list['L6a']+lever_list['L6b']+lever_list['L6c']+\
                    lever_list['L6d']+lever_list['L6e']+lever_list['L7a']+lever_list['L7b'],
            'L7c': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f']+lever_list['L4']+lever_list['L6a']+lever_list['L6b']+lever_list['L6c']+\
                    lever_list['L6d']+lever_list['L6e']+lever_list['L7a']+lever_list['L7b']+lever_list['L7c'],
            'L8a': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f']+lever_list['L4']+lever_list['L6a']+lever_list['L6b']+lever_list['L6c']+\
                    lever_list['L6d']+lever_list['L6e']+lever_list['L7a']+lever_list['L7b']+lever_list['L7c']+lever_list['L8a'],
            'L8b': lever_list['Base']+lever_list['L1a']+lever_list['L1b']+lever_list['L1c']+lever_list['L3']+lever_list['L2']+\
                lever_list['L5a']+lever_list['L5f']+lever_list['L4']+lever_list['L6a']+lever_list['L6b']+lever_list['L6c']+\
                    lever_list['L6d']+lever_list['L6e']+lever_list['L7a']+lever_list['L7b']+lever_list['L7c']+lever_list['L8a']+lever_list['L8b']                    
        }
        data_inputs = scenario_steps_inputs.copy()
        scenario_inputs = scenario_demand_inputs.copy()
    
    print("levers include " + str(lever_list_story.keys()))

    for scenario_index in lever_list_story:
        print(lever_list_story[scenario_index])
        scenarios = lever_list_story[scenario_index]

        if len(scenarios)>0:
            print("Implementing Scenario" + str(scenarios))
            for i in range(0,len(scenarios)):
                data_inputs.iloc[:,scenarios[i]] = scenario_inputs.iloc[:,scenarios[i]]
        else:
            print("Implementing Base scenario")

        ##################new road length####################################
        #unit: km;
        new_km_Mot_Unp = data_inputs.iloc[0:118,0]
        new_km_Mot_Bit = data_inputs.iloc[0:118,1]
        new_km_Mot_Con = data_inputs.iloc[0:118,2]
        new_km_Mot_Com = data_inputs.iloc[0:118,3]

        new_km_Hig_Unp = data_inputs.iloc[0:118,4]
        new_km_Hig_Bit = data_inputs.iloc[0:118,5]
        new_km_Hig_Con = data_inputs.iloc[0:118,6]
        new_km_Hig_Com = data_inputs.iloc[0:118,7]

        new_km_Sec_Unp = data_inputs.iloc[0:118,8]
        new_km_Sec_Bit = data_inputs.iloc[0:118,9]
        new_km_Sec_Con = data_inputs.iloc[0:118,10]
        new_km_Sec_Com = data_inputs.iloc[0:118,11]

        new_km_Oth_Unp = data_inputs.iloc[0:118,12]
        new_km_Oth_Bit = data_inputs.iloc[0:118,13]
        new_km_Oth_Con = data_inputs.iloc[0:118,14]
        new_km_Oth_Com = data_inputs.iloc[0:118,15]

        ##################Scenario parameters 2017-2060#########################################
        #unit: km
        Use_km_Mot = data_inputs.iloc[117:161,16]
        Use_km_Hig = data_inputs.iloc[117:161,17]
        Use_km_Sec = data_inputs.iloc[117:161,18]
        Use_km_Oth = data_inputs.iloc[117:161,19]

        #Motorways
        Mot_Unp = data_inputs.iloc[117:161,20]
        Mot_Bit = data_inputs.iloc[117:161,21]
        Mot_Con = data_inputs.iloc[117:161,22]
        Mot_Com = data_inputs.iloc[117:161,23]

        #Highways
        Hig_Unp = data_inputs.iloc[117:161,24]
        Hig_Bit = data_inputs.iloc[117:161,25]
        Hig_Con = data_inputs.iloc[117:161,26]
        Hig_Com = data_inputs.iloc[117:161,27]

        #Secondary
        Sec_Unp = data_inputs.iloc[117:161,28]
        Sec_Bit = data_inputs.iloc[117:161,29]
        Sec_Con = data_inputs.iloc[117:161,30]
        Sec_Com = data_inputs.iloc[117:161,31]

        #Other roads
        Oth_Unp = data_inputs.iloc[117:161,32]
        Oth_Bit = data_inputs.iloc[117:161,33]
        Oth_Con = data_inputs.iloc[117:161,34]
        Oth_Com = data_inputs.iloc[117:161,35]

        #lifetime: years; cement intensity; kg cement/km
        Life = data_inputs.iloc[:,36]

        int_Con_Mot_Unp = data_inputs.iloc[:,37]
        int_Con_Mot_Bit = data_inputs.iloc[:,38]
        int_Con_Mot_Con = data_inputs.iloc[:,39]
        int_Con_Mot_Com = data_inputs.iloc[:,40]
        int_Con_Hig_Unp = data_inputs.iloc[:,41]
        int_Con_Hig_Bit = data_inputs.iloc[:,42]
        int_Con_Hig_Con = data_inputs.iloc[:,43]
        int_Con_Hig_Com = data_inputs.iloc[:,44]
        int_Con_Sec_Unp = data_inputs.iloc[:,45]
        int_Con_Sec_Bit = data_inputs.iloc[:,46]
        int_Con_Sec_Con = data_inputs.iloc[:,47]
        int_Con_Sec_Com = data_inputs.iloc[:,48]
        int_Con_Oth_Unp = data_inputs.iloc[:,49]
        int_Con_Oth_Bit = data_inputs.iloc[:,50]
        int_Con_Oth_Con = data_inputs.iloc[:,51]
        int_Con_Oth_Com = data_inputs.iloc[:,52]

        cemen_content = data_inputs.iloc[:,53]

        int_Als_matrix = data_inputs.iloc[:,54:70]

        bitumen_content = data_inputs.iloc[:,70]

        int_Gra_matrix = data_inputs.iloc[:,71:87]

        #############################################################################################
        #kg CO2/t clinker; 2018-2060; 43 yrs
        Pro_Em_factor = data_inputs.iloc[118:161,87]

        #MJ/t clinker
        Them_eff = data_inputs.iloc[118:161,88]

        Coal_share_cem = data_inputs.iloc[118:161,89]
        Oil_share_cem  = data_inputs.iloc[118:161,90]
        NaGa_share_cem = data_inputs.iloc[118:161,91]
        WaFu_share_cem = data_inputs.iloc[118:161,92]
        Biom_share_cem = data_inputs.iloc[118:161,93]

        Coal_CO2_cem = data_inputs.iloc[118:161,94]
        Oil_CO2_cem  = data_inputs.iloc[118:161,95]
        NaGa_CO2_cem = data_inputs.iloc[118:161,96]
        WaFu_CO2_cem = data_inputs.iloc[118:161,97]
        Biom_CO2_cem = data_inputs.iloc[118:161,98]

        #clinker is also used in the carbonation module; 1931-2060
        clinker = data_inputs.iloc[:,99]

        Gyp_share = data_inputs.iloc[118:161,100]
        Lim_share = data_inputs.iloc[118:161,101]
        Poz_share = data_inputs.iloc[118:161,102]
        Sla_share = data_inputs.iloc[118:161,103]
        Fly_share = data_inputs.iloc[118:161,104]
        Cla_share = data_inputs.iloc[118:161,105]
        Ene_Cla   = data_inputs.iloc[118:161,106]

        CCS_Oxy_percent = data_inputs.iloc[118:161,107]
        CCS_Pos_percent = data_inputs.iloc[118:161,108]
        Eff_Oxy = data_inputs.iloc[118:161,109]
        Eff_Pos = data_inputs.iloc[118:161,110]
        Ene_Oxy = data_inputs.iloc[118:161,111]
        Ene_Pos = data_inputs.iloc[118:161,112]

        Ele_eff        = data_inputs.iloc[118:161,113]
        Coal_share_ele = data_inputs.iloc[118:161,114]
        Oil_share_ele  = data_inputs.iloc[118:161,115]
        Naga_share_ele = data_inputs.iloc[118:161,116]
        Hydr_share_ele = data_inputs.iloc[118:161,117]
        Geot_share_ele = data_inputs.iloc[118:161,118]
        Sol_share_ele  = data_inputs.iloc[118:161,119]
        Wind_share_ele = data_inputs.iloc[118:161,120]
        Tide_share_ele = data_inputs.iloc[118:161,121]
        Nuc_share_ele  = data_inputs.iloc[118:161,122]
        Bio_share_ele  = data_inputs.iloc[118:161,123]
        Was_share_ele  = data_inputs.iloc[118:161,124]
        SoTh_share_ele = data_inputs.iloc[118:161,125]

        Coal_CO2_ele = data_inputs.iloc[118:161,126]
        Oil_CO2_ele  = data_inputs.iloc[118:161,127]
        Naga_CO2_ele = data_inputs.iloc[118:161,128]
        Hydr_CO2_ele = data_inputs.iloc[118:161,129]
        Geot_CO2_ele = data_inputs.iloc[118:161,130]
        Sol_CO2_ele  = data_inputs.iloc[118:161,131]
        Wind_CO2_ele = data_inputs.iloc[118:161,132]
        Tide_CO2_ele = data_inputs.iloc[118:161,133]
        Nuc_CO2_ele  = data_inputs.iloc[118:161,134]
        Bio_CO2_ele  = data_inputs.iloc[118:161,135]
        Was_CO2_ele  = data_inputs.iloc[118:161,136]
        SoTh_CO2_ele = data_inputs.iloc[118:161,137]

        Belite_share = data_inputs.iloc[:,138]
        BYF_share    = data_inputs.iloc[:,139]
        CCSC_share   = data_inputs.iloc[:,140]
        CSAB_share   = data_inputs.iloc[:,141]
        Celite_share = data_inputs.iloc[:,142]
        MOMS_share   = data_inputs.iloc[:,143]

        Belite_pro_save = data_inputs.iloc[:,144]
        BYF_pro_save    = data_inputs.iloc[:,145]
        CCSC_pro_save   = data_inputs.iloc[:,146]
        CSAB_pro_save   = data_inputs.iloc[:,147]
        Celite_pro_save = data_inputs.iloc[:,148]
        MOMS_pro_save   = data_inputs.iloc[:,149]

        Belite_fue_save = data_inputs.iloc[:,150]
        BYF_fue_save    = data_inputs.iloc[:,151]
        CCSC_fue_save   = data_inputs.iloc[:,152]
        CSAB_fue_save   = data_inputs.iloc[:,153]
        Celite_fue_save = data_inputs.iloc[:,154]
        MOMS_fue_save   = data_inputs.iloc[:,155]

        Belite_alk_save = data_inputs.iloc[:,156]
        BYF_alk_save    = data_inputs.iloc[:,157]
        CCSC_alk_save   = data_inputs.iloc[:,158]
        CSAB_alk_save   = data_inputs.iloc[:,159]
        Celite_alk_save = data_inputs.iloc[:,160]
        MOMS_alk_save   = data_inputs.iloc[:,161]

        CCSC_CO2_credit = data_inputs.iloc[:,162]

        CO2_curing_share = data_inputs.iloc[:,163]
        CO2_curing_coef  = data_inputs.iloc[:,164]
        CO2_curing_pen   = data_inputs.iloc[:,165]
        CO2_curing_red   = data_inputs.iloc[:,166]

        CO2_mine_EOL_share   = data_inputs.iloc[:,167]
        CO2_mine_EOL_coef    = data_inputs.iloc[:,168]
        CO2_mine_EOL_pen     = data_inputs.iloc[:,169]
        CO2_mine_EOL_red     = data_inputs.iloc[:,170]

        CO2_mine_slag_share  = data_inputs.iloc[:,171]
        CO2_mine_slag_upt    = data_inputs.iloc[:,172]
        CO2_mine_slag_pen    = data_inputs.iloc[:,173]

        CO2_mine_ash_share  = data_inputs.iloc[:,174]
        CO2_mine_ash_upt    = data_inputs.iloc[:,175]
        CO2_mine_ash_pen    = data_inputs.iloc[:,176]

        CO2_mine_lime_share  = data_inputs.iloc[:,177]
        CO2_mine_lime_upt    = data_inputs.iloc[:,178]
        CO2_mine_lime_pen    = data_inputs.iloc[:,179]

        CO2_mine_red_share  = data_inputs.iloc[:,180]
        CO2_mine_red_upt    = data_inputs.iloc[:,181]
        CO2_mine_red_pen    = data_inputs.iloc[:,182]

        Leaner_rate = data_inputs.iloc[:,183]
        Leaner_con_reduction = data_inputs.iloc[:,184]

        Mat_sub_rate  = data_inputs.iloc[:,185]
        Timber_to_cem = data_inputs.iloc[:,186]
        CO2_tim_pen   = data_inputs.iloc[:,187]
        CO2_tim_sto   = data_inputs.iloc[:,188]

        EoL = data_inputs.iloc[:,189:194]

        CO2_tran_cem  = data_inputs.iloc[:,194]

        CO2_prod_agg  = data_inputs.iloc[:,195]
        CO2_tran_agg  = data_inputs.iloc[:,196]

        CO2_prod_rec  = data_inputs.iloc[:,197]
        CO2_tran_rec  = data_inputs.iloc[:,198]

        CO2_tran_bur  = data_inputs.iloc[:,199]

        CO2_mix_con   = data_inputs.iloc[:,200]
        CO2_tran_con  = data_inputs.iloc[:,201]

        CO2_onsite_con= data_inputs.iloc[:,202]

        Dem_spread_time = data_inputs.iloc[:,203]
        Dem_spread_ratio= data_inputs.iloc[:,204]

        #assemble data series into dataframe
        store_t = Dem_spread_time*Dem_spread_ratio+\
            (1-Dem_spread_ratio)*0.4

        #assemble data series into dataframe
        new_km_matrix = pd.concat([new_km_Mot_Unp,new_km_Mot_Bit,new_km_Mot_Con,new_km_Mot_Com,\
            new_km_Hig_Unp,new_km_Hig_Bit,new_km_Hig_Con,new_km_Hig_Com,\
                new_km_Sec_Unp,new_km_Sec_Bit,new_km_Sec_Con,new_km_Sec_Com,\
                    new_km_Oth_Unp,new_km_Oth_Bit,new_km_Oth_Con,new_km_Oth_Com], axis = 1)

        CO2_curing_sav = CO2_curing_share*CO2_curing_red
        CO2_curing_emi = CO2_curing_share*CO2_curing_pen

        CO2_mine_EOL_sav = CO2_mine_EOL_share*CO2_mine_EOL_red
        CO2_mine_EOL_emi = CO2_mine_EOL_share*CO2_mine_EOL_pen

        CO2_mine_slag_sto = CO2_mine_slag_share*CO2_mine_slag_upt
        CO2_mine_slag_emi = CO2_mine_slag_share*CO2_mine_slag_pen

        CO2_mine_ash_sto = CO2_mine_ash_share*CO2_mine_ash_upt
        CO2_mine_ash_emi = CO2_mine_ash_share*CO2_mine_ash_pen

        CO2_mine_lime_sto = CO2_mine_lime_share*CO2_mine_lime_upt
        CO2_mine_lime_emi = CO2_mine_lime_share*CO2_mine_lime_pen

        CO2_mine_red_sto = CO2_mine_red_share*CO2_mine_red_upt
        CO2_mine_red_emi = CO2_mine_red_share*CO2_mine_red_pen

        Leaner_con_saving = Leaner_rate*Leaner_con_reduction

        waste_saving = EoL.iloc[:,0]*EoL.iloc[:,1]
        recRate = EoL.iloc[:,2]
        reuseRate = EoL.iloc[:,3]*EoL.iloc[:,4]
        burRate = 1-recRate-reuseRate

        int_Con_matrix = pd.concat([int_Con_Mot_Unp*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Mot_Bit*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Mot_Con*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Mot_Com*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Hig_Unp*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Hig_Bit*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Hig_Con*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Hig_Com*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Sec_Unp*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Sec_Bit*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Sec_Con*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Sec_Com*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Oth_Unp*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Oth_Bit*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Oth_Con*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Oth_Com*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)], axis = 1)

        int_Tim_matrix = pd.concat([int_Con_Mot_Unp*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem,\
                                    int_Con_Mot_Bit*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem,\
                                    int_Con_Mot_Con*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem,\
                                    int_Con_Mot_Com*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem,\
                                    int_Con_Hig_Unp*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem,\
                                    int_Con_Hig_Bit*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem,\
                                    int_Con_Hig_Con*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem,\
                                    int_Con_Hig_Com*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem,\
                                    int_Con_Sec_Unp*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem,\
                                    int_Con_Sec_Bit*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem,\
                                    int_Con_Sec_Con*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem,\
                                    int_Con_Sec_Com*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem,\
                                    int_Con_Oth_Unp*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem,\
                                    int_Con_Oth_Bit*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem,\
                                    int_Con_Oth_Con*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem,\
                                    int_Con_Oth_Com*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem], axis = 1)

        int_Cem_Mot_Unp = int_Con_Mot_Unp*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Mot_Bit = int_Con_Mot_Bit*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Mot_Con = int_Con_Mot_Con*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Mot_Com = int_Con_Mot_Com*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Hig_Unp = int_Con_Hig_Unp*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Hig_Bit = int_Con_Hig_Bit*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Hig_Con = int_Con_Hig_Con*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Hig_Com = int_Con_Hig_Com*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Sec_Unp = int_Con_Sec_Unp*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Sec_Bit = int_Con_Sec_Bit*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Sec_Con = int_Con_Sec_Con*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Sec_Com = int_Con_Sec_Com*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Oth_Unp = int_Con_Oth_Unp*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Oth_Bit = int_Con_Oth_Bit*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Oth_Con = int_Con_Oth_Con*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Oth_Com = int_Con_Oth_Com*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)

        int_Cem_matrix = pd.concat([int_Cem_Mot_Unp,int_Cem_Mot_Bit,int_Cem_Mot_Con,int_Cem_Mot_Com,\
            int_Cem_Hig_Unp,int_Cem_Hig_Bit,int_Cem_Hig_Con,int_Cem_Hig_Com,\
                int_Cem_Sec_Unp,int_Cem_Sec_Bit,int_Cem_Sec_Con,int_Cem_Sec_Com,\
                    int_Cem_Oth_Unp,int_Cem_Oth_Bit,int_Cem_Oth_Con,int_Cem_Oth_Com], axis = 1)

        #Use_km; unit: km
        Use_km_Mot_Unp = Use_km_Mot * Mot_Unp
        Use_km_Mot_Bit = Use_km_Mot * (1-Mot_Unp) * Mot_Bit
        Use_km_Mot_Con = Use_km_Mot * (1-Mot_Unp) * Mot_Con
        Use_km_Mot_Com = Use_km_Mot * (1-Mot_Unp) * Mot_Com

        Use_km_Hig_Unp = Use_km_Hig * Hig_Unp
        Use_km_Hig_Bit = Use_km_Hig * (1-Hig_Unp) * Hig_Bit
        Use_km_Hig_Con = Use_km_Hig * (1-Hig_Unp) * Hig_Con
        Use_km_Hig_Com = Use_km_Hig * (1-Hig_Unp) * Hig_Com

        Use_km_Sec_Unp = Use_km_Sec * Sec_Unp
        Use_km_Sec_Bit = Use_km_Sec * (1-Sec_Unp) * Sec_Bit
        Use_km_Sec_Con = Use_km_Sec * (1-Sec_Unp) * Sec_Con
        Use_km_Sec_Com = Use_km_Sec * (1-Sec_Unp) * Sec_Com

        Use_km_Oth_Unp = Use_km_Oth * Oth_Unp
        Use_km_Oth_Bit = Use_km_Oth * (1-Oth_Unp) * Oth_Bit
        Use_km_Oth_Con = Use_km_Oth * (1-Oth_Unp) * Oth_Con
        Use_km_Oth_Com = Use_km_Oth * (1-Oth_Unp) * Oth_Com

        Use_km_matrix = pd.concat([
            Use_km_Mot_Unp,Use_km_Mot_Bit,Use_km_Mot_Con,Use_km_Mot_Com,\
                Use_km_Hig_Unp,Use_km_Hig_Bit,Use_km_Hig_Con,Use_km_Hig_Com,\
                    Use_km_Sec_Unp,Use_km_Sec_Bit,Use_km_Sec_Con,Use_km_Sec_Com,\
                        Use_km_Oth_Unp,Use_km_Oth_Bit,Use_km_Oth_Con,Use_km_Oth_Com
                    ], axis = 1)

        ##############################################################################
        #stock-driven model#
        ##############################################################################
        new_km_extended_matrix = pd.DataFrame()
        con_app_matrix = pd.DataFrame()
        cem_app_matrix = pd.DataFrame()
        tim_con_matrix = pd.DataFrame()
        dem_con_matrix = pd.DataFrame()
        dem_tim_matrix = pd.DataFrame()

        for i in range(0, len(new_km_matrix.columns)):
            new_km_extended = new_km_matrix.iloc[:,i]
            year_complete = np.arange(1900,2018)
            demol_Con_extended = np.repeat(0,len(year_complete))
            reuse_Con_extended = np.repeat(0,len(year_complete))
            reuse_Cem_extended = np.repeat(0,len(year_complete))
            demol_Tim_extended = np.repeat(0,len(year_complete))
            
            for k in range(2018, 2061):
                #!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                Life_Road  = Life[0:len(year_complete)] 
                int_Con = int_Con_matrix.iloc[0:len(year_complete),i]
                int_Cem = int_Cem_matrix.iloc[0:len(year_complete),i]
                int_Tim = int_Tim_matrix.iloc[0:len(year_complete),i]

                reuse_rate = reuseRate.iloc[0:len(year_complete)]
                rec_rate   = recRate[0:len(year_complete)]

                demolish_km_list = new_km_extended * (norm.cdf(k-year_complete,Life_Road,Life_Road*0.2)-\
                    norm.cdf(k-1-year_complete,Life_Road,Life_Road*0.2))
                if k-1900 <= Life_Road[0]:
                    demolish_km_list[0] = new_km_extended[0] * 1/Life_Road[0]
                else:
                    demolish_km_list[0] = 0
                demolish_km = sum(demolish_km_list)

                demolish_Con_list = demolish_km_list * int_Con
                demolish_Con      = sum(demolish_Con_list)
                demol_Con_extended= np.append(demol_Con_extended,demolish_Con)

                demolish_Tim_list = demolish_km_list * int_Tim
                demolish_Tim      = sum(demolish_Tim_list)
                demol_Tim_extended= np.append(demol_Tim_extended,demolish_Tim)
                
                new_km = Use_km_matrix.iloc[k-2017,i] - Use_km_matrix.iloc[k-2018,i] + demolish_km
                if new_km < 0:
                    new_km = 0
                new_km_extended = np.append(new_km_extended, new_km)

                reuse_Con_list    = demolish_km_list * int_Con
                reuse_Cem_list    = demolish_km_list * int_Cem * reuse_rate
                reuse_Con         = sum(reuse_Con_list)
                reuse_Cem         = sum(reuse_Cem_list)
                if new_km == 0:
                    reuse_Con = 0
                    reuse_Cem = 0
                reuse_Con_extended= np.append(reuse_Con_extended,reuse_Cem)
                reuse_Cem_extended= np.append(reuse_Cem_extended,reuse_Cem)

                year_complete = np.append(year_complete,k)

            new_km_extended_matrix = pd.concat([new_km_extended_matrix, pd.Series(new_km_extended)],\
                axis = 1, ignore_index = True)
            #unit: t cement
            reuse_rate   = reuseRate.iloc[0:len(year_complete)]
            print(reuse_rate)

            con_app = new_km_extended * int_Con_matrix.iloc[:,i] - reuse_Con_extended
            con_app_matrix = pd.concat([con_app_matrix, pd.Series(con_app)], axis = 1, ignore_index = True)

            cem_app = new_km_extended * int_Cem_matrix.iloc[:,i] - reuse_Cem_extended
            cem_app_matrix = pd.concat([cem_app_matrix, pd.Series(cem_app)], axis = 1, ignore_index = True)

            tim_con = new_km_extended * int_Tim_matrix.iloc[:,i]
            tim_con_matrix = pd.concat([tim_con_matrix, pd.Series(tim_con)], axis = 1, ignore_index = True)

            dem_con = demol_Con_extended
            dem_con_matrix = pd.concat([dem_con_matrix, pd.Series(dem_con)], axis = 1, ignore_index = True)

            dem_tim = demol_Tim_extended
            dem_tim_matrix = pd.concat([dem_tim_matrix, pd.Series(dem_tim)], axis = 1, ignore_index = True)

        Concrete_demand = np.sum(con_app_matrix, axis=1)
        Cement_demand = np.sum(cem_app_matrix, axis=1)
        Timber_demand = np.sum(tim_con_matrix, axis=1)
        Concrete_demolish = np.sum(dem_con_matrix, axis=1)
        Timber_demolish = np.sum(dem_tim_matrix, axis=1)

        pd.DataFrame(new_km_extended_matrix).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\USA-Roads' + \
                    r'\New_km_list_' + str(scenario_index+"_"+storyline) + '.xlsx')

        pd.DataFrame(Concrete_demand).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\USA-Roads' + \
                    r'\Concrete_demand_' + str(scenario_index+"_"+storyline) + '.xlsx')

        pd.DataFrame(Cement_demand).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\USA-Roads' + \
                    r'\Cement_demand_' + str(scenario_index+"_"+storyline) + '.xlsx')

        pd.DataFrame(Timber_demand).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\USA-Roads' + \
                    r'\Timber_demand_' + str(scenario_index+"_"+storyline) + '.xlsx')

        pd.DataFrame(Concrete_demolish).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\USA-Roads' + \
                    r'\Concrete_demolish_' + str(scenario_index+"_"+storyline) + '.xlsx')

        pd.DataFrame(Timber_demolish).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\USA-Roads' + \
                    r'\Timber_demolish_' + str(scenario_index+"_"+storyline) + '.xlsx')  

        import Emissions

        Emissions.emissions_calculate(
        year_len = 130,
        index_year2018 = 118,
        index_year2060 = 161,
        Belite_share = Belite_share,
        Belite_fue_save = Belite_fue_save,
        Belite_pro_save = Belite_pro_save,
        BYF_share = BYF_share,
        BYF_fue_save = BYF_fue_save,
        BYF_pro_save = BYF_pro_save,
        CCSC_share = CCSC_share,
        CCSC_fue_save = CCSC_fue_save,
        CCSC_pro_save = CCSC_pro_save,
        CCSC_CO2_credit = CCSC_CO2_credit,
        CSAB_share = CSAB_share,
        CSAB_fue_save = CSAB_fue_save,
        CSAB_pro_save = CSAB_pro_save,
        Celite_share = Celite_share,
        Celite_fue_save = Celite_fue_save,
        Celite_pro_save = Celite_pro_save,
        MOMS_share = MOMS_share,
        MOMS_fue_save = MOMS_fue_save,
        MOMS_pro_save = MOMS_pro_save,
        Pro_Em_factor = Pro_Em_factor,
        Them_eff = Them_eff,
        Coal_share_cem = Coal_share_cem,
        Coal_CO2_cem = Coal_CO2_cem,
        Oil_share_cem = Oil_share_cem,
        Oil_CO2_cem = Oil_CO2_cem,
        NaGa_share_cem = NaGa_share_cem,
        NaGa_CO2_cem = NaGa_CO2_cem,
        WaFu_share_cem = WaFu_share_cem,
        WaFu_CO2_cem = WaFu_CO2_cem,
        Biom_share_cem = Biom_share_cem,
        Biom_CO2_cem = Biom_CO2_cem,
        Ele_eff = Ele_eff,
        Coal_share_ele = Coal_share_ele,
        Coal_CO2_ele = Coal_CO2_ele,
        Oil_share_ele = Oil_share_ele,
        Oil_CO2_ele = Oil_CO2_ele,
        Naga_share_ele = Naga_share_ele,
        Naga_CO2_ele = Naga_CO2_ele,
        Hydr_share_ele = Hydr_share_ele,
        Hydr_CO2_ele = Hydr_CO2_ele,
        Geot_share_ele = Geot_share_ele,
        Geot_CO2_ele = Geot_CO2_ele,
        Sol_share_ele = Sol_share_ele,
        Sol_CO2_ele = Sol_CO2_ele,
        Wind_share_ele = Wind_share_ele,
        Wind_CO2_ele = Wind_CO2_ele,
        Tide_share_ele = Tide_share_ele,
        Tide_CO2_ele = Tide_CO2_ele,
        Nuc_share_ele = Nuc_share_ele,
        Nuc_CO2_ele = Nuc_CO2_ele,
        Bio_share_ele = Bio_share_ele,
        Bio_CO2_ele = Bio_CO2_ele,
        Was_share_ele = Was_share_ele,
        Was_CO2_ele = Was_CO2_ele,
        SoTh_share_ele = SoTh_share_ele,
        SoTh_CO2_ele = SoTh_CO2_ele,
        clinker = clinker,
        Cla_share = Cla_share,
        Ene_Cla = Ene_Cla,
        CCS_Oxy_percent = CCS_Oxy_percent,
        Eff_Oxy = Eff_Oxy,
        Ene_Oxy = Ene_Oxy,
        CCS_Pos_percent = CCS_Pos_percent,
        Eff_Pos = Eff_Pos,
        Ene_Pos = Ene_Pos,
        Cement_demand = Cement_demand,
        CO2_tran_cem = CO2_tran_cem,
        Concrete_demand = Concrete_demand,
        CO2_prod_agg = CO2_prod_agg,
        CO2_tran_agg = CO2_tran_agg,
        Concrete_demolish = Concrete_demolish,
        recRate = recRate,
        CO2_prod_rec = CO2_prod_rec,
        CO2_tran_rec = CO2_tran_rec,
        burRate = burRate,
        CO2_tran_bur = CO2_tran_bur,
        CO2_mix_con = CO2_mix_con,
        CO2_tran_con = CO2_tran_con,
        CO2_onsite_con = CO2_onsite_con,
        CO2_curing_emi = CO2_curing_emi,
        CO2_mine_EOL_emi = CO2_mine_EOL_emi,
        CO2_mine_slag_emi = CO2_mine_slag_emi,
        CO2_mine_slag_sto = CO2_mine_slag_sto,
        CO2_mine_ash_emi = CO2_mine_ash_emi,
        CO2_mine_ash_sto = CO2_mine_ash_sto,
        CO2_mine_lime_emi = CO2_mine_lime_emi,
        CO2_mine_lime_sto = CO2_mine_lime_sto,
        CO2_mine_red_emi = CO2_mine_red_emi,
        CO2_mine_red_sto = CO2_mine_red_sto,
        CO2_mine_slag_share = CO2_mine_slag_share,
        CO2_mine_ash_share = CO2_mine_ash_share,
        CO2_mine_lime_share = CO2_mine_lime_share,
        CO2_mine_red_share = CO2_mine_red_share,        
        Timber_demand = Timber_demand,
        CO2_tim_pen = CO2_tim_pen,
        CO2_tim_sto = CO2_tim_sto,
        Timber_demolish = Timber_demolish,
        EoL_tim_land = pd.Series(np.zeros(161)),
        Land_tim_deg = pd.Series(np.zeros(161)),
        Deg_tim_CH4 = pd.Series(np.zeros(161)),
        Deg_tim_CO2 = pd.Series(np.zeros(161)),
        CH4_to_CO2 = pd.Series(np.zeros(161)),
        CH4_recover = pd.Series(np.zeros(161)),
        EoL_tim_inci = pd.Series(np.zeros(161)),
        scenario_index = scenario_index,
        storyline = storyline,
        filepath = 'USA-Roads'
        )

        import Carbonation_USA
        #index_year2018 as in Cement_demand[31:161]
        Carbonation_USA.cement_carbonation(
        year_len = 130,
        index_year2018 = 87,
        Belite_share = (Belite_share[31:161]).reset_index(drop=True),
        Belite_alk_save = (Belite_alk_save[31:161]).reset_index(drop=True),
        BYF_share = (BYF_share[31:161]).reset_index(drop=True),
        BYF_alk_save = (BYF_alk_save[31:161]).reset_index(drop=True),
        CCSC_share = (CCSC_share[31:161]).reset_index(drop=True),
        CCSC_alk_save = (CCSC_alk_save[31:161]).reset_index(drop=True),
        CSAB_share = (CSAB_share[31:161]).reset_index(drop=True),
        CSAB_alk_save = (CSAB_alk_save[31:161]).reset_index(drop=True),
        Celite_share = (Celite_share[31:161]).reset_index(drop=True),
        Celite_alk_save = (Celite_alk_save[31:161]).reset_index(drop=True),
        MOMS_share = (MOMS_share[31:161]).reset_index(drop=True),
        MOMS_alk_save = (MOMS_alk_save[31:161]).reset_index(drop=True),
        CO2_curing_share = (CO2_curing_share[31:161]).reset_index(drop=True),
        CO2_curing_coef = (CO2_curing_coef[31:161]).reset_index(drop=True),
        CO2_mine_EOL_share = (CO2_mine_EOL_share[31:161]).reset_index(drop=True),
        CO2_mine_EOL_coef = (CO2_mine_EOL_coef[31:161]).reset_index(drop=True),
        clinker = (clinker[31:161]).reset_index(drop=True),
        Cement_demand = (Cement_demand[31:161]).reset_index(drop=True),
        waste_saving = (waste_saving[31:161]).reset_index(drop=True),
        scenario_index = scenario_index,
        storyline = storyline,
        Life = (Life[31:161]).reset_index(drop=True),
        reuse_rate = (reuseRate[31:161]).reset_index(drop=True),
        recRate = (recRate[31:161]).reset_index(drop=True),
        burRate = (burRate[31:161]).reset_index(drop=True),
        store_t = (store_t[31:161]).reset_index(drop=True),
        filepath = 'USA-Roads'
        )
