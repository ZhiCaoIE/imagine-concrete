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
header = None, sheet_name="India-Buildings", usecols= "B:FH", nrows = 130, skiprows = 4)

scenario_steps_inputs = pd.read_excel (io = r'C:\Users\caozh\Box\ERSAL Cement 2018-2019\Zhi Cao\Raw data\Data inputs.xlsx', 
header = None, sheet_name="India-Buildings-STEPS", usecols= "B:FH", nrows = 130, skiprows = 4)

scenario_production_inputs = pd.read_excel (io = r'C:\Users\caozh\Box\ERSAL Cement 2018-2019\Zhi Cao\Raw data\Data inputs.xlsx', 
header = None, sheet_name="India-Buildings-Production", usecols= "B:FH", nrows = 130, skiprows = 4)

scenario_demand_inputs = pd.read_excel (io = r'C:\Users\caozh\Box\ERSAL Cement 2018-2019\Zhi Cao\Raw data\Data inputs.xlsx', 
header = None, sheet_name="India-Buildings-Demand", usecols= "B:FH", nrows = 130, skiprows = 4)

lever_list = {
     'Base': [],
     'L1a': [36],
     'L1b': [61],
     'L1c': [37,38,39,40,41],
     'L3':  [47,48,49,50,51,52,53],
     'L2':  [86,87,88,89,90,91],
     'L5a': [111],
     'L5f': [115,119,122,125,128],
     'L4':  [55,56,57,58],
     'L6a': [131],
     'L6b': [133],
     'L6c': [147],
     'L6d': [10,11],
     'L6e': [23,24],
     'L7a': [149],
     'L7b': [150],
     'L7c': [162],
     'L8a': [62,63,64,65,66,67,68,69,70,71,72,73],
     'L8b': [152,153,154,155,156,157,158,159,160]
    }

storylines = ['STEPS','Production','Demand']
#data_inputs = scenario_frozen_inputs
#storylines = ['Production','Demand']

for storyline in storylines:
    #storyline = 'Demand'
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

        ##################new residential Urban floor area####################################
        #unit: m2; TO: Timber-others
        new_m2_Res_Urban_TO = data_inputs.iloc[0:87,0]

        #unit: m2; BT: Brick-timber
        new_m2_Res_Urban_BT = data_inputs.iloc[0:87,1]

        #unit: m2; CB: Concrete-brick
        new_m2_Res_Urban_CB = data_inputs.iloc[0:87,2]

        ##################new residential Rural floor area####################################
        #unit: m2; TO: Timber-others
        new_m2_Res_Rural_TO = data_inputs.iloc[0:87,3]

        #unit: m2; BT: Brick-timber
        new_m2_Res_Rural_BT = data_inputs.iloc[0:87,4]

        #unit: m2; CB: Concrete-brick
        new_m2_Res_Rural_CB = data_inputs.iloc[0:87,5]

        ##################new nonresidential floor area########################################
        #unit: m2; TO: Concrete-steel
        new_m2_NonR_TO = data_inputs.iloc[0:87,6]

        #unit: m2; BT: Concrete-brick
        new_m2_NonR_BT = data_inputs.iloc[0:87,7]

        #unit: m2; CB: Brick-timber
        new_m2_NonR_CB = data_inputs.iloc[0:87,8]

        ##################Scenario parameters 2017-2060#########################################
        #unit: 1,000
        pop = data_inputs.iloc[86:130,9]

        #unit: m2/cap
        m2_per_Res = data_inputs.iloc[86:130,10]

        #unit: m2; in-use residential floor area
        Use_m2_Res = pop*m2_per_Res*1000

        #unit: million m2; in-use non-residential floor area
        Use_m2_NonR = data_inputs.iloc[86:130,11]

        #unit: m2
        Use_m2_NonR = Use_m2_NonR*1000000

        #Res-Urban and Res-Rural
        Res_Urban = data_inputs.iloc[86:130,12]
        Res_Rural = data_inputs.iloc[86:130,13]

        #Res_Urban_TO; Res_Urban_BT; Res_Urban_CB
        Res_Urban_TO = data_inputs.iloc[86:130,14]
        Res_Urban_BT = data_inputs.iloc[86:130,15]
        Res_Urban_CB = data_inputs.iloc[86:130,16]

        #Res_Rural_TO; Res_Rural_BT; Res_Rural_CB
        Res_Rural_TO = data_inputs.iloc[86:130,17]
        Res_Rural_BT = data_inputs.iloc[86:130,18]
        Res_Rural_CB = data_inputs.iloc[86:130,19]

        #NonR_TO; NonR_BT; NonR_CB
        NonR_TO = data_inputs.iloc[86:130,20]
        NonR_BT = data_inputs.iloc[86:130,21]
        NonR_CB = data_inputs.iloc[86:130,22]

        #lifetime: years; cement intensity: kg cement/m2
        Life = data_inputs.iloc[:,23:25]

        int_Con_Res_Urban_TO = data_inputs.iloc[:,25]
        int_Con_Res_Urban_BT = data_inputs.iloc[:,26]
        int_Con_Res_Urban_CB = data_inputs.iloc[:,27]
        int_Con_Res_Rural_TO = data_inputs.iloc[:,28]
        int_Con_Res_Rural_BT = data_inputs.iloc[:,29]
        int_Con_Res_Rural_CB = data_inputs.iloc[:,30]
        int_Con_NonR_TO = data_inputs.iloc[:,31]
        int_Con_NonR_BT = data_inputs.iloc[:,32]
        int_Con_NonR_CB = data_inputs.iloc[:,33]

        cemen_content = data_inputs.iloc[:,34]

        #############################################################################################
        #kg CO2/t clinker; 2018-2060; 43 yrs
        Pro_Em_factor = data_inputs.iloc[87:130,35]

        #MJ/t clinker
        Them_eff = data_inputs.iloc[87:130,36]

        Coal_share_cem = data_inputs.iloc[87:130,37]
        Oil_share_cem  = data_inputs.iloc[87:130,38]
        NaGa_share_cem = data_inputs.iloc[87:130,39]
        WaFu_share_cem = data_inputs.iloc[87:130,40]
        Biom_share_cem = data_inputs.iloc[87:130,41]

        Coal_CO2_cem = data_inputs.iloc[87:130,42]
        Oil_CO2_cem  = data_inputs.iloc[87:130,43]
        NaGa_CO2_cem = data_inputs.iloc[87:130,44]
        WaFu_CO2_cem = data_inputs.iloc[87:130,45]
        Biom_CO2_cem = data_inputs.iloc[87:130,46]

        #clinker is also used in the carbonation module; 1931-2060
        clinker = data_inputs.iloc[:,47]

        Gyp_share = data_inputs.iloc[87:130,48]
        Lim_share = data_inputs.iloc[87:130,49]
        Poz_share = data_inputs.iloc[87:130,50]
        Sla_share = data_inputs.iloc[87:130,51]
        Fly_share = data_inputs.iloc[87:130,52]
        Cla_share = data_inputs.iloc[87:130,53]
        Ene_Cla   = data_inputs.iloc[87:130,54]

        CCS_Oxy_percent = data_inputs.iloc[87:130,55]
        CCS_Pos_percent = data_inputs.iloc[87:130,56]
        Eff_Oxy = data_inputs.iloc[87:130,57]
        Eff_Pos = data_inputs.iloc[87:130,58]
        Ene_Oxy = data_inputs.iloc[87:130,59]
        Ene_Pos = data_inputs.iloc[87:130,60]

        Ele_eff        = data_inputs.iloc[87:130,61]
        Coal_share_ele = data_inputs.iloc[87:130,62]
        Oil_share_ele  = data_inputs.iloc[87:130,63]
        Naga_share_ele = data_inputs.iloc[87:130,64]
        Hydr_share_ele = data_inputs.iloc[87:130,65]
        Geot_share_ele = data_inputs.iloc[87:130,66]
        Sol_share_ele  = data_inputs.iloc[87:130,67]
        Wind_share_ele = data_inputs.iloc[87:130,68]
        Tide_share_ele = data_inputs.iloc[87:130,69]
        Nuc_share_ele  = data_inputs.iloc[87:130,70]
        Bio_share_ele  = data_inputs.iloc[87:130,71]
        Was_share_ele  = data_inputs.iloc[87:130,72]
        SoTh_share_ele = data_inputs.iloc[87:130,73]

        Coal_CO2_ele = data_inputs.iloc[87:130,74]
        Oil_CO2_ele  = data_inputs.iloc[87:130,75]
        Naga_CO2_ele = data_inputs.iloc[87:130,76]
        Hydr_CO2_ele = data_inputs.iloc[87:130,77]
        Geot_CO2_ele = data_inputs.iloc[87:130,78]
        Sol_CO2_ele  = data_inputs.iloc[87:130,79]
        Wind_CO2_ele = data_inputs.iloc[87:130,80]
        Tide_CO2_ele = data_inputs.iloc[87:130,81]
        Nuc_CO2_ele  = data_inputs.iloc[87:130,82]
        Bio_CO2_ele  = data_inputs.iloc[87:130,83]
        Was_CO2_ele  = data_inputs.iloc[87:130,84]
        SoTh_CO2_ele = data_inputs.iloc[87:130,85]

        Belite_share = data_inputs.iloc[:,86]
        BYF_share    = data_inputs.iloc[:,87]
        CCSC_share   = data_inputs.iloc[:,88]
        CSAB_share   = data_inputs.iloc[:,89]
        Celite_share = data_inputs.iloc[:,90]
        MOMS_share   = data_inputs.iloc[:,91]

        Belite_pro_save = data_inputs.iloc[:,92]
        BYF_pro_save    = data_inputs.iloc[:,93]
        CCSC_pro_save   = data_inputs.iloc[:,94]
        CSAB_pro_save   = data_inputs.iloc[:,95]
        Celite_pro_save = data_inputs.iloc[:,96]
        MOMS_pro_save   = data_inputs.iloc[:,97]

        Belite_fue_save = data_inputs.iloc[:,98]
        BYF_fue_save    = data_inputs.iloc[:,99]
        CCSC_fue_save   = data_inputs.iloc[:,100]
        CSAB_fue_save   = data_inputs.iloc[:,101]
        Celite_fue_save = data_inputs.iloc[:,102]
        MOMS_fue_save   = data_inputs.iloc[:,103]

        Belite_alk_save = data_inputs.iloc[:,104]
        BYF_alk_save    = data_inputs.iloc[:,105]
        CCSC_alk_save   = data_inputs.iloc[:,106]
        CSAB_alk_save   = data_inputs.iloc[:,107]
        Celite_alk_save = data_inputs.iloc[:,108]
        MOMS_alk_save   = data_inputs.iloc[:,109]

        CCSC_CO2_credit = data_inputs.iloc[:,110]

        CO2_curing_share = data_inputs.iloc[:,111]
        CO2_curing_coef  = data_inputs.iloc[:,112]
        CO2_curing_pen   = data_inputs.iloc[:,113]
        CO2_curing_red   = data_inputs.iloc[:,114]

        CO2_mine_EOL_share   = data_inputs.iloc[:,115]
        CO2_mine_EOL_coef    = data_inputs.iloc[:,116]
        CO2_mine_EOL_pen     = data_inputs.iloc[:,117]
        CO2_mine_EOL_red     = data_inputs.iloc[:,118]

        CO2_mine_slag_share  = data_inputs.iloc[:,119]
        CO2_mine_slag_upt    = data_inputs.iloc[:,120]
        CO2_mine_slag_pen    = data_inputs.iloc[:,121]

        CO2_mine_ash_share  = data_inputs.iloc[:,122]
        CO2_mine_ash_upt    = data_inputs.iloc[:,123]
        CO2_mine_ash_pen    = data_inputs.iloc[:,124]

        CO2_mine_lime_share  = data_inputs.iloc[:,125]
        CO2_mine_lime_upt    = data_inputs.iloc[:,126]
        CO2_mine_lime_pen    = data_inputs.iloc[:,127]

        CO2_mine_red_share  = data_inputs.iloc[:,128]
        CO2_mine_red_upt    = data_inputs.iloc[:,129]
        CO2_mine_red_pen    = data_inputs.iloc[:,130]

        Leaner_rate = data_inputs.iloc[:,131]
        Leaner_con_reduction = data_inputs.iloc[:,132]

        Mat_sub_rate  = data_inputs.iloc[:,133]
        Timber_to_cem_res = data_inputs.iloc[:,134]
        Timber_to_cem_nonr = data_inputs.iloc[:,135]
        CO2_tim_pen   = data_inputs.iloc[:,136]
        CO2_tim_sto   = data_inputs.iloc[:,137]
        EoL_tim_land  = data_inputs.iloc[:,138]
        EoL_tim_inci  = data_inputs.iloc[:,139]
        EoL_tim_recy  = data_inputs.iloc[:,140]
        Land_tim_deg  = data_inputs.iloc[:,141]
        Land_tim_per  = data_inputs.iloc[:,142]
        Deg_tim_CO2   = data_inputs.iloc[:,143]
        Deg_tim_CH4   = data_inputs.iloc[:,144]
        CH4_to_CO2    = data_inputs.iloc[:,145]
        CH4_recover   = data_inputs.iloc[:,146]    

        EoL = data_inputs.iloc[:,147:152]

        CO2_tran_cem  = data_inputs.iloc[:,152]

        CO2_prod_agg  = data_inputs.iloc[:,153]
        CO2_tran_agg  = data_inputs.iloc[:,154]

        CO2_prod_rec  = data_inputs.iloc[:,155]
        CO2_tran_rec  = data_inputs.iloc[:,156]

        CO2_tran_bur  = data_inputs.iloc[:,157]

        CO2_mix_con   = data_inputs.iloc[:,158]
        CO2_tran_con  = data_inputs.iloc[:,159]

        CO2_onsite_con= data_inputs.iloc[:,160]

        Dem_spread_time = data_inputs.iloc[:,161]
        Dem_spread_ratio= data_inputs.iloc[:,162]

        #assemble data series into dataframe
        store_t = Dem_spread_time*Dem_spread_ratio+\
            (1-Dem_spread_ratio)*0.4

        #assemble data series into dataframe
        new_m2_matrix = pd.concat([new_m2_Res_Urban_TO,new_m2_Res_Urban_BT,new_m2_Res_Urban_CB,\
            new_m2_Res_Rural_TO,new_m2_Res_Rural_BT,new_m2_Res_Rural_CB,\
                new_m2_NonR_TO,new_m2_NonR_BT,new_m2_NonR_CB], axis = 1)

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

        int_Con_matrix = pd.concat([int_Con_Res_Urban_TO*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015),\
                                    int_Con_Res_Urban_BT*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015),\
                                    int_Con_Res_Urban_CB*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Res_Rural_TO*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015),\
                                    int_Con_Res_Rural_BT*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015),\
                                    int_Con_Res_Rural_CB*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_NonR_TO*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015),\
                                    int_Con_NonR_BT*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015),\
                                    int_Con_NonR_CB*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)], axis = 1)

        int_Tim_matrix = pd.concat([int_Con_Res_Urban_TO*(1-Leaner_con_saving)*0,\
                                    int_Con_Res_Urban_BT*(1-Leaner_con_saving)*0,\
                                    int_Con_Res_Urban_CB*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_Res_Rural_TO*(1-Leaner_con_saving)*0,\
                                    int_Con_Res_Rural_BT*(1-Leaner_con_saving)*0,\
                                    int_Con_Res_Rural_CB*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_NonR_TO*(1-Leaner_con_saving)*0,\
                                    int_Con_NonR_BT*(1-Leaner_con_saving)*0,\
                                    int_Con_NonR_CB*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_nonr], axis = 1)

        int_Cem_Res_Urban_TO = int_Con_Res_Urban_TO*cemen_content*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_Urban_BT = int_Con_Res_Urban_BT*cemen_content*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_Urban_CB = int_Con_Res_Urban_CB*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_Rural_TO = int_Con_Res_Rural_TO*cemen_content*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_Rural_BT = int_Con_Res_Rural_BT*cemen_content*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_Rural_CB = int_Con_Res_Rural_CB*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_NonR_TO = int_Con_NonR_TO * cemen_content*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_NonR_BT = int_Con_NonR_BT * cemen_content*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_NonR_CB = int_Con_NonR_CB * cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)

        int_Cem_matrix = pd.concat([
            int_Cem_Res_Urban_TO,int_Cem_Res_Urban_BT,int_Cem_Res_Urban_CB,\
                int_Cem_Res_Rural_TO,int_Cem_Res_Rural_BT,int_Cem_Res_Rural_CB,\
                    int_Cem_NonR_TO,int_Cem_NonR_BT,int_Cem_NonR_CB], axis = 1)

        #Use_m2; unit: m2
        Use_m2_Res_Urban_TO = Use_m2_Res * Res_Urban * Res_Urban_TO
        Use_m2_Res_Urban_BT = Use_m2_Res * Res_Urban * Res_Urban_BT
        Use_m2_Res_Urban_CB = Use_m2_Res * Res_Urban * Res_Urban_CB

        Use_m2_Res_Rural_TO = Use_m2_Res * Res_Rural * Res_Rural_TO
        Use_m2_Res_Rural_BT = Use_m2_Res * Res_Rural * Res_Rural_BT
        Use_m2_Res_Rural_CB = Use_m2_Res * Res_Rural * Res_Rural_CB

        Use_m2_NonR_TO = Use_m2_NonR * NonR_TO
        Use_m2_NonR_BT = Use_m2_NonR * NonR_BT
        Use_m2_NonR_CB = Use_m2_NonR * NonR_CB

        Use_m2_matrix = pd.concat([Use_m2_Res_Urban_TO,Use_m2_Res_Urban_BT,Use_m2_Res_Urban_CB,\
            Use_m2_Res_Rural_TO,Use_m2_Res_Rural_BT,Use_m2_Res_Rural_CB,\
                Use_m2_NonR_TO,Use_m2_NonR_BT,Use_m2_NonR_CB], axis = 1)

        ##############################################################################
        #stock-driven model#
        ##############################################################################
        new_m2_extended_matrix = pd.DataFrame()
        con_app_matrix = pd.DataFrame()
        cem_app_matrix = pd.DataFrame()
        tim_con_matrix = pd.DataFrame()
        dem_con_matrix = pd.DataFrame()
        dem_tim_matrix = pd.DataFrame()

        for i in range(0, len(new_m2_matrix.columns)):
            new_m2_extended = new_m2_matrix.iloc[:,i]
            year_complete = np.arange(1931,2018)
            demol_Con_extended = np.repeat(0,len(year_complete))
            reuse_Con_extended = np.repeat(0,len(year_complete))
            reuse_Cem_extended = np.repeat(0,len(year_complete))
            demol_Tim_extended = np.repeat(0,len(year_complete))

            for k in range(2018, 2061):
                Life_Res  = Life.iloc[0:len(year_complete),0]
                Life_NonR = Life.iloc[0:len(year_complete),1]
                int_Con = int_Con_matrix.iloc[0:len(year_complete),i]
                int_Cem = int_Cem_matrix.iloc[0:len(year_complete),i]
                int_Tim = int_Tim_matrix.iloc[0:len(year_complete),i]
                
                reuse_rate = reuseRate.iloc[0:len(year_complete)]
                rec_rate   = recRate[0:len(year_complete)]

                demolish_m2_list = new_m2_extended * (norm.cdf(k-year_complete,Life_Res,Life_Res*0.2)-\
                    norm.cdf(k-1-year_complete,Life_Res,Life_Res*0.2))
                demolish_m2 = sum(demolish_m2_list)

                demolish_Con_list = demolish_m2_list * int_Con / 1000
                demolish_Con      = sum(demolish_Con_list)
                demol_Con_extended= np.append(demol_Con_extended,demolish_Con)

                demolish_Tim_list = demolish_m2_list * int_Tim / 1000
                demolish_Tim      = sum(demolish_Tim_list)
                demol_Tim_extended= np.append(demol_Tim_extended,demolish_Tim)

                new_m2 = Use_m2_matrix.iloc[k-2017,i] - Use_m2_matrix.iloc[k-2018,i] + demolish_m2
                if new_m2 < 0:
                    new_m2 = 0
                new_m2_extended = np.append(new_m2_extended, new_m2)
                
                reuse_Con_list    = demolish_m2_list * int_Con / 1000 * reuse_rate
                reuse_Cem_list    = demolish_m2_list * int_Cem / 1000 * reuse_rate
                reuse_Con         = sum(reuse_Con_list)
                reuse_Cem         = sum(reuse_Cem_list)
                if new_m2 == 0:
                    reuse_Con = 0
                    reuse_Cem = 0
                reuse_Con_extended= np.append(reuse_Con_extended,reuse_Con)
                reuse_Cem_extended= np.append(reuse_Cem_extended,reuse_Cem)

                year_complete = np.append(year_complete,k)

            new_m2_extended_matrix = pd.concat([new_m2_extended_matrix, pd.Series(new_m2_extended)],\
                axis = 1, ignore_index = True)
            #unit: t cement
            reuse_rate   = reuseRate.iloc[0:len(year_complete)]
            print(reuse_rate) 

            con_app = new_m2_extended * int_Con_matrix.iloc[:,i] /1000 - reuse_Con_extended
            con_app_matrix = pd.concat([con_app_matrix, pd.Series(con_app)], axis = 1, ignore_index = True)

            cem_app = new_m2_extended * int_Cem_matrix.iloc[:,i] /1000 - reuse_Cem_extended
            cem_app_matrix = pd.concat([cem_app_matrix, pd.Series(cem_app)], axis = 1, ignore_index = True)

            tim_con = new_m2_extended * int_Tim_matrix.iloc[:,i] /1000
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

        pd.DataFrame(new_m2_extended_matrix).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\India-Buildings' + \
                    r'\New_m2_list_' + str(scenario_index+"_"+storyline) + '.xlsx')

        pd.DataFrame(Concrete_demand).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\India-Buildings' + \
                    r'\Concrete_demand_' + str(scenario_index+"_"+storyline) + '.xlsx')

        pd.DataFrame(Cement_demand).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\India-Buildings' + \
                    r'\Cement_demand_' + str(scenario_index+"_"+storyline) + '.xlsx')

        pd.DataFrame(Timber_demand).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\India-Buildings' + \
                    r'\Timber_demand_' + str(scenario_index+"_"+storyline) + '.xlsx')

        pd.DataFrame(Concrete_demolish).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\India-Buildings' + \
                    r'\Concrete_demolish_' + str(scenario_index+"_"+storyline) + '.xlsx')

        pd.DataFrame(Timber_demolish).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\India-Buildings' + \
                    r'\Timber_demolish_' + str(scenario_index+"_"+storyline) + '.xlsx')            

        import Emissions

        Emissions.emissions_calculate(
        year_len = 130,
        index_year2018 = 87,
        index_year2060 = 130,
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
        EoL_tim_land = EoL_tim_land,
        Land_tim_deg = Land_tim_deg,
        Deg_tim_CH4 = Deg_tim_CH4,
        Deg_tim_CO2 = Deg_tim_CO2,
        CH4_to_CO2 = CH4_to_CO2,
        CH4_recover = CH4_recover,
        EoL_tim_inci = EoL_tim_inci,
        scenario_index = scenario_index,
        storyline = storyline,
        filepath = 'India-Buildings'
        )

        import Carbonation_India

        Carbonation_India.cement_carbonation(
        year_len = 130,
        index_year2018 = 87,
        Belite_share = Belite_share,
        Belite_alk_save = Belite_alk_save,
        BYF_share = BYF_share,
        BYF_alk_save = BYF_alk_save,
        CCSC_share = CCSC_share,
        CCSC_alk_save = CCSC_alk_save,
        CSAB_share = CSAB_share,
        CSAB_alk_save = CSAB_alk_save,
        Celite_share = Celite_share,
        Celite_alk_save = Celite_alk_save,
        MOMS_share = MOMS_share,
        MOMS_alk_save = MOMS_alk_save,
        CO2_curing_share = CO2_curing_share,
        CO2_curing_coef = CO2_curing_coef,
        CO2_mine_EOL_share = CO2_mine_EOL_share,
        CO2_mine_EOL_coef = CO2_mine_EOL_coef,
        clinker = clinker,
        Cement_demand = Cement_demand,
        waste_saving = waste_saving,
        scenario_index = scenario_index,
        storyline = storyline,
        Life = Life.iloc[:,0],
        reuse_rate = reuseRate,
        recRate = recRate,
        burRate = burRate,
        store_t = store_t,
        filepath = 'India-Buildings'
        )
