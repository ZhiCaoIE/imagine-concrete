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
header = None, sheet_name="USA-Buildings", usecols= "B:HB", nrows = 172, skiprows = 4)

scenario_steps_inputs = pd.read_excel (io = r'C:\Users\caozh\Box\ERSAL Cement 2018-2019\Zhi Cao\Raw data\Data inputs.xlsx', 
header = None, sheet_name="USA-Buildings-STEPS", usecols= "B:HB", nrows = 172, skiprows = 4)

scenario_production_inputs = pd.read_excel (io = r'C:\Users\caozh\Box\ERSAL Cement 2018-2019\Zhi Cao\Raw data\Data inputs.xlsx', 
header = None, sheet_name="USA-Buildings-Production", usecols= "B:HB", nrows = 172, skiprows = 4)

scenario_demand_inputs = pd.read_excel (io = r'C:\Users\caozh\Box\ERSAL Cement 2018-2019\Zhi Cao\Raw data\Data inputs.xlsx', 
header = None, sheet_name="USA-Buildings-Demand", usecols= "B:HB", nrows = 172, skiprows = 4)

lever_list = {
     'Base': [],
     'L1a': [82],
     'L1b': [107],
     'L1c': [83,84,85,86,87],
     'L3':  [93,94,95,96,97,98,99],
     'L2':  [132,133,134,135,136,137],
     'L5a': [157],
     'L5f': [161,165,168,171,174],
     'L4':  [101,102,103,104],
     'L6a': [177],
     'L6b': [179],
     'L6c': [193],
     'L6d': [25,26],
     'L6e': [54,55],
     'L7a': [195],
     'L7b': [196],
     'L7c': [208],
     'L8a': [108,109,110,111,112,113,114,115,116,117,118,119],
     'L8b': [198,199,200,201,202,203,204,205,206]
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

        ##################new single family floor area#####################################
        #unit: m2; SF_WB: Single family-Wood frame-Basement
        new_m2_SF_WB = data_inputs.iloc[0:129,0]

        #unit: m2; SF-WS: Single family-Wood frame-Slab
        new_m2_SF_WS = data_inputs.iloc[0:129,1]

        #unit: m2; SF-WC: Single family-Wood frame-Crawlspace
        new_m2_SF_WC = data_inputs.iloc[0:129,2]

        #unit: m2; SF-CB: Single family-Concrete frame-Basement
        new_m2_SF_CB = data_inputs.iloc[0:129,3]

        #unit: m2; SF-CS: Single family-Concrete frame-Slab
        new_m2_SF_CS = data_inputs.iloc[0:129,4]

        #unit: m2; SF-CC: Single family-Concrete frame-Crawlspace
        new_m2_SF_CC = data_inputs.iloc[0:129,5]

        ##################new multiple family floor area####################################
        #unit: m2; MF-WB: Multi family-Wood frame-Basement
        new_m2_MF_WB = data_inputs.iloc[0:129,6]

        #unit: m2; MF-WS: Multi family-Wood frame-Slab
        new_m2_MF_WS = data_inputs.iloc[0:129,7]

        #unit: m2; MF-WC: Multi family-Wood frame-Crawlspace
        new_m2_MF_WC = data_inputs.iloc[0:129,8]

        #unit: m2; MF-SB: Multi family-Steel frame-Basement
        new_m2_MF_SB = data_inputs.iloc[0:129,9]

        #unit: m2; MF-SS: Multi family-Steel frame-Slab
        new_m2_MF_SS = data_inputs.iloc[0:129,10]

        #unit: m2; MF-SC: Multi family-Steel frame-Crawlspace
        new_m2_MF_SC = data_inputs.iloc[0:129,11]

        #unit: m2; MF-CB: Multi family-Concrete frame-Basement
        new_m2_MF_CB = data_inputs.iloc[0:129,12]

        #unit: m2; MF-CS: Multi family-Concrete frame-Slab
        new_m2_MF_CS = data_inputs.iloc[0:129,13]

        #unit: m2; MF-CC: Multi family-Concrete frame-Crawlspace
        new_m2_MF_CC = data_inputs.iloc[0:129,14]

        ##################new manufactured house floor area###################################
        #unit: m2; MH-WB: Manufactured house-Wood frame-Basement
        new_m2_MH_WB = data_inputs.iloc[0:129,15]

        #unit: m2; MH-WS: Manufactured house-Wood frame-Slab
        new_m2_MH_WS = data_inputs.iloc[0:129,16]

        #unit: m2; MH-WC: Manufactured house-Wood frame-Crawlspace
        new_m2_MH_WC = data_inputs.iloc[0:129,17]

        #unit: m2; MH-CB: Manufactured house-Concrete frame-Basement
        new_m2_MH_CB = data_inputs.iloc[0:129,18]

        #unit: m2; MH-CS: Manufactured house-Concrete frame-Slab
        new_m2_MH_CS = data_inputs.iloc[0:129,19]

        #unit: m2; MH-CC: Manufactured house-Concrete frame-Crawlspace
        new_m2_MH_CC = data_inputs.iloc[0:129,20]

        ##################new non-residential floor area###################################
        #unit: m2; NonR-W: NonR-Wood frame
        new_m2_NonR_W = data_inputs.iloc[0:129,21]

        #unit: m2; NonR-S: NonR-Steel frame
        new_m2_NonR_S = data_inputs.iloc[0:129,22]

        #unit: m2; NonR-C: NonR-Concrete frame
        new_m2_NonR_C = data_inputs.iloc[0:129,23]

        ##################Scenario parameters 2017-2060#########################################
        #unit: 1,000
        pop = data_inputs.iloc[128:172,24]

        #unit: m2/cap
        m2_per_Res = data_inputs.iloc[128:172,25]

        #unit: m2; in-use residential floor area
        Use_m2_Res = pop*m2_per_Res*1000

        #unit: million m2; in-use non-residential floor area
        Use_m2_NonR = data_inputs.iloc[128:172,26]

        #unit: m2
        Use_m2_NonR = Use_m2_NonR*1000000

        #Res-SF; Res-MF; Res-MH
        Res_SF = data_inputs.iloc[128:172,27]
        Res_MF = data_inputs.iloc[128:172,28]
        Res_MH = data_inputs.iloc[128:172,29]

        #Res_SF_WB; Res_SF_WS; Res_SF_WC
        Res_SF_WB = data_inputs.iloc[128:172,30]
        Res_SF_WS = data_inputs.iloc[128:172,31]
        Res_SF_WC = data_inputs.iloc[128:172,32]

        #Res_SF_CB; Res_SF_CS; Res_SF_CC
        Res_SF_CB = data_inputs.iloc[128:172,33]
        Res_SF_CS = data_inputs.iloc[128:172,34]
        Res_SF_CC = data_inputs.iloc[128:172,35]

        #Res_MF_WB; Res_MF_WS; Res_MF_WC
        Res_MF_WB = data_inputs.iloc[128:172,36]
        Res_MF_WS = data_inputs.iloc[128:172,37]
        Res_MF_WC = data_inputs.iloc[128:172,38]

        #Res_MF_SB; Res_MF_SS; Res_MF_SC
        Res_MF_SB = data_inputs.iloc[128:172,39]
        Res_MF_SS = data_inputs.iloc[128:172,40]
        Res_MF_SC = data_inputs.iloc[128:172,41]

        #Res_MF_CB; Res_MF_CS; Res_MF_CC
        Res_MF_CB = data_inputs.iloc[128:172,42]
        Res_MF_CS = data_inputs.iloc[128:172,43]
        Res_MF_CC = data_inputs.iloc[128:172,44]

        #Res_MH_WB; Res_MH_WS; Res_MH_WC
        Res_MH_WB = data_inputs.iloc[128:172,45]
        Res_MH_WS = data_inputs.iloc[128:172,46]
        Res_MH_WC = data_inputs.iloc[128:172,47]

        #Res_MH_CB; Res_MH_CS; Res_MH_CC
        Res_MH_CB = data_inputs.iloc[128:172,48]
        Res_MH_CS = data_inputs.iloc[128:172,49]
        Res_MH_CC = data_inputs.iloc[128:172,50]

        #NonR_W; NonR_S; NonR_C
        NonR_W = data_inputs.iloc[128:172,51]
        NonR_S = data_inputs.iloc[128:172,52]
        NonR_C = data_inputs.iloc[128:172,53]

        #lifetime: years; cement intensity; kg cement/m2
        Life = data_inputs.iloc[:,54:56]

        int_Con_Res_SF_WB = data_inputs.iloc[:,56]
        int_Con_Res_SF_WS = data_inputs.iloc[:,57]
        int_Con_Res_SF_WC = data_inputs.iloc[:,58]
        int_Con_Res_SF_CB = data_inputs.iloc[:,59]
        int_Con_Res_SF_CS = data_inputs.iloc[:,60]
        int_Con_Res_SF_CC = data_inputs.iloc[:,61]

        int_Con_Res_MF_WB = data_inputs.iloc[:,62]
        int_Con_Res_MF_WS = data_inputs.iloc[:,63]
        int_Con_Res_MF_WC = data_inputs.iloc[:,64]
        int_Con_Res_MF_SB = data_inputs.iloc[:,65]
        int_Con_Res_MF_SS = data_inputs.iloc[:,66]
        int_Con_Res_MF_SC = data_inputs.iloc[:,67]
        int_Con_Res_MF_CB = data_inputs.iloc[:,68]
        int_Con_Res_MF_CS = data_inputs.iloc[:,69]
        int_Con_Res_MF_CC = data_inputs.iloc[:,70]

        int_Con_Res_MH_WB = data_inputs.iloc[:,71]
        int_Con_Res_MH_WS = data_inputs.iloc[:,72]
        int_Con_Res_MH_WC = data_inputs.iloc[:,73]
        int_Con_Res_MH_CB = data_inputs.iloc[:,74]
        int_Con_Res_MH_CS = data_inputs.iloc[:,75]
        int_Con_Res_MH_CC = data_inputs.iloc[:,76]

        int_Con_NonR_W = data_inputs.iloc[:,77]
        int_Con_NonR_S = data_inputs.iloc[:,78]
        int_Con_NonR_C = data_inputs.iloc[:,79]

        cemen_content = data_inputs.iloc[:,80]

        #############################################################################################
        #kg CO2/t clinker; 2018-2060; 43 yrs
        Pro_Em_factor = data_inputs.iloc[129:172,81]

        #MJ/t clinker
        Them_eff = data_inputs.iloc[129:172,82]

        Coal_share_cem = data_inputs.iloc[129:172,83]
        Oil_share_cem  = data_inputs.iloc[129:172,84]
        NaGa_share_cem = data_inputs.iloc[129:172,85]
        WaFu_share_cem = data_inputs.iloc[129:172,86]
        Biom_share_cem = data_inputs.iloc[129:172,87]

        Coal_CO2_cem = data_inputs.iloc[129:172,88]
        Oil_CO2_cem  = data_inputs.iloc[129:172,89]
        NaGa_CO2_cem = data_inputs.iloc[129:172,90]
        WaFu_CO2_cem = data_inputs.iloc[129:172,91]
        Biom_CO2_cem = data_inputs.iloc[129:172,92]

        #clinker is also used in the carbonation module; 1931-2060
        clinker = data_inputs.iloc[:,93]

        Gyp_share = data_inputs.iloc[129:172,94]
        Lim_share = data_inputs.iloc[129:172,95]
        Poz_share = data_inputs.iloc[129:172,96]
        Sla_share = data_inputs.iloc[129:172,97]
        Fly_share = data_inputs.iloc[129:172,98]
        Cla_share = data_inputs.iloc[129:172,99]
        Ene_Cla   = data_inputs.iloc[129:172,100]

        CCS_Oxy_percent = data_inputs.iloc[129:172,101]
        CCS_Pos_percent = data_inputs.iloc[129:172,102]
        Eff_Oxy = data_inputs.iloc[129:172,103]
        Eff_Pos = data_inputs.iloc[129:172,104]
        Ene_Oxy = data_inputs.iloc[129:172,105]
        Ene_Pos = data_inputs.iloc[129:172,106]

        Ele_eff        = data_inputs.iloc[129:172,107]
        Coal_share_ele = data_inputs.iloc[129:172,108]
        Oil_share_ele  = data_inputs.iloc[129:172,109]
        Naga_share_ele = data_inputs.iloc[129:172,110]
        Hydr_share_ele = data_inputs.iloc[129:172,111]
        Geot_share_ele = data_inputs.iloc[129:172,112]
        Sol_share_ele  = data_inputs.iloc[129:172,113]
        Wind_share_ele = data_inputs.iloc[129:172,114]
        Tide_share_ele = data_inputs.iloc[129:172,115]
        Nuc_share_ele  = data_inputs.iloc[129:172,116]
        Bio_share_ele  = data_inputs.iloc[129:172,117]
        Was_share_ele  = data_inputs.iloc[129:172,118]
        SoTh_share_ele = data_inputs.iloc[129:172,119]

        Coal_CO2_ele = data_inputs.iloc[129:172,120]
        Oil_CO2_ele  = data_inputs.iloc[129:172,121]
        Naga_CO2_ele = data_inputs.iloc[129:172,122]
        Hydr_CO2_ele = data_inputs.iloc[129:172,123]
        Geot_CO2_ele = data_inputs.iloc[129:172,124]
        Sol_CO2_ele  = data_inputs.iloc[129:172,125]
        Wind_CO2_ele = data_inputs.iloc[129:172,126]
        Tide_CO2_ele = data_inputs.iloc[129:172,127]
        Nuc_CO2_ele  = data_inputs.iloc[129:172,128]
        Bio_CO2_ele  = data_inputs.iloc[129:172,129]
        Was_CO2_ele  = data_inputs.iloc[129:172,130]
        SoTh_CO2_ele = data_inputs.iloc[129:172,131]

        Belite_share = data_inputs.iloc[:,132]
        BYF_share    = data_inputs.iloc[:,133]
        CCSC_share   = data_inputs.iloc[:,134]
        CSAB_share   = data_inputs.iloc[:,135]
        Celite_share = data_inputs.iloc[:,136]
        MOMS_share   = data_inputs.iloc[:,137]

        Belite_pro_save = data_inputs.iloc[:,138]
        BYF_pro_save    = data_inputs.iloc[:,139]
        CCSC_pro_save   = data_inputs.iloc[:,140]
        CSAB_pro_save   = data_inputs.iloc[:,141]
        Celite_pro_save = data_inputs.iloc[:,142]
        MOMS_pro_save   = data_inputs.iloc[:,143]

        Belite_fue_save = data_inputs.iloc[:,144]
        BYF_fue_save    = data_inputs.iloc[:,145]
        CCSC_fue_save   = data_inputs.iloc[:,146]
        CSAB_fue_save   = data_inputs.iloc[:,147]
        Celite_fue_save = data_inputs.iloc[:,148]
        MOMS_fue_save   = data_inputs.iloc[:,149]

        Belite_alk_save = data_inputs.iloc[:,150]
        BYF_alk_save    = data_inputs.iloc[:,151]
        CCSC_alk_save   = data_inputs.iloc[:,152]
        CSAB_alk_save   = data_inputs.iloc[:,153]
        Celite_alk_save = data_inputs.iloc[:,154]
        MOMS_alk_save   = data_inputs.iloc[:,155]

        CCSC_CO2_credit = data_inputs.iloc[:,156]

        CO2_curing_share = data_inputs.iloc[:,157]
        CO2_curing_coef  = data_inputs.iloc[:,158]
        CO2_curing_pen   = data_inputs.iloc[:,159]
        CO2_curing_red   = data_inputs.iloc[:,160]

        CO2_mine_EOL_share   = data_inputs.iloc[:,161]
        CO2_mine_EOL_coef    = data_inputs.iloc[:,162]
        CO2_mine_EOL_pen     = data_inputs.iloc[:,163]
        CO2_mine_EOL_red     = data_inputs.iloc[:,164]

        CO2_mine_slag_share  = data_inputs.iloc[:,165]
        CO2_mine_slag_upt    = data_inputs.iloc[:,166]
        CO2_mine_slag_pen    = data_inputs.iloc[:,167]

        CO2_mine_ash_share  = data_inputs.iloc[:,168]
        CO2_mine_ash_upt    = data_inputs.iloc[:,169]
        CO2_mine_ash_pen    = data_inputs.iloc[:,170]

        CO2_mine_lime_share  = data_inputs.iloc[:,171]
        CO2_mine_lime_upt    = data_inputs.iloc[:,172]
        CO2_mine_lime_pen    = data_inputs.iloc[:,173]

        CO2_mine_red_share  = data_inputs.iloc[:,174]
        CO2_mine_red_upt    = data_inputs.iloc[:,175]
        CO2_mine_red_pen    = data_inputs.iloc[:,176]

        Leaner_rate = data_inputs.iloc[:,177]
        Leaner_con_reduction = data_inputs.iloc[:,178]

        Mat_sub_rate  = data_inputs.iloc[:,179]
        Timber_to_cem_res = data_inputs.iloc[:,180]
        Timber_to_cem_nonr = data_inputs.iloc[:,181]
        CO2_tim_pen   = data_inputs.iloc[:,182]
        CO2_tim_sto   = data_inputs.iloc[:,183]
        EoL_tim_land  = data_inputs.iloc[:,184]
        EoL_tim_inci  = data_inputs.iloc[:,185]
        EoL_tim_recy  = data_inputs.iloc[:,186]
        Land_tim_deg  = data_inputs.iloc[:,187]
        Land_tim_per  = data_inputs.iloc[:,188]
        Deg_tim_CO2   = data_inputs.iloc[:,189]
        Deg_tim_CH4   = data_inputs.iloc[:,190]
        CH4_to_CO2    = data_inputs.iloc[:,191]
        CH4_recover   = data_inputs.iloc[:,192]
        
        EoL = data_inputs.iloc[:,193:198]

        CO2_tran_cem  = data_inputs.iloc[:,198]

        CO2_prod_agg  = data_inputs.iloc[:,199]
        CO2_tran_agg  = data_inputs.iloc[:,200]

        CO2_prod_rec  = data_inputs.iloc[:,201]
        CO2_tran_rec  = data_inputs.iloc[:,202]

        CO2_tran_bur  = data_inputs.iloc[:,203]

        CO2_mix_con   = data_inputs.iloc[:,204]
        CO2_tran_con  = data_inputs.iloc[:,205]

        CO2_onsite_con= data_inputs.iloc[:,206]

        Dem_spread_time = data_inputs.iloc[:,207]
        Dem_spread_ratio= data_inputs.iloc[:,208]

        #assemble data series into dataframe
        store_t = Dem_spread_time*Dem_spread_ratio+\
            (1-Dem_spread_ratio)*0.4

        #assemble data series into dataframe
        new_m2_matrix = pd.concat([new_m2_SF_WB,new_m2_SF_WS,new_m2_SF_WC,new_m2_SF_CB,new_m2_SF_CS,new_m2_SF_CC,\
            new_m2_MF_WB,new_m2_MF_WS,new_m2_MF_WC,new_m2_MF_SB,new_m2_MF_SS,new_m2_MF_SC,new_m2_MF_CB,new_m2_MF_CS,new_m2_MF_CC,\
                        new_m2_MH_WB,new_m2_MH_WS,new_m2_MH_WC,new_m2_MH_CB,new_m2_MH_CS,new_m2_MH_CC,\
                            new_m2_NonR_W,new_m2_NonR_S,new_m2_NonR_C], axis = 1)

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

        int_Con_matrix = pd.concat([int_Con_Res_SF_WB*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Res_SF_WS*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015),\
                                    int_Con_Res_SF_WC*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015),\
                                    int_Con_Res_SF_CB*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Res_SF_CS*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Res_SF_CC*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Res_MF_WB*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Res_MF_WS*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015),\
                                    int_Con_Res_MF_WC*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015),\
                                    int_Con_Res_MF_SB*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Res_MF_SS*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Res_MF_SC*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Res_MF_CB*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Res_MF_CS*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Res_MF_CC*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Res_MH_WB*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Res_MH_WS*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015),\
                                    int_Con_Res_MH_WC*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015),\
                                    int_Con_Res_MH_CB*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Res_MH_CS*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_Res_MH_CC*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_NonR_W*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015),\
                                    int_Con_NonR_S*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015),\
                                    int_Con_NonR_C*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)], axis = 1)

        int_Tim_matrix = pd.concat([int_Con_Res_SF_WB*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_Res_SF_WS*(1-Leaner_con_saving)*0,\
                                    int_Con_Res_SF_WC*(1-Leaner_con_saving)*0,\
                                    int_Con_Res_SF_CB*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_Res_SF_CS*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_Res_SF_CC*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_Res_MF_WB*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_Res_MF_WS*(1-Leaner_con_saving)*0,\
                                    int_Con_Res_MF_WC*(1-Leaner_con_saving)*0,\
                                    int_Con_Res_MF_SB*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_Res_MF_SS*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_Res_MF_SC*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_Res_MF_CB*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_Res_MF_CS*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_Res_MF_CC*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_Res_MH_WB*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_Res_MH_WS*(1-Leaner_con_saving)*0,\
                                    int_Con_Res_MH_WC*(1-Leaner_con_saving)*0,\
                                    int_Con_Res_MH_CB*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_Res_MH_CS*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_Res_MH_CC*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_res,\
                                    int_Con_NonR_W*(1-Leaner_con_saving)*0,\
                                    int_Con_NonR_S*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_nonr,\
                                    int_Con_NonR_C*(1-Leaner_con_saving)*Mat_sub_rate*Timber_to_cem_nonr], axis = 1)


        int_Cem_Res_SF_WB = int_Con_Res_SF_WB*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_SF_WS = int_Con_Res_SF_WS*cemen_content*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_SF_WC = int_Con_Res_SF_WC*cemen_content*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_SF_CB = int_Con_Res_SF_CB*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_SF_CS = int_Con_Res_SF_CS*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_SF_CC = int_Con_Res_SF_CC*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)

        int_Cem_Res_MF_WB = int_Con_Res_MF_WB*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_MF_WS = int_Con_Res_MF_WS*cemen_content*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_MF_WC = int_Con_Res_MF_WC*cemen_content*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_MF_SB = int_Con_Res_MF_SB*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_MF_SS = int_Con_Res_MF_SS*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_MF_SC = int_Con_Res_MF_SC*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_MF_CB = int_Con_Res_MF_CB*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_MF_CS = int_Con_Res_MF_CS*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_MF_CC = int_Con_Res_MF_CC*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)

        int_Cem_Res_MH_WB = int_Con_Res_MH_WB*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_MH_WS = int_Con_Res_MH_WS*cemen_content*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_MH_WC = int_Con_Res_MH_WC*cemen_content*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_MH_CB = int_Con_Res_MH_CB*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_MH_CS = int_Con_Res_MH_CS*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_Res_MH_CC = int_Con_Res_MH_CC*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)

        int_Cem_NonR_W = int_Con_NonR_W*cemen_content*(1-Leaner_con_saving)*(1)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_NonR_S = int_Con_NonR_S*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)
        int_Cem_NonR_C = int_Con_NonR_C*cemen_content*(1-Leaner_con_saving)*(1-Mat_sub_rate)*(1-waste_saving*0.015)*(1-CO2_curing_sav-CO2_mine_EOL_sav)

        int_Cem_matrix = pd.concat([
            int_Cem_Res_SF_WB,int_Cem_Res_SF_WS,int_Cem_Res_SF_WC,\
                int_Cem_Res_SF_CB,int_Cem_Res_SF_CS,int_Cem_Res_SF_CC,\
                    int_Cem_Res_MF_WB,int_Cem_Res_MF_WS,int_Cem_Res_MF_WC,\
                        int_Cem_Res_MF_SB,int_Cem_Res_MF_SS,int_Cem_Res_MF_SC,\
                            int_Cem_Res_MF_CB,int_Cem_Res_MF_CS,int_Cem_Res_MF_CC,\
                                int_Cem_Res_MH_WB,int_Cem_Res_MH_WS,int_Cem_Res_MH_WC,\
                                    int_Cem_Res_MH_CB,int_Cem_Res_MH_CS,int_Cem_Res_MH_CC,\
                                        int_Cem_NonR_W,int_Cem_NonR_S,int_Cem_NonR_C], axis = 1)

        #Use_m2; unit: m2
        Use_m2_Res_SF_WB = Use_m2_Res * Res_SF * Res_SF_WB
        Use_m2_Res_SF_WS = Use_m2_Res * Res_SF * Res_SF_WS
        Use_m2_Res_SF_WC = Use_m2_Res * Res_SF * Res_SF_WC
        Use_m2_Res_SF_CB = Use_m2_Res * Res_SF * Res_SF_CB
        Use_m2_Res_SF_CS = Use_m2_Res * Res_SF * Res_SF_CS
        Use_m2_Res_SF_CC = Use_m2_Res * Res_SF * Res_SF_CC

        Use_m2_Res_MF_WB = Use_m2_Res * Res_MF * Res_MF_WB
        Use_m2_Res_MF_WS = Use_m2_Res * Res_MF * Res_MF_WS
        Use_m2_Res_MF_WC = Use_m2_Res * Res_MF * Res_MF_WC
        Use_m2_Res_MF_SB = Use_m2_Res * Res_MF * Res_MF_SB
        Use_m2_Res_MF_SS = Use_m2_Res * Res_MF * Res_MF_SS
        Use_m2_Res_MF_SC = Use_m2_Res * Res_MF * Res_MF_SC
        Use_m2_Res_MF_CB = Use_m2_Res * Res_MF * Res_MF_CB
        Use_m2_Res_MF_CS = Use_m2_Res * Res_MF * Res_MF_CS
        Use_m2_Res_MF_CC = Use_m2_Res * Res_MF * Res_MF_CC

        Use_m2_Res_MH_WB = Use_m2_Res * Res_MH * Res_MH_WB
        Use_m2_Res_MH_WS = Use_m2_Res * Res_MH * Res_MH_WS
        Use_m2_Res_MH_WC = Use_m2_Res * Res_MH * Res_MH_WC
        Use_m2_Res_MH_CB = Use_m2_Res * Res_MH * Res_MH_CB
        Use_m2_Res_MH_CS = Use_m2_Res * Res_MH * Res_MH_CS
        Use_m2_Res_MH_CC = Use_m2_Res * Res_MH * Res_MH_CC

        Use_m2_NonR_W = Use_m2_NonR * NonR_W
        Use_m2_NonR_S = Use_m2_NonR * NonR_S
        Use_m2_NonR_C = Use_m2_NonR * NonR_C

        Use_m2_matrix = pd.concat([
            Use_m2_Res_SF_WB,Use_m2_Res_SF_WS,Use_m2_Res_SF_WC,\
                Use_m2_Res_SF_CB,Use_m2_Res_SF_CS,Use_m2_Res_SF_CC,\
                    Use_m2_Res_MF_WB,Use_m2_Res_MF_WS,Use_m2_Res_MF_WC,\
                        Use_m2_Res_MF_SB,Use_m2_Res_MF_SS,Use_m2_Res_MF_SC,\
                            Use_m2_Res_MF_CB,Use_m2_Res_MF_CS,Use_m2_Res_MF_CC,\
                                Use_m2_Res_MH_WB,Use_m2_Res_MH_WS,Use_m2_Res_MH_WC,\
                                    Use_m2_Res_MH_CB,Use_m2_Res_MH_CS,Use_m2_Res_MH_CC,\
                                        Use_m2_NonR_W,Use_m2_NonR_S,Use_m2_NonR_C], axis = 1)

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
            year_complete = np.arange(1889,2018)
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

                demolish_Con_list = demolish_m2_list * int_Con /1000
                demolish_Con      = sum(demolish_Con_list)
                demol_Con_extended= np.append(demol_Con_extended,demolish_Con)

                demolish_Tim_list = demolish_m2_list * int_Tim /1000
                demolish_Tim      = sum(demolish_Tim_list)
                demol_Tim_extended= np.append(demol_Tim_extended,demolish_Tim)

                new_m2 = Use_m2_matrix.iloc[k-2017,i] - Use_m2_matrix.iloc[k-2018,i] + demolish_m2
                if new_m2 < 0:
                    new_m2 = 0
                new_m2_extended = np.append(new_m2_extended, new_m2)
                
                reuse_Con_list    = demolish_m2_list * int_Con /1000 * reuse_rate
                reuse_Cem_list    = demolish_m2_list * int_Cem /1000 * reuse_rate
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
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\USA-Buildings' + \
                    r'\New_m2_list_' + str(scenario_index+"_"+storyline) + '.xlsx')

        pd.DataFrame(Concrete_demand).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\USA-Buildings' + \
                    r'\Concrete_demand_' + str(scenario_index+"_"+storyline) + '.xlsx')

        pd.DataFrame(Cement_demand).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\USA-Buildings' + \
                    r'\Cement_demand_' + str(scenario_index+"_"+storyline) + '.xlsx')

        pd.DataFrame(Timber_demand).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\USA-Buildings' + \
                    r'\Timber_demand_' + str(scenario_index+"_"+storyline) + '.xlsx')

        pd.DataFrame(Concrete_demolish).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\USA-Buildings' + \
                    r'\Concrete_demolish_' + str(scenario_index+"_"+storyline) + '.xlsx')

        pd.DataFrame(Timber_demolish).to_excel(
            r'C:\Users\caozh\Documents\2020-ClimateWork\Carbonation\USA-Buildings' + \
                    r'\Timber_demolish_' + str(scenario_index+"_"+storyline) + '.xlsx')            

        import Emissions

        Emissions.emissions_calculate(
        year_len = 130,
        index_year2018 = 129,
        index_year2060 = 172,
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
        filepath = 'USA-Buildings'
        )

        import Carbonation_USA
        #index_year2018 as in Cement_demand[42:172]
        Carbonation_USA.cement_carbonation(
        year_len = 130,
        index_year2018 = 87,
        Belite_share = (Belite_share[42:172]).reset_index(drop=True),
        Belite_alk_save = (Belite_alk_save[42:172]).reset_index(drop=True),
        BYF_share = (BYF_share[42:172]).reset_index(drop=True),
        BYF_alk_save = (BYF_alk_save[42:172]).reset_index(drop=True),
        CCSC_share = (CCSC_share[42:172]).reset_index(drop=True),
        CCSC_alk_save = (CCSC_alk_save[42:172]).reset_index(drop=True),
        CSAB_share = (CSAB_share[42:172]).reset_index(drop=True),
        CSAB_alk_save = (CSAB_alk_save[42:172]).reset_index(drop=True),
        Celite_share = (Celite_share[42:172]).reset_index(drop=True),
        Celite_alk_save = (Celite_alk_save[42:172]).reset_index(drop=True),
        MOMS_share = (MOMS_share[42:172]).reset_index(drop=True),
        MOMS_alk_save = (MOMS_alk_save[42:172]).reset_index(drop=True),
        CO2_curing_share = (CO2_curing_share[42:172]).reset_index(drop=True),
        CO2_curing_coef = (CO2_curing_coef[42:172]).reset_index(drop=True),
        CO2_mine_EOL_share = (CO2_mine_EOL_share[42:172]).reset_index(drop=True),
        CO2_mine_EOL_coef = (CO2_mine_EOL_coef[42:172]).reset_index(drop=True),
        clinker = (clinker[42:172]).reset_index(drop=True),
        Cement_demand = (Cement_demand[42:172]).reset_index(drop=True),
        waste_saving = (waste_saving[42:172]).reset_index(drop=True),
        scenario_index = scenario_index,
        storyline = storyline,
        Life = (Life.iloc[42:172,0]).reset_index(drop=True),
        reuse_rate = (reuseRate[42:172]).reset_index(drop=True),
        recRate = (recRate[42:172]).reset_index(drop=True),
        burRate = (burRate[42:172]).reset_index(drop=True),
        store_t = (store_t[42:172]).reset_index(drop=True),
        filepath = 'USA-Buildings'
        )
