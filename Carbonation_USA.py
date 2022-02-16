import pandas as pd
import numpy as np
from scipy.stats import norm

import random
import statistics
import pathlib

def cement_carbonation(
    year_len,index_year2018,
    Belite_share,Belite_alk_save,
    BYF_share,BYF_alk_save,
    CCSC_share,CCSC_alk_save,
    CSAB_share,CSAB_alk_save,
    Celite_share,Celite_alk_save,
    MOMS_share,MOMS_alk_save,
    CO2_curing_share,CO2_curing_coef,
    CO2_mine_EOL_share,CO2_mine_EOL_coef,
    clinker,Cement_demand,waste_saving,
    scenario_index,storyline,
    Life,
    reuse_rate,recRate,burRate,store_t,
    filepath
):
    ##############################################################################
    ########################simulate cement carbonation###########################
    ##############################################################################
    #determine the impacts of using green cement chemistries on alkaline content in cement
    Chem_alk_reduction = Belite_share*Belite_alk_save + BYF_share*BYF_alk_save +\
        CCSC_share*CCSC_alk_save + CSAB_share*CSAB_alk_save +\
            Celite_share*Celite_alk_save + MOMS_share*MOMS_alk_save

    #determine the impacts of CO2 curing on coefficients
    Curing_coef_increase = CO2_curing_share*CO2_curing_coef

    #determine the impacts of CO2 mineralization on coefficients
    Mine_coef_increase   = CO2_mine_EOL_share*CO2_mine_EOL_coef

    #uptake X44-47: cement additives (Beta ad)
    #uptake X49-52: CO2 concentration (Beta CO2)
    #uptake X54-57: coating and cover (Beta CC)
    #alpah is scale; beta is shape
    Beta_ad_random = []
    Beta_CO2_random = []
    Beat_CC_random = []

    for i in range(0,1000):
        random.seed(a=i);Beta_ad_random.append(random.weibullvariate(alpha = 1.16, beta = 20))
        random.seed(a=i);Beta_CO2_random.append(random.weibullvariate(alpha = 1.18, beta = 25))
        random.seed(a=i);Beat_CC_random.append(random.weibullvariate(alpha = 1, beta = 6))

    Beta_ad = statistics.median(Beta_ad_random)
    Beta_CO2 = statistics.median(Beta_CO2_random)
    Beta_CC = statistics.median(Beat_CC_random)

    impact_factor = Beta_ad * Beta_CO2 * Beta_CC

    #uptake X32-33, X35-36, X38-39, X41-42
    #carbonation rate coefficients under exposure condition (Beta csec)
    coefC15_random = []; coefC16_random = []; coefC23_random = []; coefC35_random = []

    for i in range(0,1000):
        random.seed(a=i);coefC15_random.append(random.uniform(a = 7.1, b = 2.15))
        random.seed(a=i);coefC16_random.append(random.uniform(a = 6.9, b = 3.5))
        random.seed(a=i);coefC23_random.append(random.uniform(a = 5.4, b = 2.7))
        random.seed(a=i);coefC35_random.append(random.uniform(a = 3.8, b = 2.5))

    coefC15 = statistics.median(coefC15_random) * impact_factor * (1+Curing_coef_increase)
    coefC16 = statistics.median(coefC16_random) * impact_factor * (1+Curing_coef_increase)
    coefC23 = statistics.median(coefC23_random) * impact_factor * (1+Curing_coef_increase)
    coefC35 = statistics.median(coefC35_random) * impact_factor * (1+Curing_coef_increase)

    #uptake X124-127: carbonation rate coefficients under buried condition
    coefC15B = 3   * impact_factor * (1+Mine_coef_increase)
    coefC16B = 1.5 * impact_factor * (1+Mine_coef_increase)
    coefC23B = 1   * impact_factor * (1+Mine_coef_increase)
    coefC35B = 0.75* impact_factor * (1+Mine_coef_increase)

    #uptake X59-60: wall thickness
    wallthick_random = []

    for i in range(0,1000):
        random.seed(a=i);wallthick_random.append(random.uniform(a = 610, b = 60)/2)

    wallthick = statistics.median(wallthick_random)

    #uptake X79-81: average CaO content of clinker in cement (fCaO)
    caocontent_random = []

    for i in range(0,1000):
        random.seed(a=i);caocontent_random.append(random.triangular(low = 0.6, high = 0.72, mode = 0.675))

    caocontent = statistics.median(caocontent_random)*(1-Chem_alk_reduction)

    #uptake X83-86: proportion of CaO within fully carbonated cement that converts to CaCO3 for concrete cement (?)
    #uptake X87: ratio of mole mass of CO2 to CaO (Mr)
    convertconcrete_random = []

    for i in range(0,1000):
        random.seed(a=i);convertconcrete_random.append(random.weibullvariate(alpha = 0.86, beta = 25))

    convertconcrete = statistics.median(convertconcrete_random) * 0.784808140177683

    #uptake X2-5: cement for concrete; X7-X10: cement for mortar; percentage
    #uptake X12-15: <=C15; X17-20: C16-C23, X22-25: C23-C35, X27-X30: >C35; percentage
    C15percentage_random = []; C16percentage_random = []
    C23percentage_random = []; C35percentage_random = []

    for i in range(0,1000):
        random.seed(a=i);C15percentage_random.append(random.weibullvariate(alpha = 0.891, beta = 25.5)*
        random.weibullvariate(alpha = 0.222, beta = 12))
        random.seed(a=i);C16percentage_random.append(random.weibullvariate(alpha = 0.891, beta = 25.5)*
        random.weibullvariate(alpha = 0.405, beta = 12))
        random.seed(a=i);C23percentage_random.append(random.weibullvariate(alpha = 0.891, beta = 25.5)*
        random.weibullvariate(alpha = 0.295, beta = 8))
        random.seed(a=i);C35percentage_random.append(random.weibullvariate(alpha = 0.891, beta = 25.5)*
        random.weibullvariate(alpha = 0.127, beta = 16))

    C15percentage = statistics.median(C15percentage_random)
    C16percentage = statistics.median(C16percentage_random)
    C23percentage = statistics.median(C23percentage_random)
    C35percentage = statistics.median(C35percentage_random)

    #uptake X129-132: rendering, palstering and decoration; percentage
    #uptake X134-137: masonry; percentage
    #uptake X139-142: maintenance and repairing; percentage
    RENpercentage_random = []
    MASpercentage_random = []
    MAIpercentage_random = []

    for i in range(0,1000):
        random.seed(a=i);RENpercentage_random.append((1-random.weibullvariate(alpha = 0.891, beta = 25.5))*
        random.weibullvariate(alpha = 0.394, beta = 12))
        random.seed(a=i);MASpercentage_random.append((1-random.weibullvariate(alpha = 0.891, beta = 25.5))*
        random.weibullvariate(alpha = 0.288, beta = 12))
        random.seed(a=i);MAIpercentage_random.append((1-random.weibullvariate(alpha = 0.891, beta = 25.5))*
        random.weibullvariate(alpha = 0.318, beta = 12))

    RENpercentage = statistics.median(RENpercentage_random)
    MASpercentage = statistics.median(MASpercentage_random)
    MAIpercentage = statistics.median(MAIpercentage_random)

    #uptake X144-147: thickness of motar used for rendering, palstering and decoration
    #uptake X149-152: thickness of motar used for masonry (use X59-60)
    #uptake X154-157: thickness of motar used for maintenance and repairing
    RENthick_random = []
    MASthick_random = []
    MAIthick_random = []

    for i in range(0,1000):
        random.seed(a=i);RENthick_random.append(random.weibullvariate(alpha = 22, beta = 4))
        random.seed(a=i);MASthick_random.append(random.uniform(a = 610, b = 60)/2)
        random.seed(a=i);MAIthick_random.append(random.weibullvariate(alpha = 26.8, beta = 7))

    RENthick = statistics.median(RENthick_random)
    MASthick = statistics.median(MASthick_random)
    MAIthick = statistics.median(MAIthick_random)

    #uptake X159-161: carbonation coefficient of mortar
    coefMortar_random = []

    for i in range(0,1000):
        random.seed(a=i);coefMortar_random.append(random.triangular(low = 6.1, high = 36.8, mode = 19.6))

    coefMortar = statistics.median(coefMortar_random) * (1+Curing_coef_increase)

    #uptake X163-166: proportion of CaO within fully carbonated cement that converts to CaCO3 for mortar cement
    convertmortar_random = []

    for i in range(0,1000):
        random.seed(a=i);convertmortar_random.append(random.weibullvariate(alpha = 0.92, beta = 20))

    convertmortar = statistics.median(convertmortar_random) * 0.784808140177683

    #uptake X92-99: percentage of waste concrete (for new concrete)
    #uptake X100-107: percentage of waste concrete (for road base, backfills materials and other use)
    #uptake X108-115: percentage of waste concrete (landfill, dumped and stacking)
    #uptake X116-123: percentage of waste concrete (asphalt concrete)
    radiusRec_random = []; radiusBur_random = []
    radiusRecA_random = []; radiusRecB_random = []
    radiusBurA_random = []; radiusBurB_random = []

    for i in range(0,1000):
        random.seed(a=i);radiusRec_random.append(
            random.uniform(a = 0.1, b= 0.36)*2.5+random.uniform(a = 0.05, b= 0.3)*3.75+
            random.uniform(a = 0.2, b= 0.44)*7.5+random.uniform(a = 0.1, b= 0.3)*15)

        random.seed(a=i);radiusBur_random.append((
            (random.uniform(a = 0.1, b= 0.21)*0.5+random.uniform(a = 0.25, b= 0.3)*2.75+
            random.uniform(a = 0.2, b= 0.44)*10+random.uniform(a = 0.05, b= 0.45)*20.75)*0.51+

            (random.uniform(a = 0.122, b= 0.256)*5+random.uniform(a = 0.195, b= 0.354)*10+
            random.uniform(a = 0.106, b= 0.225)*20+random.uniform(a = 0.248, b= 0.484)*25)*0.4+

            (random.uniform(a = 0.1, b= 0.36)*2.5+random.uniform(a = 0.05, b= 0.3)*3.75+
            random.uniform(a = 0.2, b= 0.44)*7.5+random.uniform(a = 0.1, b= 0.3)*15)*0.054)/
            (0.51 + 0.4 + 0.054)
        )

        random.seed(a=i);radiusRecA_random.append(
            random.uniform(a = 0.1, b= 0.36)*2.5+random.uniform(a = 0.05, b= 0.3)*2.5+
            random.uniform(a = 0.2, b= 0.44)*5+random.uniform(a = 0.1, b= 0.3)*10)
        
        random.seed(a=i);radiusRecB_random.append(
            random.uniform(a = 0.1, b= 0.36)*2.5+random.uniform(a = 0.05, b= 0.3)*5+
            random.uniform(a = 0.2, b= 0.44)*10+random.uniform(a = 0.1, b= 0.3)*20)

        random.seed(a=i);radiusBurA_random.append((
            (random.uniform(a = 0.1, b= 0.21)*0.5+
            random.uniform(a = 0.25, b= 0.3)*0.5+
            random.uniform(a = 0.2, b= 0.44)*5+
            random.uniform(a = 0.05, b= 0.45)*15)*0.51+

            (random.uniform(a = 0.122, b= 0.256)*5+
            random.uniform(a = 0.195, b= 0.354)*5+
            random.uniform(a = 0.106, b= 0.225)*15+
            random.uniform(a = 0.248, b= 0.484)*25)*0.4+

            (random.uniform(a = 0.1, b= 0.36)*2.5+
            random.uniform(a = 0.05, b= 0.3)*2.5+
            random.uniform(a = 0.2, b= 0.44)*5+
            random.uniform(a = 0.1, b= 0.3)*10)*0.054)/
            (0.51 + 0.4 + 0.054)
        )

        random.seed(a=i);radiusBurB_random.append((
            (random.uniform(a = 0.1, b= 0.21)*0.5+
            random.uniform(a = 0.25, b= 0.3)*5+
            random.uniform(a = 0.2, b= 0.44)*15+
            random.uniform(a = 0.05, b= 0.45)*26.5)*0.51+

            (random.uniform(a = 0.122, b= 0.256)*5+
            random.uniform(a = 0.195, b= 0.354)*15+
            random.uniform(a = 0.106, b= 0.225)*25+
            random.uniform(a = 0.248, b= 0.484)*25)*0.4+

            (random.uniform(a = 0.1, b= 0.36)*2.5+
            random.uniform(a = 0.05, b= 0.3)*5+
            random.uniform(a = 0.2, b= 0.44)*10+
            random.uniform(a = 0.1, b= 0.3)*20)*0.054)/
            (0.51 + 0.4 + 0.054)
        )

    radiusRecA = statistics.median(radiusRecA_random)
    radiusRecB = statistics.median(radiusRecB_random)
    radiusBurA = statistics.median(radiusBurA_random)
    radiusBurB = statistics.median(radiusBurB_random)

    shellthickRec = radiusRecB-radiusRecA
    shellthickBur = radiusBurB-radiusBurA

    #uptake X88: fate of waste concrete (for new concrete)
    #has been moved up

    #determine the carbonation capacity of one tonne of cement
    adjustC = caocontent*convertconcrete*clinker
    adjustM = caocontent*convertmortar*clinker

    #parameters for CO2 uptake by concrete waste
    waste_rate_random = []
    CONpercentage_random = []

    for i in range(0,1000):
        random.seed(a=i);waste_rate_random.append(random.triangular(low = 0.01, high = 0.03, mode = 0.015))
        random.seed(a=i);CONpercentage_random.append(random.weibullvariate(alpha = 0.891, beta = 25.5))

    waste_rate = statistics.median(waste_rate_random)
    CONpercentage = statistics.median(CONpercentage_random)

    CKD_rate_random = []
    CKD_lanfill_random = []
    caocontent_CKD_random = []

    #parameters for CO2 uptake by CKD
    for i in range(0,1000):
        random.seed(a=i);CKD_rate_random.append(random.triangular(low = 0.041, high = 0.115, mode =0.06))
        random.seed(a=i);CKD_lanfill_random.append(random.triangular(low = 0.52, high = 0.9, mode =0.8))
        random.seed(a=i);caocontent_CKD_random.append(random.normalvariate(mu = 0.44, sigma = 0.0801))

    CKD_rate = statistics.median(CKD_rate_random)
    CKD_lanfill = statistics.median(CKD_lanfill_random)
    caocontent_CKD = statistics.median(caocontent_CKD_random)*(1-Chem_alk_reduction)

    #####################################################################################
    ##################################marked intentionally###############################
    #####################################################################################
    #####################################################################################
    #CO2 upake simulation
    #for i in range(0, len(cem_app_matrix.columns)):
    #cem_app_matrix.iloc[:,i]
    #define cement apparent consumption
    cem_app_individual = Cement_demand

    #CO2 uptake by concrete waste and mortar waste; unit: t CO2
    uptake_wasteConcrete = cem_app_individual * CONpercentage * waste_rate * (1-waste_saving) * adjustC
    uptake_wasteMortar = cem_app_individual * (1-CONpercentage) * waste_rate * (1-waste_saving) * adjustM

    uptake_wasteConcrete_Matrix = np.zeros((year_len,year_len))
    uptake_wasteMortar_Matrix = np.zeros((year_len,year_len))

    print("success")

    for k in range(0,year_len):
        for n in range(0,year_len-k):
            if n <= 4:
                uptake_wasteConcrete_Matrix[k,n+k] = uptake_wasteConcrete[k]/5
            if n <= 0:
                uptake_wasteMortar_Matrix[k,n+k] = uptake_wasteMortar[k]


    
    pd.DataFrame(uptake_wasteConcrete_Matrix).to_excel('C:\\Users\\caozh\\Documents\\2020-ClimateWork\\Carbonation\\' + filepath + '\\wasteConcrete_'+\
        str(scenario_index+"_"+storyline) + '.xlsx')

    pd.DataFrame(uptake_wasteMortar_Matrix).to_excel('C:\\Users\\caozh\\Documents\\2020-ClimateWork\\Carbonation\\' + filepath + '\\wasteMortar_'+\
        str(scenario_index+"_"+storyline) +'.xlsx')

    #CO2 uptake by CKD; unit: t CO2
    uptake_CKD = cem_app_individual * clinker * CKD_rate * CKD_lanfill * caocontent_CKD * 0.784808140177683
    pd.DataFrame(uptake_CKD).T.to_excel('C:\\Users\\caozh\\Documents\\2020-ClimateWork\\Carbonation\\' + filepath + '\\CKD_'+\
        str(scenario_index+"_"+storyline) + '.xlsx')

    #CO2 uptake by in-use concrete and in-use mortar
    uptake_mortar_Ren_use_Matrix = np.zeros((year_len,year_len))
    uptake_mortar_Mas_use_Matrix = np.zeros((year_len,year_len))
    uptake_mortar_Mai_use_Matrix = np.zeros((year_len,year_len))

    uptake_concrete_C15_use_Matrix = np.zeros((year_len,year_len))
    uptake_concrete_C16_use_Matrix = np.zeros((year_len,year_len))
    uptake_concrete_C23_use_Matrix = np.zeros((year_len,year_len))
    uptake_concrete_C35_use_Matrix = np.zeros((year_len,year_len))

    uptake_mortar_demolish_Matrix = np.zeros((year_len,year_len))
    uptake_concrete_demolish_Matrix = np.zeros((year_len,year_len))

    #loop by the year when cement was consumed
    for k in range(0,year_len):
        print("k="); print(k)
        
        Life_k = Life[k]

        #uptake by demolished concrete (recycled)
        uptake_concrete_C15_Rec_Matrix = np.zeros((year_len,year_len))
        uptake_concrete_C16_Rec_Matrix = np.zeros((year_len,year_len))
        uptake_concrete_C23_Rec_Matrix = np.zeros((year_len,year_len))
        uptake_concrete_C35_Rec_Matrix = np.zeros((year_len,year_len))

        #uptake by demolished concrete (buried)
        uptake_concrete_C15_Bur_Matrix = np.zeros((year_len,year_len))
        uptake_concrete_C16_Bur_Matrix = np.zeros((year_len,year_len))
        uptake_concrete_C23_Bur_Matrix = np.zeros((year_len,year_len))
        uptake_concrete_C35_Bur_Matrix = np.zeros((year_len,year_len))

        #uptake by demolished mortar (recycled)
        uptake_mortar_Ren_Rec_Matrix = np.zeros((year_len,year_len))
        uptake_mortar_Mas_Rec_Matrix = np.zeros((year_len,year_len))
        uptake_mortar_Mai_Rec_Matrix = np.zeros((year_len,year_len))

        #uptake by demolished mortar (buried)
        uptake_mortar_Ren_Bur_Matrix = np.zeros((year_len,year_len))
        uptake_mortar_Mas_Bur_Matrix = np.zeros((year_len,year_len))
        uptake_mortar_Mai_Bur_Matrix = np.zeros((year_len,year_len))

        #determine reuse rate at the k-th year
        reuse_rate_k = reuse_rate[k]

        #loop by the time span of cement being 'alive' (in-use stock)
        if k >= index_year2018 or scenario_index == 'Base':
            st = 0
        else:
            st = index_year2018-k
        for n in range(st,year_len-k):
            #in-use mortar carbonation
            #k = 0; n = 0
            CO2Ren = (np.sqrt(n+1)-np.sqrt(n)) * coefMortar[k]/RENthick * adjustM[k]
            CO2Mas = (np.sqrt(n+1)-np.sqrt(n)) * coefMortar[k]/MASthick * adjustM[k]
            CO2Mai = (np.sqrt(n+1)-np.sqrt(n)) * coefMortar[k]/MAIthick * adjustM[k]
            depth1Ren = np.sqrt(n)  *coefMortar[k]; depth1Mas = np.sqrt(n)  *coefMortar[k]; depth1Mai = np.sqrt(n)  *coefMortar[k]
            depth2Ren = np.sqrt(n+1)*coefMortar[k]; depth2Mas = np.sqrt(n+1)*coefMortar[k]; depth2Mai = np.sqrt(n+1)*coefMortar[k]

            #carbonized mortar in last year is greater than the thickness of mortar
            if depth1Ren >= RENthick:
                CO2Ren = 0
            if depth1Mas >= MASthick:
                CO2Mas = 0
            if depth1Mai >= MAIthick:
                CO2Mai = 0
            #carbonized mortar in last year is less than the thickness of mortar but those in this year is thicker
            if depth1Ren <= RENthick and depth2Ren >= RENthick:
                CO2Ren = (1-depth1Ren/RENthick) * adjustM[k]
            if depth1Mas <= MASthick and depth2Mas >= MASthick:
                CO2Mas = (1-depth1Mas/MASthick) * adjustM[k]
            if depth1Mai <= MAIthick and depth2Mai >= MAIthick:
                CO2Mai = (1-depth1Mai/MAIthick) * adjustM[k]
                
            #reuse rate refers to % of demolished components being reused
            #reuse fraction refers to cement embodied in reused components over apparent cement consumption
            reuse_fraction = (norm.cdf(n,Life_k,Life_k*0.2) - norm.cdf(n-1,Life_k,Life_k*0.2))*reuse_rate_k
            
            uptake_mortar_Ren_use_Matrix[k,n+k] = cem_app_individual[k]*RENpercentage*(1-norm.cdf(n,Life_k,Life_k*0.2))*(1+reuse_fraction)*CO2Ren
            uptake_mortar_Mas_use_Matrix[k,n+k] = cem_app_individual[k]*MASpercentage*(1-norm.cdf(n,Life_k,Life_k*0.2))*(1+reuse_fraction)*CO2Mas
            uptake_mortar_Mai_use_Matrix[k,n+k] = cem_app_individual[k]*MAIpercentage*(1-norm.cdf(n,Life_k,Life_k*0.2))*(1+reuse_fraction)*CO2Mai

            #in-use concrete carbonation
            CO2C15 = (np.sqrt(n+1)-np.sqrt(n)) * coefC15[k]/wallthick*adjustC[k]
            CO2C16 = (np.sqrt(n+1)-np.sqrt(n)) * coefC16[k]/wallthick*adjustC[k]
            CO2C23 = (np.sqrt(n+1)-np.sqrt(n)) * coefC23[k]/wallthick*adjustC[k]
            CO2C35 = (np.sqrt(n+1)-np.sqrt(n)) * coefC35[k]/wallthick*adjustC[k]
            depth1C15 = np.sqrt(n)  *coefC15[k]; depth1C16 = np.sqrt(n)  *coefC16[k]
            depth1C23 = np.sqrt(n)  *coefC23[k]; depth1C35 = np.sqrt(n)  *coefC35[k]
            depth2C15 = np.sqrt(n+1)*coefC15[k]; depth2C16 = np.sqrt(n+1)*coefC16[k]
            depth2C23 = np.sqrt(n+1)*coefC23[k]; depth2C35 = np.sqrt(n+1)*coefC35[k]

            #carbonized concrete in last year is greater than the thickness of wall
            if depth1C15 >= wallthick:
                CO2C15 = 0
            if depth1C16 >= wallthick:
                CO2C16 = 0
            if depth1C23 >= wallthick:
                CO2C23 = 0
            if depth1C35 >= wallthick:
                CO2C35 = 0
            #carbonized concrete in last year is less than the thickness of wall but those in this year is thicker
            if depth1C15 < wallthick and depth2C15 >= wallthick:
                CO2C15 = (1-depth1C15/wallthick) * adjustC[k]
            if depth1C16 < wallthick and depth2C16 >= wallthick:
                CO2C16 = (1-depth1C16/wallthick) * adjustC[k]
            if depth1C23 < wallthick and depth2C23 >= wallthick:
                CO2C23 = (1-depth1C23/wallthick) * adjustC[k]
            if depth1C35 < wallthick and depth2C35 >= wallthick:
                CO2C35 = (1-depth1C35/wallthick) * adjustC[k]

            uptake_concrete_C15_use_Matrix[k,k+n] = cem_app_individual[k]*C15percentage*(1-norm.cdf(n,Life_k,Life_k*0.2))*(1+reuse_fraction)*CO2C15
            uptake_concrete_C16_use_Matrix[k,k+n] = cem_app_individual[k]*C16percentage*(1-norm.cdf(n,Life_k,Life_k*0.2))*(1+reuse_fraction)*CO2C16
            uptake_concrete_C23_use_Matrix[k,k+n] = cem_app_individual[k]*C23percentage*(1-norm.cdf(n,Life_k,Life_k*0.2))*(1+reuse_fraction)*CO2C23
            uptake_concrete_C35_use_Matrix[k,k+n] = cem_app_individual[k]*C35percentage*(1-norm.cdf(n,Life_k,Life_k*0.2))*(1+reuse_fraction)*CO2C35

            #demolition carbonation
            demolish_rate = norm.cdf(n,Life_k,Life_k*0.2) - norm.cdf(n-1,Life_k,Life_k*0.2) - reuse_fraction
            
            #loop by the time span of cement in the secondary stage
            for t in range(0,year_len-k-n):
                if t == 0:
                    #depths of demolished mortar
                    depth1RenRec = np.sqrt(n)*coefMortar[k]/RENthick*shellthickRec
                    depth1MasRec = np.sqrt(n)*coefMortar[k]/MASthick*shellthickRec
                    depth1MaiRec = np.sqrt(n)*coefMortar[k]/MAIthick*shellthickRec
                    
                    depth1RenBur = np.sqrt(n)*coefMortar[k]/RENthick*shellthickBur
                    depth1MasBur = np.sqrt(n)*coefMortar[k]/MASthick*shellthickBur
                    depth1MaiBur = np.sqrt(n)*coefMortar[k]/MAIthick*shellthickBur

                    #depths of demolished concrete
                    depth1C15Rec = np.sqrt(n)*coefC15[k]/wallthick*shellthickRec
                    depth1C16Rec = np.sqrt(n)*coefC16[k]/wallthick*shellthickRec
                    depth1C23Rec = np.sqrt(n)*coefC23[k]/wallthick*shellthickRec
                    depth1C35Rec = np.sqrt(n)*coefC35[k]/wallthick*shellthickRec
        
                    depth1C15Bur = np.sqrt(n)*coefC15[k]/wallthick*shellthickBur
                    depth1C16Bur = np.sqrt(n)*coefC16[k]/wallthick*shellthickBur
                    depth1C23Bur = np.sqrt(n)*coefC23[k]/wallthick*shellthickBur
                    depth1C35Bur = np.sqrt(n)*coefC35[k]/wallthick*shellthickBur
                else:
                    depth1RenRec = np.sqrt(n)*coefMortar[k]/RENthick*shellthickRec+(np.sqrt(n+t-1)-np.sqrt(n))*coefMortar[k]
                    depth1MasRec = np.sqrt(n)*coefMortar[k]/MASthick*shellthickRec+(np.sqrt(n+t-1)-np.sqrt(n))*coefMortar[k]
                    depth1MaiRec = np.sqrt(n)*coefMortar[k]/MAIthick*shellthickRec+(np.sqrt(n+t-1)-np.sqrt(n))*coefMortar[k]
                    
                    depth1RenBur = np.sqrt(n)*coefMortar[k]/RENthick*shellthickBur+(np.sqrt(n+store_t[k])-np.sqrt(n))*coefMortar[k]+\
                        (np.sqrt(n+t)-np.sqrt(n+store_t[k]))*coefMortar[k]
                    depth1MasBur = np.sqrt(n)*coefMortar[k]/MASthick*shellthickBur+(np.sqrt(n+store_t[k])-np.sqrt(n))*coefMortar[k]+\
                        (np.sqrt(n+t)-np.sqrt(n+store_t[k]))*coefMortar[k]
                    depth1MaiBur = np.sqrt(n)*coefMortar[k]/MAIthick*shellthickBur+(np.sqrt(n+store_t[k])-np.sqrt(n))*coefMortar[k]+\
                        (np.sqrt(n+t)-np.sqrt(n+store_t[k]))*coefMortar[k]
                    
                    depth1C15Rec = np.sqrt(n)*coefC15[k]/wallthick*shellthickRec+(np.sqrt(n+t-1)  -np.sqrt(n))*coefC15[k]
                    depth1C16Rec = np.sqrt(n)*coefC16[k]/wallthick*shellthickRec+(np.sqrt(n+t-1)  -np.sqrt(n))*coefC16[k]
                    depth1C23Rec = np.sqrt(n)*coefC23[k]/wallthick*shellthickRec+(np.sqrt(n+t-1)  -np.sqrt(n))*coefC23[k]
                    depth1C35Rec = np.sqrt(n)*coefC35[k]/wallthick*shellthickRec+(np.sqrt(n+t-1)  -np.sqrt(n))*coefC35[k]
                    
                    depth1C15Bur = np.sqrt(n)*coefC15[k]/wallthick*shellthickBur+(np.sqrt(n+store_t[k])  -np.sqrt(n))*coefC15[k]+\
                        (np.sqrt(n+t)-np.sqrt(n+store_t[k]))*coefC15B[k]
                    depth1C16Bur = np.sqrt(n)*coefC16[k]/wallthick*shellthickBur+(np.sqrt(n+store_t[k])  -np.sqrt(n))*coefC16[k]+\
                        (np.sqrt(n+t)-np.sqrt(n+store_t[k]))*coefC16B[k]
                    depth1C23Bur = np.sqrt(n)*coefC23[k]/wallthick*shellthickBur+(np.sqrt(n+store_t[k])  -np.sqrt(n))*coefC23[k]+\
                        (np.sqrt(n+t)-np.sqrt(n+store_t[k]))*coefC23B[k]
                    depth1C35Bur = np.sqrt(n)*coefC35[k]/wallthick*shellthickBur+(np.sqrt(n+store_t[k])  -np.sqrt(n))*coefC35[k]+\
                        (np.sqrt(n+t)-np.sqrt(n+store_t[k]))*coefC35B[k] 

                depth2RenRec = np.sqrt(n)*coefMortar[k]/RENthick*shellthickRec+(np.sqrt(n+t+1)-np.sqrt(n))*coefMortar[k]
                depth2MasRec = np.sqrt(n)*coefMortar[k]/MASthick*shellthickRec+(np.sqrt(n+t+1)-np.sqrt(n))*coefMortar[k]
                depth2MaiRec = np.sqrt(n)*coefMortar[k]/MAIthick*shellthickRec+(np.sqrt(n+t+1)-np.sqrt(n))*coefMortar[k]

                depth2RenBur = np.sqrt(n)*coefMortar[k]/RENthick*shellthickBur+(np.sqrt(n+store_t[k])-np.sqrt(n))*coefMortar[k]+\
                    (np.sqrt(n+t+1)-np.sqrt(n+store_t[k]))*coefMortar[k]
                depth2MasBur = np.sqrt(n)*coefMortar[k]/MASthick*shellthickBur+(np.sqrt(n+store_t[k])-np.sqrt(n))*coefMortar[k]+\
                    (np.sqrt(n+t+1)-np.sqrt(n+store_t[k]))*coefMortar[k]
                depth2MaiBur = np.sqrt(n)*coefMortar[k]/MAIthick*shellthickBur+(np.sqrt(n+store_t[k])-np.sqrt(n))*coefMortar[k]+\
                    (np.sqrt(n+t+1)-np.sqrt(n+store_t[k]))*coefMortar[k]

                depth2C15Rec = np.sqrt(n)*coefC15[k]/wallthick*shellthickRec+(np.sqrt(n+t+1)  -np.sqrt(n))*coefC15[k]
                depth2C16Rec = np.sqrt(n)*coefC16[k]/wallthick*shellthickRec+(np.sqrt(n+t+1)  -np.sqrt(n))*coefC16[k]
                depth2C23Rec = np.sqrt(n)*coefC23[k]/wallthick*shellthickRec+(np.sqrt(n+t+1)  -np.sqrt(n))*coefC23[k]
                depth2C35Rec = np.sqrt(n)*coefC35[k]/wallthick*shellthickRec+(np.sqrt(n+t+1)  -np.sqrt(n))*coefC35[k]

                depth2C15Bur = np.sqrt(n)*coefC15[k]/wallthick*shellthickBur+(np.sqrt(n+store_t[k])-np.sqrt(n))*coefC15[k]+\
                    (np.sqrt(n+t+1)-np.sqrt(n+store_t[k]))*coefC15B[k]
                depth2C16Bur = np.sqrt(n)*coefC16[k]/wallthick*shellthickBur+(np.sqrt(n+store_t[k])-np.sqrt(n))*coefC16[k]+\
                    (np.sqrt(n+t+1)-np.sqrt(n+store_t[k]))*coefC16B[k]
                depth2C23Bur = np.sqrt(n)*coefC23[k]/wallthick*shellthickBur+(np.sqrt(n+store_t[k])-np.sqrt(n))*coefC23[k]+\
                    (np.sqrt(n+t+1)-np.sqrt(n+store_t[k]))*coefC23B[k]
                depth2C35Bur = np.sqrt(n)*coefC35[k]/wallthick*shellthickBur+(np.sqrt(n+store_t[k])-np.sqrt(n))*coefC35[k]+\
                    (np.sqrt(n+t+1)-np.sqrt(n+store_t[k]))*coefC35B[k]

                CO2RenRec = ((radiusRecB-depth1RenRec)**3-(radiusRecB-depth2RenRec)**3)/(radiusRecB**3-radiusRecA**3)*adjustM[k]
                CO2MasRec = ((radiusRecB-depth1MasRec)**3-(radiusRecB-depth2MasRec)**3)/(radiusRecB**3-radiusRecA**3)*adjustM[k]
                CO2MaiRec = ((radiusRecB-depth1MaiRec)**3-(radiusRecB-depth2MaiRec)**3)/(radiusRecB**3-radiusRecA**3)*adjustM[k]
                CO2RenBur = ((radiusBurB-depth1RenBur)**3-(radiusBurB-depth2RenBur)**3)/(radiusBurB**3-radiusBurA**3)*adjustM[k]
                CO2MasBur = ((radiusBurB-depth1MasBur)**3-(radiusBurB-depth2MasBur)**3)/(radiusBurB**3-radiusBurA**3)*adjustM[k]
                CO2MaiBur = ((radiusBurB-depth1MaiBur)**3-(radiusBurB-depth2MaiBur)**3)/(radiusBurB**3-radiusBurA**3)*adjustM[k]

                CO2C15Rec = ((radiusRecB-depth1C15Rec)**3-(radiusRecB-depth2C15Rec)**3)/(radiusRecB**3-radiusRecA**3)*adjustC[k]
                CO2C16Rec = ((radiusRecB-depth1C16Rec)**3-(radiusRecB-depth2C16Rec)**3)/(radiusRecB**3-radiusRecA**3)*adjustC[k]
                CO2C23Rec = ((radiusRecB-depth1C23Rec)**3-(radiusRecB-depth2C23Rec)**3)/(radiusRecB**3-radiusRecA**3)*adjustC[k]
                CO2C35Rec = ((radiusRecB-depth1C35Rec)**3-(radiusRecB-depth2C35Rec)**3)/(radiusRecB**3-radiusRecA**3)*adjustC[k]
                CO2C15Bur = ((radiusBurB-depth1C15Bur)**3-(radiusBurB-depth2C15Bur)**3)/(radiusBurB**3-radiusBurA**3)*adjustC[k]
                CO2C16Bur = ((radiusBurB-depth1C16Bur)**3-(radiusBurB-depth2C16Bur)**3)/(radiusBurB**3-radiusBurA**3)*adjustC[k]
                CO2C23Bur = ((radiusBurB-depth1C23Bur)**3-(radiusBurB-depth2C23Bur)**3)/(radiusBurB**3-radiusBurA**3)*adjustC[k]
                CO2C35Bur = ((radiusBurB-depth1C35Bur)**3-(radiusBurB-depth2C35Bur)**3)/(radiusBurB**3-radiusBurA**3)*adjustC[k]

                if depth1RenRec >= shellthickRec:
                    CO2RenRec = 0
                if depth1MasRec >= shellthickRec:
                    CO2MasRec = 0
                if depth1MaiRec >= shellthickRec:
                    CO2MaiRec = 0
                if depth1RenBur >= shellthickBur:
                    CO2RenBur = 0
                if depth1MasBur >= shellthickBur:
                    CO2MasBur = 0
                if depth1MaiBur >= shellthickBur:
                    CO2MaiBur = 0
                
                if depth1RenRec < shellthickRec and depth2RenRec >= shellthickRec:
                    CO2RenRec = ((radiusRecB-depth1RenRec)**3-radiusRecA**3)/(radiusRecB**3-radiusRecA**3)*adjustM[k]
                if depth1MasRec < shellthickRec and depth2MasRec >= shellthickRec:
                    CO2MasRec = ((radiusRecB-depth1MasRec)**3-radiusRecA**3)/(radiusRecB**3-radiusRecA**3)*adjustM[k]
                if depth1MaiRec < shellthickRec and depth2MaiRec >= shellthickRec:
                    CO2MaiRec = ((radiusRecB-depth1MaiRec)**3-radiusRecA**3)/(radiusRecB**3-radiusRecA**3)*adjustM[k]
                if depth1RenBur < shellthickBur and depth2RenBur >= shellthickBur:
                    CO2RenBur = ((radiusBurB-depth1RenBur)**3-radiusBurA**3)/(radiusBurB**3-radiusBurA**3)*adjustM[k]
                if depth1MasBur < shellthickBur and depth2MasBur >= shellthickBur:
                    CO2MasBur = ((radiusBurB-depth1MasBur)**3-radiusBurA**3)/(radiusBurB**3-radiusBurA**3)*adjustM[k]
                if depth1MaiBur < shellthickBur and depth2MaiBur >= shellthickBur:
                    CO2MaiBur = ((radiusBurB-depth1MaiBur)**3-radiusBurA**3)/(radiusBurB**3-radiusBurA**3)*adjustM[k]

                if depth1C15Rec >= shellthickRec:
                    CO2C15Rec = 0
                if depth1C16Rec >= shellthickRec:
                    CO2C16Rec = 0
                if depth1C23Rec >= shellthickRec:
                    CO2C23Rec = 0
                if depth1C35Rec >= shellthickRec:
                    CO2C35Rec = 0
                if depth1C15Bur >= shellthickBur:
                    CO2C15Bur = 0
                if depth1C16Bur >= shellthickBur:
                    CO2C16Bur = 0
                if depth1C23Bur >= shellthickBur:
                    CO2C23Bur = 0
                if depth1C35Bur >= shellthickBur:
                    CO2C35Bur = 0

                if depth1C15Rec < shellthickRec and depth2C15Rec >= shellthickRec:
                    CO2C15Rec = ((radiusRecB-depth1C15Rec)**3-radiusRecA**3)/(radiusRecB**3-radiusRecA**3)*adjustC[k]
                if depth1C16Rec < shellthickRec and depth2C16Rec >= shellthickRec:
                    CO2C16Rec = ((radiusRecB-depth1C16Rec)**3-radiusRecA**3)/(radiusRecB**3-radiusRecA**3)*adjustC[k]
                if depth1C23Rec < shellthickRec and depth2C23Rec >= shellthickRec:
                    CO2C23Rec = ((radiusRecB-depth1C23Rec)**3-radiusRecA**3)/(radiusRecB**3-radiusRecA**3)*adjustC[k]
                if depth1C35Rec < shellthickRec and depth2C35Rec >= shellthickRec:
                    CO2C35Rec = ((radiusRecB-depth1C35Rec)**3-radiusRecA**3)/(radiusRecB**3-radiusRecA**3)*adjustC[k]
                if depth1C15Bur < shellthickBur and depth2C15Bur >= shellthickBur:
                    CO2C15Bur = ((radiusBurB-depth1C15Bur)**3-radiusBurA**3)/(radiusBurB**3-radiusBurA**3)*adjustC[k]
                if depth1C16Bur < shellthickBur and depth2C16Bur >= shellthickBur:
                    CO2C16Bur = ((radiusBurB-depth1C16Bur)**3-radiusBurA**3)/(radiusBurB**3-radiusBurA**3)*adjustC[k]
                if depth1C23Bur < shellthickBur and depth2C23Bur >= shellthickBur:
                    CO2C23Bur = ((radiusBurB-depth1C23Bur)**3-radiusBurA**3)/(radiusBurB**3-radiusBurA**3)*adjustC[k]
                if depth1C35Bur < shellthickBur and depth2C35Bur >= shellthickBur:
                    CO2C35Bur = ((radiusBurB-depth1C35Bur)**3-radiusBurA**3)/(radiusBurB**3-radiusBurA**3)*adjustC[k]
                
                uptake_mortar_Ren_Rec_Matrix[k+n,k+t+n] = cem_app_individual[k]*RENpercentage*demolish_rate*recRate[k]*CO2RenRec
                uptake_mortar_Mas_Rec_Matrix[k+n,k+t+n] = cem_app_individual[k]*MASpercentage*demolish_rate*recRate[k]*CO2MasRec
                uptake_mortar_Mai_Rec_Matrix[k+n,k+t+n] = cem_app_individual[k]*MAIpercentage*demolish_rate*recRate[k]*CO2MaiRec

                uptake_mortar_Ren_Bur_Matrix[k+n,k+t+n] = cem_app_individual[k]*RENpercentage*demolish_rate*burRate[k]*CO2RenBur
                uptake_mortar_Mas_Bur_Matrix[k+n,k+t+n] = cem_app_individual[k]*MASpercentage*demolish_rate*burRate[k]*CO2MasBur
                uptake_mortar_Mai_Bur_Matrix[k+n,k+t+n] = cem_app_individual[k]*MAIpercentage*demolish_rate*burRate[k]*CO2MaiBur

                uptake_concrete_C15_Rec_Matrix[k+n,k+t+n] = cem_app_individual[k]*C15percentage*demolish_rate*recRate[k]*CO2C15Rec
                uptake_concrete_C16_Rec_Matrix[k+n,k+t+n] = cem_app_individual[k]*C16percentage*demolish_rate*recRate[k]*CO2C16Rec
                uptake_concrete_C23_Rec_Matrix[k+n,k+t+n] = cem_app_individual[k]*C23percentage*demolish_rate*recRate[k]*CO2C23Rec
                uptake_concrete_C35_Rec_Matrix[k+n,k+t+n] = cem_app_individual[k]*C35percentage*demolish_rate*recRate[k]*CO2C35Rec
                
                uptake_concrete_C15_Bur_Matrix[k+n,k+t+n] = cem_app_individual[k]*C15percentage*demolish_rate*burRate[k]*CO2C15Bur
                uptake_concrete_C16_Bur_Matrix[k+n,k+t+n] = cem_app_individual[k]*C16percentage*demolish_rate*burRate[k]*CO2C16Bur
                uptake_concrete_C23_Bur_Matrix[k+n,k+t+n] = cem_app_individual[k]*C23percentage*demolish_rate*burRate[k]*CO2C23Bur
                uptake_concrete_C35_Bur_Matrix[k+n,k+t+n] = cem_app_individual[k]*C35percentage*demolish_rate*burRate[k]*CO2C35Bur

        uptake_concrete_demolish_Matrix[k,:] =\
            pd.Series(np.sum(
                uptake_concrete_C15_Rec_Matrix + uptake_concrete_C16_Rec_Matrix +\
                    uptake_concrete_C23_Rec_Matrix + uptake_concrete_C35_Rec_Matrix +\
                    uptake_concrete_C15_Bur_Matrix + uptake_concrete_C16_Bur_Matrix +\
                    uptake_concrete_C23_Bur_Matrix + uptake_concrete_C35_Bur_Matrix,
                    axis = 0)
                    )
        
        uptake_mortar_demolish_Matrix[k,:] =\
            pd.Series(np.sum(
                    uptake_mortar_Ren_Rec_Matrix + uptake_mortar_Mas_Rec_Matrix + uptake_mortar_Mai_Rec_Matrix +\
                    uptake_mortar_Ren_Bur_Matrix + uptake_mortar_Mas_Bur_Matrix + uptake_mortar_Mai_Bur_Matrix,
                    axis = 0)
                    )

    pd.DataFrame(uptake_concrete_demolish_Matrix).to_excel(
        'C:\\Users\\caozh\\Documents\\2020-ClimateWork\\Carbonation\\' + filepath + '\\demolishConcrete_' + str(scenario_index+"_"+storyline) + '.xlsx')

    pd.DataFrame(uptake_mortar_demolish_Matrix).to_excel(
        'C:\\Users\\caozh\\Documents\\2020-ClimateWork\\Carbonation\\' + filepath + '\\demolishMortar_' + str(scenario_index+"_"+storyline) + '.xlsx')

    upatake_concrete_use =\
        uptake_concrete_C15_use_Matrix + uptake_concrete_C16_use_Matrix +\
            uptake_concrete_C23_use_Matrix + uptake_concrete_C35_use_Matrix

    pd.DataFrame(upatake_concrete_use).to_excel(
        'C:\\Users\\caozh\\Documents\\2020-ClimateWork\\Carbonation\\' + filepath + '\\useConcrete_' + str(scenario_index+"_"+storyline) + '.xlsx')

    uptake_mortar_use =\
    uptake_mortar_Ren_use_Matrix + uptake_mortar_Mas_use_Matrix + uptake_mortar_Mai_use_Matrix

    pd.DataFrame(uptake_mortar_use).to_excel(
        'C:\\Users\\caozh\\Documents\\2020-ClimateWork\\Carbonation\\' + filepath + '\\useMortar_' + str(scenario_index+"_"+storyline) + '.xlsx')

    uptake = np.sum(uptake_wasteConcrete_Matrix + uptake_wasteMortar_Matrix+\
        uptake_concrete_demolish_Matrix + uptake_mortar_demolish_Matrix+\
            upatake_concrete_use + uptake_mortar_use,axis=0) + uptake_CKD

    pd.DataFrame(uptake).to_excel(
        'C:\\Users\\caozh\\Documents\\2020-ClimateWork\\Carbonation\\' + filepath + '\\uptake_' + str(scenario_index+"_"+storyline) + '.xlsx')
