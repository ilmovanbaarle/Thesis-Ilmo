#!/usr/bin/env python
# coding: utf-8

# In[1]:


import win32com.client
import pandas as pd
import numpy as np
import math

import seaborn as sns
import matplotlib.pyplot as plt

from scipy.optimize import brentq

wd = "C:/Users/nlilbm/Documents/Thesis/20201221 Mobilitievisie Groningen - complete"


# In[2]:


Visum = win32com.client.Dispatch("Visum.Visum")
Visum.LoadVersion(wd + "/Version for Thesis good.ver")


# In[6]:


def small_model(numlanes=1,v0=50,nsamples=5):
    
    link2_vol = 0
    link3_vol = 0
    link5_vol = 0
    link10_vol = 0
    
    Visum.Net.Links.ItemByKey(3,2).SetAttValue("V0PRT",v0)
    Visum.Net.Links.ItemByKey(5,8).SetAttValue("V0PRT",v0)
    Visum.Net.Links.ItemByKey(3,2).SetAttValue("NUMLANES",numlanes)
    Visum.Net.Links.ItemByKey(5,8).SetAttValue("NUMLANES",numlanes)
    
    for i in range(nsamples):
        myvisum.Procedures.Execute()

        link2_vol += Visum.Net.Links.ItemByKey(3,2).AttValue("VOLVEHPRT(AP)")
        link3_vol = Visum.Net.Links.ItemByKey(2,5).AttValue("VOLVEHPRT(AP)")
        link5_vol = Visum.Net.Links.ItemByKey(5,8).AttValue("VOLVEHPRT(AP)")
        link10_vol = Visum.Net.Links.ItemByKey(3,4).AttValue("VOLVEHPRT(AP)")
    
    link2_vol = link2_vol / nsamples
    link3_vol = link3_vol / nsamples
    link5_vol = link5_vol / nsamples
    link10_vol = link10_vol / nsamples
    
    return link10_vol, link5_vol, link3_vol, link2_vol


# In[7]:


def scenario_instellingen(KMKOSTEN=0.7,OVKOSTEN=1.028,EBIKE_BS=0.28,EBIKE_OW=0.11,THUISW=0.95):

    # naam uitvoer
    Visum.Net.SetAttValue("SCENARIO","GroVem_2040H")

    # WLO scenario: 2030L 2030H 2040L 2040H (andere waarden niet toegestaan
    Visum.Net.SetAttValue("WLO","2040H")

    # Parameters
    # KMKOSTEN=Visum.Net.AttValue("KMKOSTENINDEX")
    # OVKOSTEN=Visum.Net.AttValue("OVKOSTENINDEX")
    # EBIKE_BS=Visum.Net.AttValue("EBIKE_BASIS")
    # EBIKE_OW=Visum.Net.AttValue("EBIKE_OW")
    # THUISW=Visum.Net.AttValue("THUISWERKREDUCTIE")



    #=============================================== Hieronder Niet wijzigen =================================
    # SCENARIO INSTELLINGEN
    Visum.Net.SetAttValue("INDEXOVKOSTEN",OVKOSTEN)

    Visum.Net.Activities.SetMultiAttValues("IndexAutoKosten",((1,KMKOSTEN),(2,KMKOSTEN),(3,KMKOSTEN),(4,KMKOSTEN),(5,KMKOSTEN),(6,KMKOSTEN),(7,KMKOSTEN),(8,KMKOSTEN),(9,KMKOSTEN),(10,KMKOSTEN),(11,KMKOSTEN),(12,KMKOSTEN),(13,KMKOSTEN),(14,KMKOSTEN)),Add=False)# Autokosten via Add1 toedeling
    # Eerste iteratie gaat uit van Basisjaar LOS
    # Daarom hier een algemene kosten index voor de eerste iteratie
    Visum.Net.SetAttValue("IndexAutoKostBasis",KMKOSTEN)

    # Thuiswerken
    Visum.Net.SetAttValue("THUIS_FT",THUISW)
    Visum.Net.SetAttValue("THUIS_PT",THUISW)
    Visum.Net.SetAttValue("THUIS_STUD",THUISW)
    Visum.Net.SetAttValue("THUIS_PENS",THUISW)
    Visum.Net.SetAttValue("THUIS_OVER",THUISW)

    # % Ebike
    Visum.Net.Activities.SetMultiAttValues("Ebike_frac",((1,EBIKE_BS),(2,EBIKE_BS),(3,EBIKE_OW),(4,EBIKE_BS),(5,EBIKE_BS),(6,EBIKE_BS),(7,EBIKE_OW),(8,EBIKE_BS),(9,EBIKE_BS),(10,EBIKE_BS),(11,EBIKE_BS),(12,0.0),(13,0.0),(14,0.0)),Add=False)


# In[8]:


def bigmodel(KMKOSTEN,OVKOSTEN,EBIKE_BS,EBIKE_OW,THUISW,nsamples=1):
    scenario_instellingen(KMKOSTEN,OVKOSTEN,EBIKE_BS,EBIKE_OW,THUISW)
   
    for i in range(nsamples):
        Visum.Procedures.Execute()
        
    TAB_GroVem_2040H_Iter1.csv
    df1 = pd.read_csv('20201221 Mobilitievisie Groningen - complete/TAB_GroVem_2040H_Iter1.csv')
    df2 = pd.read_csv('20201221 Mobilitievisie Groningen - complete/TAB_all_GroVem_2040H_Iter1.csv')
    df3 = pd.read_csv('20201221 Mobilitievisie Groningen - complete/TAB_km_GroVem_2040H_Iter1.csv', header =None)
    
    car_km = df1[df1.VVW == 'AB'].Value.sum()
    OV_km = df1[df1.VVW == 'OV'].Value.sum()
    total_km = df1.Value.sum()
    
    car_share = car_km/total_km
    OV_share = OV_km/total_km
    
    return car_share, OV_share, total_km


# df1 = pd.read_csv('20201221 Mobilitievisie Groningen - complete/TAB_GroVem_2040H_Iter1.csv')
# df1 = df1[df1.O == 1].Value.sum()
# df1

# xls = pd.ExcelFile('20201221 Mobilitievisie Groningen - complete/Grovem_TAB_GroVem_2040H1.xlsx')
# df1 = pd.read_excel(xls, 'SYNT_VERPL', index_col = 0, header = None, skiprows = 0,
#                 nrows = 1000, usecols = 'B:H')
# df1.iloc[54,5]
# df1.loc['TAB_V45_BASIS',2]

# In[13]:


from ema_workbench import (RealParameter, IntegerParameter, ScalarOutcome, Constant,
                           Model)

model = Model('bigmodel', function=bigmodel)

#specify uncertainties
model.uncertainties = [RealParameter('EBIKE_BS', 0.2, 0.5),
                      RealParameter('EBIKE_OW', 0.08, 0.25),
                      RealParameter('THUISW', 0.8, 0.99)]

# set levers
model.levers = [RealParameter('KMKOSTEN', 0.5 ,0.9),
               RealParameter('OVKOSTEN', 0.9, 1.1)]

#specify outcomes
model.outcomes = [ScalarOutcome('car_share'),
                  ScalarOutcome('OV_share'),
                  ScalarOutcome('total_km'),]

# override some of the defaults of the model
model.constants = [Constant('nsamples', 1)]


# In[37]:


from ema_workbench import (MultiprocessingEvaluator, SequentialEvaluator, ema_logging,
                           perform_experiments)
ema_logging.log_to_stderr(ema_logging.INFO)

with SequentialEvaluator(model) as evaluator:
    results = evaluator.perform_experiments(scenarios=4, policies=4)


# In[40]:


policies = experiments['policy']
for i, policy in enumerate(np.unique(policies)):
    experiments.loc[policies==policy, 'policy'] = str(i)

data = pd.DataFrame(outcomes)
data['policy'] = policies


# In[41]:


sns.pairplot(data, hue='policy', vars=list(outcomes.keys()))
plt.show()


# In[42]:


from ema_workbench.analysis import pairs_plotting

fig, axes = pairs_plotting.pairs_scatter(experiments, outcomes, group_by='policy',
                                         legend=False)
fig.set_size_inches(8,8)
plt.show()


# In[43]:


from ema_workbench import save_results
save_results(results, '16 scenarios 4 policies.tar.gz')

from ema_workbench import load_results
results = load_results('16 scenarios 4 policies.tar.gz')


# In[45]:


from ema_workbench.analysis import prim

x = experiments
y = outcomes['car_share'] < 0.49
prim_alg = prim.Prim(x, y, threshold=0.8)
box1 = prim_alg.find_box()


# In[46]:


box1.show_tradeoff()
plt.show()


# In[50]:


box1.inspect(2)
box1.inspect(2, style='graph')
plt.show()


# In[52]:


box1.show_pairs_scatter(2)
plt.show()


# In[53]:


from ema_workbench.analysis import feature_scoring

x = experiments
y = outcomes

fs = feature_scoring.get_feature_scores_all(x, y)
sns.heatmap(fs, cmap='viridis', annot=True)
plt.show()


# In[76]:


from ema_workbench.analysis import dimensional_stacking

x = experiments
y = outcomes['car_share'] < 0.49
dimensional_stacking.create_pivot_plot(x,y, 1, nbins=3)
plt.show()


# In[66]:


from ema_workbench import (RealParameter, IntegerParameter, ScalarOutcome, Constant,
                           Model)

model = Model('bigmodel', function=big_model)

#specify uncertainties
model.uncertainties = [RealParameter('EBIKE_BS', 0.2, 0.5),
                      RealParameter('EBIKE_OW', 0.08, 0.25),
                      RealParameter('THUISW', 0.8, 0.99)]

# set levers
model.levers = [RealParameter('KMKOSTEN', 0.5 ,0.9),
               RealParameter('OVKOSTEN', 0.9, 1.1),
               RealParameter('EBIKE_BS', 0.2, 0.5),
                RealParameter('EBIKE_OW', 0.08, 0.25),
                RealParameter('THUISW', 0.8, 0.99)]

#specify outcomes
model.outcomes = [ScalarOutcome('car_share',ScalarOutcome.MINIMIZE),
                  ScalarOutcome('OV_share',ScalarOutcome.MAXIMIZE),
                  ScalarOutcome('total_km',ScalarOutcome.MAXIMIZE),]

model.constants = [Constant('nsamples', 1)]


# In[67]:


from ema_workbench import MultiprocessingEvaluator, ema_logging

ema_logging.log_to_stderr(ema_logging.INFO)

with SequentialEvaluator(model) as evaluator:
    results = evaluator.optimize(nfe=250, searchover='levers',
                                 epsilons=[0.1,]*len(model.outcomes))


# In[ ]:




