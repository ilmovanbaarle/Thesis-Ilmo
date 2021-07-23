import os
import win32com.client
import time

from ema_workbench import ema_logging, perform_experiments

import pandas as pd
import decimal
import math
import numpy as np
import os
import warnings

from ema_workbench.em_framework import TimeSeriesOutcome, FileModel
from ema_workbench.em_framework.parameters import (Parameter, RealParameter,
                                       CategoricalParameter)
from ema_workbench.em_framework.util import NamedObjectMap
from ema_workbench.em_framework.model import SingleReplication

from ema_workbench.util import CaseError, EMAError, EMAWarning, get_module_logger
from ema_workbench.util.ema_logging import method_logger

from logging import DEBUG as debug

_logger = get_module_logger(__name__)


class VisumModel(SingleReplication, FileModel):

    def __init__(self, name, wd=None, model_file=None, output_file=None,
                 n_replications=1):
        super(VisumModel, self).__init__(name, wd=wd, model_file=model_file)
        self.Visum = None
        self.n_replications = n_replications
        self.output_file = output_file
        self.path_to_output_file = None

    @method_logger(__name__)
    def model_init(self, policy):
        time.sleep(5)
        super(VisumModel, self).model_init(policy)
        path_to_model = os.path.join(self.working_directory, self.model_file)
        self.path_to_output_file = os.path.join(self.working_directory,
                                                self.output_file)

        #debug("trying to start visum")
        self.Visum = win32com.client.Dispatch("Visum.Visum")

        #debug(f"trying to load visum model from {path_to_model}")
        self.Visum.LoadVersion(path_to_model)
        #debug("model loaded succesfully")

        # start visum

    @method_logger(__name__)
    def run_experiment(self, experiment):

        for k, v in experiment.items():
            print(k,v)
            self.Visum.Net.SetAttValue(k, v)
            

        # run experiment
        for i in range(self.n_replications):
            # in order to run te model
            self.Visum.Procedures.Execute()

        # return outputs
        output = output_changer(self.path_to_output_file)

        results = {}
        for variable in self.output_variables:
            
            results[variable] = output[variable]

        return results


def output_changer(df):
    
    xls = pd.ExcelFile(df)
    df_verpl = pd.read_excel(xls, 'SYNT_TRIPLENGTE', index_col = 0, skiprows = 12,
                    nrows = 6, usecols = 'B:H')
    df_km = pd.read_excel(xls, 'SYNT_TRIPLENGTE', index_col = 0, skiprows = 12,
                    nrows = 6, usecols = 'J:P')

    car_verpl = df_verpl.loc["Totaal","Autobestuurder"]
    OV_verpl = df_verpl.loc["Totaal","Openbaar Vervoer"]
    fiets_verpl = df_verpl.loc["Totaal","Fiets"]
    totaal_verpl = df_verpl.loc["Totaal","Totaal"]

    carshare_verpl = car_verpl/totaal_verpl
    OVshare_verpl = OV_verpl/totaal_verpl
    bikeshare_verpl = fiets_verpl/totaal_verpl
    
    car_km = df_km.loc["Totaal","Autobestuurder.1"]
    OV_km = df_km.loc["Totaal","Openbaar Vervoer.1"]
    fiets_km = df_km.loc["Totaal","Fiets.1"] 
    totaal_km = df_km.loc["Totaal","Totaal.1"]
    
    carshare_km = car_km/totaal_km
    OVshare_km = OV_km/totaal_km
    bikeshare_km = fiets_km/totaal_km
    
    output = {"carshare_verpl":carshare_verpl, "bikeshare_verpl":bikeshare_verpl,
             "OVshare_verpl":OVshare_verpl,"totaal_verpl":totaal_verpl,
             'carshare_km':carshare_km, "OVshare_km":OVshare_km,
             "bikeshare_km":bikeshare_km, "totaal_km":totaal_km}

    return output
  
def lever_changer(levers):
    self.Visum.Net.Links.SetAllAttValues("BIKE_SPEED", BIKE_SPEED)
    
    for i in list(df3.index):
        zone = Visum.Net.Zones.ItemByKey(i)
        zone.SetAttValue("PT_KP", PT_KP)
