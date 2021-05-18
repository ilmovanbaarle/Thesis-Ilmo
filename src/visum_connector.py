import os
import win32com.client

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
import time

from ema_workbench.util import CaseError, EMAError, EMAWarning, get_module_logger
from ema_workbench.util.ema_logging import method_logger

from logging import DEBUG as debug

_logger = get_module_logger(__name__)


class VisumModel(SingleReplication, FileModel):

    def __init__(self, name, wd=None, model_file=None, output_file=None,
                 n_replications=1):
        super(VisumModel, self).__init__(name, wd=wd, model_file=model_file)
        self.Visum = None
        self.model_loaded = False
        self.n_replications = n_replications
        self.output_file = output_file
        self.path_to_output_file = None

    @method_logger(__name__)
    def model_init(self, policy):
        #time.sleep(10)
        super(VisumModel, self).model_init(policy)
        
        if self.Visum is None:
            _logger.debug("trying to start visum")
            self.Visum = win32com.client.Dispatch("Visum.Visum")
            
        path_to_model = os.path.join(self.working_directory, self.model_file)
        self.path_to_output_file = os.path.join(self.working_directory,
                                                self.output_file)

        if ("model_file" in policy) or (not self.model_loaded): 
            _logger.debug(f"trying to load visum model from {path_to_model}")
            self.Visum.LoadVersion(path_to_model)
            _logger.debug("model loaded succesfully")
            self.model_loaded = True



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

    df1 = pd.read_csv(df)
    only_gemeente = df1[(df1.O == 1) | (df1.D == 1)]

    car_km = only_gemeente[only_gemeente.VVW == 'AB'].Value.sum()
    OV_km = only_gemeente[only_gemeente.VVW == 'OV'].Value.sum()
    fiets_km = only_gemeente[only_gemeente.VVW == 'FTS'].Value.sum()
    totaal_gemeente = only_gemeente.Value.sum()

    carshare_gemeente = car_km/totaal_gemeente
    OVshare_gemeente = OV_km/totaal_gemeente
    bikeshare_gemeente = fiets_km/totaal_gemeente
    
    output = {"carshare_gemeente":carshare_gemeente, "bikeshare_gemeente":bikeshare_gemeente,
             "OVshare_gemeente":OVshare_gemeente,"totaal_gemeente":totaal_gemeente}

    return output

    
if __name__ == '__main__':
    from ema_workbench import (RealParameter, IntegerParameter, ScalarOutcome, Constant,
                           Model, Policy, MultiprocessingEvaluator, ArrayOutcome)
    ema_logging.log_to_stderr(level=ema_logging.DEBUG)
    
    wd = "./20201221 Mobilitievisie Groningen - complete/"

    model = VisumModel('testmodel', wd=wd, model_file='Version for Thesis good.ver', output_file='TAB_GroVem_2040H_Iter1.csv',
                           n_replications=1)
    
    # values taken from data of Dick Bakker (via outlook)
    model.uncertainties = [RealParameter('EBIKE_BASIS', 0.22, 0.28),
                          RealParameter('EBIKE_OW', 0.09, 0.11),
                          RealParameter('THUISWERKREDUCTIE', 0.95, 1),
                          RealParameter('KMKOSTENINDEX', 0.7 ,0.953),
                          RealParameter('OVKOSTENINDEX', 0.9, 1.1)]

    model.outcomes = [ScalarOutcome('carshare_gemeente'),
                      ScalarOutcome('bikeshare_gemeente'),
                      ScalarOutcome('OVshare_gemeente'),
                      ScalarOutcome('totaal_gemeente')]

    policies = [Policy('basecase',
                      model_file = 'Version for Thesis good.ver', output_file = 'Basecase.csv'),
                Policy('fiets',
                      model_file='Policy fiets.ver', output_file = 'fiets.csv'),
                Policy('30 km',
                       model_file='Policy 30 km.ver', output_file = '30_km.csv'),
                Policy('knips hoog',
                       model_file='Policy knips hoog.ver', output_file = 'knips_hoog.csv')
                    ]


    #results = perform_experiments(scenarios=2, models = model, policies = policies)
    #maxtasksperchild=4
    with MultiprocessingEvaluator(model, n_processes=5) as evaluator:
        results = evaluator.perform_experiments(4, policies=policies)
