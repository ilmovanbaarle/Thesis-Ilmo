wd = "C:/Users/nlilbm/Documents/Thesis/20201221 Mobilitievisie Groningen - complete/"

def output_changer(df):
    xls = pd.ExcelFile(df)
    df1 = pd.read_excel(xls, 'SYNT_VERPL', index_col = 0, header = None, skiprows = 0,
                    nrows = 1000, usecols = 'B:H')
    
    carshare_gemeente = df1.loc["TOTAAL"].iloc[1,0]/df1.loc["TOTAAL"].iloc[1,-1]
    bikeshare_gemeente = df1.loc["TOTAAL"].iloc[1,3]/df1.loc["TOTAAL"].iloc[1,-1]
    OVshare_gemeente = df1.loc["TOTAAL"].iloc[1,2]/df1.loc["TOTAAL"].iloc[1,-1]
    totaal_gemeente = df1.loc["TOTAAL"].iloc[1,-1]
    
    output = {"carshare_gemeente":carshare_gemeente, "bikeshare_gemeente":bikeshare_gemeente,
             "OVshare_gemeente":OVshare_gemeente,"totaal_gemeente":totaal_gemeente}

    return output

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
            self.Visum.Net.SetAttValue(k, v)

        # run experiment
        for i in range(self.n_replications):
            # in order to run te model
            self.Visum.Procedures.Execute()

        # return outputs
        output = output_changer(self.path_to_output_file)

        results = {}
        for variable in self.output_variables:

            # assumption is that outcomes are stored by column
            # not sure if this is correct
            if variable in results:
                tmp = results[variable]
                tmp.append(output[variable])
                results[variable] = tmp
            if variable not in results:
                results[variable] = [output[variable]]

        return results


if __name__ == '__main__':
    from ema_workbench import (RealParameter, IntegerParameter, ScalarOutcome, Constant,
                           Model, Policy, MultiprocessingEvaluator)
    ema_logging.log_to_stderr(level=ema_logging.DEBUG)

    model = VisumModel('testmodel', wd=wd, model_file='Version for Thesis good.ver', output_file='TAB_GroVem_2040H_Iter1.csv',
                           n_replications=1)
    model.uncertainties = [RealParameter('EBIKE_BASIS', 0.2, 0.5),
                          RealParameter('EBIKE_OW', 0.08, 0.25),
                          RealParameter('THUISWERKREDUCTIE', 0.8, 0.99),
                          RealParameter('KMKOSTENINDEX', 0.5 ,0.9),
                          RealParameter('OVKOSTENINDEX', 0.9, 1.1)]

    model.outcomes = [ScalarOutcome('carshare_gemeente'),
                      ScalarOutcome('bikeshare_gemeente'),
                      ScalarOutcome('OVshare_gemeente'),
                      ScalarOutcome('totaal_gemeente')]

    policies = [Policy('basecase',
                      model_file = 'Version for Thesis good.ver', output_file = 'Basecase.xlsx'),
                    Policy('fiets',
                           model_file='Policy fiets.ver', output_file = 'fiets.xlsx'),
                    Policy('30 km',
                           model_file='Policy 30 km.ver', output_file = '30 km.xlsx'),
                    Policy('knips hoog',
                           model_file='Policy knips hoog.ver', output_file = 'knips hoog.xlsx')
                    ]


    results = perform_experiments(scenarios=1, models = model, policies = policies)
    #with MultiprocessingEvaluator(model, n_processes=4) as evaluator:
     #   results = evaluator.perform_experiments(4, policies=policies)
