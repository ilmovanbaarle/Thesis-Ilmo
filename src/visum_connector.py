
import os
import win32com

from ema_workbench.em_framework import FileModel, SingleReplication
from ema_workbench.util import (method_logger, get_module_logger, debug,
                                info)
from ema_workbenc import ema_logging, perform_experiments

_logger = get_module_logger(__name__)


class VisumModel(SingleReplication, FileModel):

    def __init__(self, name, wd=None, model_file=None, output_file=None,
                 n_replications=1):
        super(VisumModel, self).__init__(name, wd=wd, model_file=model_file)
        self.visum = None
        self.n_replications = n_replications
        self.output_file = output_file
        self.path_to_output_file = None

    @method_logger(__name__)
    def model_init(self, policy):
        super(VisumModel, self).model_init(policy)
        path_to_model = os.path.join(self.working_directory, self.model_file)
        self.path_to_output_file = os.path.join(self.working_directory,
                                                self.output_file)

        debug("trying to start visum")
        self.visum = win32com.client.Dispatch("Visum.Visum")

        debug(f"trying to load visum model from {path_to_model}")
        self.visum.LoadVersion(path_to_model)
        debug("model loaded succesfully")

        # start visum

    @method_logger(__name__)
    def run_experiment(self, experiment):

        for k, v in experiment.items():
            Visum.Net.SetAttValue(k, v)

        # run experiment
        for i in range(self.n_replications):
            # in order to run te model
            Visum.Procedures.Execute()

        # return outputs
        output = pd.read_csv(self.path_to_output_file)

        results = {}
        for variable in self.output_variables:

            # assumption is that outcomes are stored by column
            # not sure if this is correct
            results[variable] = output.loc[:, variable]

        return results


if __name__ == '__main__':
    from ema_workbench import (RealParameter, IntegerParameter, ScalarOutcome, Constant,
                           Model, Policy)
    ema_logging.log_to_stderr(level=ema_logging.DEBUG)

    model = VisumModel('testmodel', wd=wd, model_file='Version for Thesis good.ver', output_file='TAB_GroVem_2040H_Iter1.csv',
                           n_replications=1)
    model.uncertainties = [RealParameter('EBIKE_BASIS', 0.2, 0.5),
                          RealParameter('EBIKE_OW', 0.08, 0.25),
                          RealParameter('THUISWERKREDUCTIE', 0.8, 0.99),
                          RealParameter('KMKOSTENINDEX', 0.5 ,0.9),
                          RealParameter('OVKOSTENINDEX', 0.9, 1.1)]

    model.outcomes = [ScalarOutcome('car_share'),
                      ScalarOutcome('OV_share'),
                      ScalarOutcome('total_km')]

    policies = [Policy('basecase',
                      model_file = 'Version for Thesis good.ver'),
                    Policy('fiets',
                           model_file='Policy fiets.ver'),
                    Policy('30 km',
                           model_file='Policy 30 km.ver'),
                    Policy('knips hoog',
                           model_file='Policy knips hoog.ver')
                    ]


    results = perform_experiments(scenarios=1, models = model, policies = policies)
    
    
    ##########
    
[MainProcess/DEBUG] setting working directory to C:\Users\nlilbm\Documents\Thesis\20201221 Mobilitievisie Groningen - complete
[MainProcess/INFO] performing 1 scenarios * 4 policies * 1 model(s) = 4 experiments
[MainProcess/INFO] performing experiments sequentially
[MainProcess/DEBUG] running scenario 16 for policy basecase on model testmodel
[MainProcess/DEBUG] calling run_model on SingleReplication
[MainProcess/DEBUG] calling run_model on AbstractModel
[MainProcess/DEBUG] calling initialized on AbstractModel
[MainProcess/DEBUG] completed calling initialized on AbstractModel
[MainProcess/DEBUG] calling model_init on VisumModel
[MainProcess/DEBUG] calling model_init on AbstractModel
[MainProcess/DEBUG] completed calling model_init on AbstractModel
[MainProcess/DEBUG] completed calling model_init on VisumModel
[MainProcess/DEBUG] calling _transform on AbstractModel
[MainProcess/DEBUG] completed calling _transform on AbstractModel
[MainProcess/DEBUG] calling _transform on AbstractModel
[MainProcess/DEBUG] completed calling _transform on AbstractModel
[MainProcess/DEBUG] completed calling run_model on AbstractModel
[MainProcess/DEBUG] calling run_experiment on VisumModel
[MainProcess/ERROR] name 'Visum' is not defined
Traceback (most recent call last):
  File "C:\Users\nlilbm\Anaconda3\lib\site-packages\ema_workbench\em_framework\experiment_runner.py", line 85, in run_experiment
    model.run_model(scenario, policy)
  File "C:\Users\nlilbm\Anaconda3\lib\site-packages\ema_workbench\util\ema_logging.py", line 158, in wrapper
    res = func(*args, **kwargs)
  File "C:\Users\nlilbm\Anaconda3\lib\site-packages\ema_workbench\em_framework\model.py", line 338, in run_model
    outputs = self.run_experiment(experiment)
  File "C:\Users\nlilbm\Anaconda3\lib\site-packages\ema_workbench\util\ema_logging.py", line 158, in wrapper
    res = func(*args, **kwargs)
  File "<ipython-input-40-1a98a2fcc369>", line 56, in run_experiment
    Visum.Net.SetAttValue(k, v)
NameError: name 'Visum' is not defined
[MainProcess/DEBUG] calling cleanup on AbstractModel
[MainProcess/DEBUG] completed calling cleanup on AbstractModel
