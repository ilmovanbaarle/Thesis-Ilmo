
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
    ema_logging.log_to_stderr(level=ema_logging.DEBUG)

    model = VisumModel('testmodel', wd='', model_file='', output_file='',
                       n_replications=1)
    model.uncertaintes = []
    model.outcomes = []

    results = perform_experiments(scenarios=10)
