if __name__ == '__main__':
    from ema_workbench import (RealParameter, IntegerParameter, ScalarOutcome, Constant,
                           Model, Policy, MultiprocessingEvaluator, ArrayOutcome)
    from visum_connector import VisumModel
    ema_logging.log_to_stderr(level=ema_logging.DEBUG)

    n=0
    
    model = VisumModel('testmodel', wd=wd, model_file='Policy base.ver',
                           n_replications=1)
    
    # values taken from data of Dick Bakker (via outlook)
    model.uncertainties = [RealParameter('EBIKE_BASIS', 0.22, 0.28),
                          RealParameter('EBIKE_OW', 0.09, 0.11),
                          RealParameter('THUISWERKREDUCTIE', 0.95, 1),
                          RealParameter('KMKOSTENINDEX', 0.7 ,0.953),
                          RealParameter('OVKOSTENINDEX', 0.9, 1.1)]

    model.outcomes = [ScalarOutcome('carshare_verpl'),
                      ScalarOutcome('bikeshare_verpl'),
                      ScalarOutcome('OVshare_verpl'),
                      ScalarOutcome('totaal_verpl'),
                      ScalarOutcome("carshare_km"),
                      ScalarOutcome("OVshare_km"),
                      ScalarOutcome("bikeshare_km"),
                      ScalarOutcome("totaal_km")
                     ]

    policies = [Policy('basecase',
                      model_file = 'Policy base.ver', output_file = 'Grovem_TAB_GroVem_ref_1a1.xlsx'),
                Policy('fiets',
                      model_file='Policy fiets.ver', output_file = 'Grovem_TAB_GroVem_ref_1a1.xlsx'),
                Policy('30 km',
                       model_file='Policy 30 km.ver', output_file = 'Grovem_TAB_GroVem_ref_1a1.xlsx'),
                Policy('parkeren',
                       model_file='Policy parkeren.ver', output_file = 'Grovem_TAB_GroVem_ref_1a1.xlsx')
                    ]


    with MultiprocessingEvaluator(model, n_processes=5) as evaluator:
        results = evaluator.perform_experiments(250, policies=policies,  uncertainty_sampling="lhs")

