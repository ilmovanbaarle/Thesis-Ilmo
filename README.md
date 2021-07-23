# Taking deep uncertainty into account in traffic models: a case study of Groningen - Thesis Ilmo van Baarle

This is a readme for the thesis "Taking deep uncertainty into account in traffic models: a case study of Groningen", written by Ilmo van Baarle for the master Engineering
and policy analysis at the TU Delft.
The goal of this research is to apply the RDM cycle to a traffic model written in PTV Visum. This needs a connection between the two programs.
With the help of this read me you will be able to better understand the code that is written.
At the end you should be able to conduct the same research.

## Installation

The main package that is used in the EMA_workbench. This library allows for an easy application of EMA research with many inherent analysis tools.
For further information on the package go to: https://emaworkbench.readthedocs.io/en/latest/
Another important package is the pypiwin32 library. This package allows Python and Visum to connect via a COM interface.

```
pip install ema_workbench
pip install pypiwin32
```

## Usage

The .py files are divided in two parts. Part one is needed to create and run scenarios. Part two is needed to create the results. The "first_try_at_connection" file shows the
iterative process of trouble shooting and testing with Visum commands and models.

### Creating and running scenarios

The two .py files needed for this are:
- Visum_connector.py: This file contains the connection with the Visum model, the way output is changed and the way lever are changed.
  - VisumModel: This can be kept almost completly the same. However, if you want to change uncertainties of links/nodes/zones, the run_experiment has to change a little.
    By changing Visum.Net.SetAttValue(k,v) to for example Visum.Net.Links.SetAttValue(k,v). More in depth tutorials of using this are found in the Visum COM manual.
  - output_changer: The output changer decides how the data is extracted. In this case after every run an Excel file is written.
    This could be anything you want. You could also read something from the Visum model for example.
  - lever_changer: There is a possibility to change the level of policies. In this case, the first thing is that a line which reads the levers has to be added in the 
    run_experiment function, before the experiments are run. Secondly, the lever_changer function needs to be adjusted to the levers that are changed.
    This also depends on the fact if it influences links/nodes/zones.
 - run with connector.py: In this file the first and second step of the RDM are taken, which means that the model is defined. The first thing that should be changed is the 
   defintion of the workingdirectory (wd). Then the VisumModel allows for three different inputs, the model_file (Visum model), the output_file (i.e. the Excel sheet) and the
   amount of replications. After that the policies, levers and uncertainties can be defined. Make sure that these names correspond with the names in the Visum model.
   In the definition of the policies, the model_files that correspond to a policy are given. The last step is defing the amount of experiments and parralel processes. Note,
   Visum runs max 5 runs simultaniously and the amount of runs are per policy. The output of this whole part is the variable "results".
   
### Analyzing the results

In the results folder 4 different python files are provided. Additionally some examples of results are handed. With all analyses the total distance travelled by car is 
researched. This variable and the value at which is looked should be changed manually. https://emaworkbench.readthedocs.io/en/latest/ shows other opinons and more in-depth
explanation of the methods used.
- Changing policies.iypnb: Here the robustness, PRIM and feature scoring analyses are performed.
- Robustness.iypnb: The robustness analysis is performed with three different metrics.
- PRIM.iypnb: PRIM analyses with the 10% worst and 10% best case scenarios.
- Feature_scoring.iypnb: Feature scoring with a parcoords analyses between the policies.


## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## Authors and acknowledgements

Ilmo van Baarle

In addition Jan Kwakkel has helped a great lot with the creation of the Visum_connector.py file.
I would like to thank him for his contribution.
