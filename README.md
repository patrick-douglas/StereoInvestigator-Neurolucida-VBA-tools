
# Stereology and Neurolucida VBA
#### Stereology and Three-dimensional reconstruction Excel tools
Hello everyone, today I will present a VBA code to help stereologists process their data from Stereoinvestigator and Neurolucida software.
### Which kind of language this software awas made?
Is was made to run directlly in excel thought VBA
#### How to use it?
To run the macro you just need download and open the **xlsm file** into Excel (Tested on Excel 2019)
##### Allow the macros
For security reasons you need allow access to macro code before run the macro, after open for the first time you will see the screen bellow, you just need click in **Enabled content button**:

![Wiki]([https://github.com/patrick-douglas/SIandNL_Macros/blob/cad411e37ee5d8ff1ceed4b7320fd8b7d174f672/Wiki/Allow-Macro.jpg](https://github.com/patrick-douglas/SIandNL_Macros/blob/cad411e37ee5d8ff1ceed4b7320fd8b7d174f672/Wiki/Allow-Macro.jpg))

#### 1 - Create a summary table from [Optical fractiotator probe](https://www.mbfbioscience.com/help/si11/Content/SI_SPECIFIC/Probes/Optical_Fractionator.htm) from [StereoInvestigator](https://www.mbfbioscience.com/stereo-investigator)
First we need to download the `.xlsm` file called [(Stereology-Summary-Table.xlsm)]([https://github.com/patrick-douglas/SIandNL_Macros/blob/master/SI/Stereology-Summary-Table.xlsm](https://github.com/patrick-douglas/SIandNL_Macros/blob/master/SI/Stereology-Summary-Table.xlsm)) from the git-hub page, after download open the file in the Microsoft Excel (2016/2019 is preferred).

 1. Create a Parameters table
Every time  you run the macro remember that the results will be placed below the selected cell and will  be placed in the left sense, so before start the code place the cursor in the cell **A1** to get the results in cell **A2**:

![!](https://github.com/patrick-douglas/SIandNL_Macros/blob/master/Wiki/parameter-opctical-fractionator.jpg)
