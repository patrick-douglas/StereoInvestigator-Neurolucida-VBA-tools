
# Stereology and Neurolucida VBA
#### Stereology and Three-dimensional reconstruction Excel tools
Hello everyone, today I will present a VBA code to help stereologists process their data from Stereoinvestigator and Neurolucida software.
### Which kind of language this software awas made?
It was made to run directly in excel through VBA
#### How to use it?
To run the macro you just need download and open the **xlsm file** into Excel (Tested on Excel 2019)
##### Allow the macros
For security reasons you need allow access to macro code before run the macro, after open for the first time you will see the screen below, you just need click in **Enabled content**:

![!](https://github.com/patrick-douglas/SIandNL_Macros/blob/master/Wiki/Allow-Macro.jpg)

#### 1 - Create a summary table from [Optical fractiotator probe](https://www.mbfbioscience.com/help/si11/Content/SI_SPECIFIC/Probes/Optical_Fractionator.htm) from [StereoInvestigator](https://www.mbfbioscience.com/stereo-investigator)
First we need to download the `.xlsm` file called [(Stereology-Summary-Table.xlsm)]([https://github.com/patrick-douglas/SIandNL_Macros/blob/master/SI/Stereology-Summary-Table.xlsm](https://github.com/patrick-douglas/SIandNL_Macros/blob/master/SI/Stereology-Summary-Table.xlsm)) from the git-hub page, after download open the file in the Microsoft Excel (2016/2019 is preferred).

### Running the code 
#### WARNING:
IS VERY IMPORTANTE **NEVER RENAME THE `.XLSM` FILES**, IF YOU DO THIS THE MACRO WILL NOT WORK!
#### 1. Where the results will be placed?
Every time  you run the macro remember that the results will be placed below the selected cell and will  be placed in the left to right sense, so before start the code place the cursor in the cell **A1** to get the results in cell **A2**:

![!](https://github.com/patrick-douglas/SIandNL_Macros/blob/master/Wiki/parameter-opctical-fractionator.jpg)

#### 2. Getting the results
As default in **Excel** you just need press `Alt+F8` to open the **Macro Run Window**, select the macro in list and click in **Run**

![macro-window](https://github.com/patrick-douglas/SIandNL_Macros/blob/master/Wiki/macro-window-example.PNG)

After click **Run** a new window will opened and you need select which raw data you will select (You can select multiple files). You need choose the correct workbook for the correct macro you are running.

![example-selecting-files](https://github.com/patrick-douglas/SIandNL_Macros/blob/master/Wiki/Secting-files-Example.PNG)

After click in **Open** you will see the **Excel window blinking** don't worry everything is fine, the code are Running. Depending on your data set the process may take a very long time, be patient. 
When finished the results will be displayed as bellow:

![example-result](https://github.com/patrick-douglas/SIandNL_Macros/blob/master/Wiki/Results-example.PNG)
>IMPORTANT: Don't use you computer while the macro is running this may cause errors in the process.
#### 3. Which types of results can I get? 
The macros can create summary tables from three types:
 1. Summary Parameters:
This is the first **Sheet** in excel named as `Parameters-OpticalFractionator`, the column heads show which type of information the code will extract, to choose this one you just will need run the macro `optical_fractiontor_parameter` in the **Macro Run Window**
>IMPORTANT: This Macro only will run with workbooks extracted from **Optical Fractionator Probe** generated from software **StereoInvestigator**.

2. Data Summary:
This is the second **Sheet** in workbook named as `DataSummary-OpticalFractionator`, the column heads show which type of information the code will extract, to run this one you just will need choose the macro `optical_fractionator_data_summary` in the **Macro Run Window**.
>IMPORTANT: This Macro only will run with workbooks extracted from **Optical Fractionator Probe** generated from software **StereoInvestigator**.

3. Cavalieri Estimator data Summary:
This is the third **Sheet** in workbook named as `DataSummary-CavalieriEstimator`, the column heads show which type of information the code will extract, to run this one you just will need choose the macro `optical_fractionator_data_summary` in the **Macro Run Window**
>IMPORTANT: This Macro only will run with workbooks extracted from **Cavalieri Estimator** generated from software **StereoInvestigator**.

