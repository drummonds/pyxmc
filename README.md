# pyxmc
Python Excel Monte Carlo financial modelling

The aim of this is to create a framework to carry out the following:

- Create a Python model with input variables, formulas and output variables.
- Run a Monte Carlo simulation on that model
- Record the results in an Excel spreadsheet for consumption by a wider business community


By using Python the aims are:

- To create more robust and tested code
- To allow faster generation so that 10K and more runs are feasible
- To allow simple creation of input variables which are:
    - Point values or time series values
    - Discrete  or values with a standard deviation
- (Output variables)
- Potentially add more graphing types

By using Excel as a reporting medium:
- to create a numerical reporting format with easy access
- to distribute the data among a wider community


The reports will have:

- Summary of the input models
- Base line output model (using no variability)
- Monte carlo summary
- Monte carlo output model (with varation)
- Mechanism for adding plots to make graphs more meaningful


## Use
Idea is that you will import the framework, construct your model and generate your results.