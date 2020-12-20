# stock-analysis

## Overview of Project

### Purpose

The purpose is to refactor the Module 2 solution code to loop through all data one time in order to collect the same information in the module. Then determine whether refactoring code successfully made the VBA script run faster.  Since our original script ran on only a dozen stocks, this will show us if refracting can help us run a script on a much larger data set more efficiently.

## Results

In refractoring the code for this challenge (Sub AllStocksAnalysisRefactored()), the stocks still performed better in 2017 - Resources/VBA_Challenge_2017.png than in 2018 - Resources/VBA_Challenge_2018.png.  As you can see from the screen shots, the script for both years ran significantly faster! The output showed that 11 of 12 stocks had a positive return in 2017 and ran in .28 seconds compared with only 2 of 12 stocks performed positively in 2018 which ran in 0.29 seconds. The original All Stocks script (Sub AllStocksAnalysis()) for 2017 ran in 1.65 seconds and 1.99 seconds for 2018.

## Summary

One advantage to refractoring code is that it is less complex and easier to read. A disadvantage is that it is time consuming.

In refactoring the original VBA script, an advantage was that it ran the script much faster.  A disadvantage was that it was challenging and time consuming to get the final, refracted script completed. 
