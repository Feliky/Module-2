# Refactored Code Analysis

## Overview of Project

Excel VBA code was used to develop a routine to calculate the Total Daily Volume and Return on Investment (ROI) percentage for each of the 12 stocks provided. Upon completion, the data is summarized, formatted, and color coded for display. In addition to correctly automating the calculation, there is a need to expedite the execution of this code. The VBA code was refactored and execution times were collected and a comparative analysis is provided in this report.

## Results

## Unfactored 2017 Stock Analysis Results and Execution Time

![Unfactored_2017_Results](https://user-images.githubusercontent.com/84817579/185815281-20160640-72f9-46e3-ac50-8a74e271b6bf.png)

## Factored 2017 Stock Analysis Results and Execution Time

![Factored_2017_Results](https://user-images.githubusercontent.com/84817579/185815308-b2e1d5f5-e68c-4bf1-97ba-cf8ed884e05a.png)

## Unfactored 2018 Stock Analysis Results and Execution Time

![Unfactored_2018_Results](https://user-images.githubusercontent.com/84817579/185815397-8b0a58a6-687f-4f39-a499-a506512e5d8b.png)

## Factored 2018 Stock Analysis Results and Execution Time

![Factored_2018_Results](https://user-images.githubusercontent.com/84817579/185815435-de49b278-ad23-4f04-920c-c1e60835bf8d.png)

## Summary

Below is the summary of the VBA code execution times for Unfactored vs. Factored for both years 2017 and 2018. 

![Unfactored_vs_Factored_Summary](https://user-images.githubusercontent.com/84817579/185815460-10073393-da5a-440d-a3c5-77056e9315b3.png)

The outcome of refactoring code is to have the code running faster, codes that are easily able to handle all types of incoming data conditions and more date if neccesary. Refactoring code is the process of reviewing previously developed code into more efficient codes. 

The differences between the two codes is the choice as to when to update the "All Stocks Analysis" sheet. 

Another difference is removing the need to have an If Conditional statement asking whether the ticker matches prior to updating the TickerVolume. 
