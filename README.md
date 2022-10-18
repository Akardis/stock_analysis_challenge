# stock_analysis_challenge
mod 2 challenge



## Overview of Project
  Help Steve with stock analysis. He wants to be able to utlize the spreadsheet that we made in the module to be able to look at the whole stock market. Refactor the code from the module to be use for the whole stock market. 
### Purpose
  Showcase skills learned in the module and class this week for VBA. VBA allows you to code for excel. VBA can be used to write code that will hlpsort through large amounts of date and look fro information. Code can be written to utilize the formulas in excel. The code can be set up so that new information can be looked for quickly, through new inputs into the code. 
## Analysis and Challenges
  Attempted the challenge after class and was not able to get through it. I went back and redid the modules. I am having trouble with the adjustment of volumes in "If then" conditionals. I copied the hint from the module for step 3a. I get the concept, yet i am having trouble writing and fixing it. 
         
         3a) Increase volume for current ticker
              If Cells(j, 1).Value = ticker Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
              End If
              'This is where i have gotten stuck. Something about the tickerVolume array is off.
              'I copied this code from the module
          3b) Check if the current row is the first row with the selected tickerIndex.
             If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               tickersStartingPrice = Cells(j, 6).Value
            End If
 
 I attempted this question in the challenge the same way i looked at it in the practice work. I ran into an issue with the tickerVolume. Asked the class about thier thoughts. 
    
 With data sets like the stock market having code that allows for you to change a few inputs is helpful. When just dealing with a year or two it was simplier to write the code. Yet, that code wouldn't work for the whole stock market. The changes, eventhough I had some challegnes, made sense for the larger data. Being able to just add a new ticker and or a year is the best way to write code in VBA.
 VBA makes using Excel more applicable than I ha initally thought possible. I can see how this skill set is applicable for jobs. Understanding VBA makes unitlizing Excel easier and allows one to do more things. 
### Analysis of Outcomes of Stocks

### Challenges and Difficulties Encountered
  Loop and conditional is something that i struggled with. I understand in theory how they work, I just have a hard time applying it. 
## Results
  Was not able to finish the chalenge. Got stuck and therefor was not able to get results. It seems though that refactoring the code from Module 2 would make the program run faster. Leaving the code how it was oringally written would take much longer to proccess the whole stock market. I assume that the time for the code to run after it was refactored was significantly faster. 
- What are the advantages or disadvantages of refactoriung code?
    If it aint broke, don't fix it. That was how I was beginning to feel. 

- How do these pros and cons apply to refactoring the orginal VBA sript?
