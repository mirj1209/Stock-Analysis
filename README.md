# Stock-Analysis.

## Overview of Project
### Background
During this module, I created a workbook. At a click of a button, one can analyze the entire data set for stocks during 2017 and 2018. However, the code I created although works, it’s not as time-efficient as it can be. Thus, this leads to the purpose of this project.
### Purpose
The purpose of this project is to edit/refactor the workbook. With the goal to make the code more time-efficient and let the script run faster. Also to see whether or not the modifications will have any effect on the codes running time.

## Results
During 2017, it seems that most stocks; that was invested on; were thriving returning at least 20% back. With a few exceptions that were below 10% return and one returning a negative percentage. However, during 2018 almost every single stock (apart from RUN and ENPH) began returning a negative percentage. This tells us that apart from RUN and ENPH, every single stock is extremely volatile (this could be further investigated if there were any monthly data) for example SEDG went from returning 184.5% to -7.8%. That is almost a 200% drop.  
In terms of the code, before refactoring the time stamps for the previous code were of a little over a second, the excel sheet buffered a bit (blinked a lot) before giving the numbers/results. After refactoring, however, the timestamps have reduced significantly by an entire second as the code ran in under half a second for both 2017 and 2018. 

![refactored code running time](https://user-images.githubusercontent.com/104941338/170326951-1f99ae72-e354-4eb5-9d2d-98e74bdac632.png)

The biggest difference between the original code and the refactored code is that the original code went through multiple If statements within a single loop rather than making multiple smaller loops. As the original code had a loop that had 2 separate If statements, where the code had to go back and forth between cells to see if they fit the If statement requirements. Whereas, in the refactored code, there was an index to specify which kind of cells it was looking into for the computer to not go back and forth within every single cell that could possibly fit the If statement requirements. The creation of the index was the difference that cut the running time significantly as it cut a great number of cells the computer was looking at while running the code. 
![code comparisson](https://user-images.githubusercontent.com/104941338/170327084-359b11a9-41e5-4162-81da-811db05e3b02.png)

## Summary
Some of the advantages on refactoring a code are that the code length gets significantly smaller thus having less lines to look at when trying to debug any issues. However, this also means that it’s much easier to make a mistake within the code where you’ll have to change some fundamental things either at the beginning or after where VBA tells you where the bug is. I know this from personal experience as during the refactoring process, I came across multiple bugs where they all highlighted the exact same line but the issues were either because I hadn’t referenced to the index properly or because I missed a small typo or forgot to add a minor thing later on. 
