Overview of Project:

The main purpose of this project is to refactor and edit our existing code by taking fewer steps, using less memory, and improving the logic of our VBA code to improve its performance and to speed up the process. 
Two different blocks of codes were used in this project. One code was written during the learning module and the second code was a refactored code written for the purpose of this project.

The two codes used for this comparison can be found by downloading the VBA_challenge.xlsm file at the link below:

https://github.com/kbehyar/stock-analysis/blob/main/VBA_Challenge.xlsm

To distinguish between the two codes: Module_1 was the original code and Module_2 is the refactored version.


Results:

Analysis of module_1 code:

This code was written during the module learnings. I added several additional features to this code to match module_2:
-	A timer to show the process time of the code
-	An option to choose the year (2017 or 2018).
-	A formatting feature to present our data.
       
The two links below show the process time for the two years of 2017 and 2018 using the Module_1 code:

https://github.com/kbehyar/stock-analysis/blob/main/Resources/VBA_Challenge_2017_Prechange.PNG

https://github.com/kbehyar/stock-analysis/blob/main/Resources/VBA_Challenge_2018_Prechange.PNG


Analysis of module_2 code:

This code is the refactored version of the original code. 

The two links below show the process time for the two years of 2017 and 2018 using the Module_2 code:

https://github.com/kbehyar/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG

https://github.com/kbehyar/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png

Summary:

Advantages and disadvantages of refactoring code:

The main purpose of this project is to refactor and edit our existing code by taking fewer steps, using less memory, and improving the logic of our VBA code to enhance its efficiency and performance and to speed up the process. One of advantages of the refactored code is the process time which was significantly decreased. This will allow a much faster analysis of even larger data. 
While Module_1 code was more simplified and easier to follow for someone wanting to the learn the process. 


Detailed Analysis of the code:

The refactored code included to separate "for loops" as opose to the original code which included a loop within a loop.

While the two codes showed a very significant difference in processing time, I also ran the same codes on two different operating systems (PC and Mac). I noticed the OS with a superior CPU had a faster processing time. I also noticed that if there are other applications and software running in the background which I ran the codes, would cause a difference in the processing time since part of the CPU is processing these applications too.
