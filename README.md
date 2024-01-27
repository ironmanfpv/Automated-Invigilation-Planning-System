# Automated Invigilation Planning System

 Rostering manpower around sets of constraints can be a laborious and time consuming task. 
 This project successfully implements an automated approach. Originally written to roster teacher manpower to invigilate on exam days, this excel planning program could be generalised to roster human resource manpower for duty over a 10 day work cycle. Input the constraints and automate the planning process. The program was finalised and last upgraded in 2021.


<img src= >

# SETUP STEPS: # 

## A. Set up parameters ##

1. Go to Exam Dates tab.
2. Key in the Date, Day and Day number for the period of use. (Column B, C and D)
   The day number is an indicator for staff's work schedule to reference, for constraint check.
3. Key in or Adjust the time period of each slot (Column H; default 30 mins time duration)
4. Key in the staff members' name (Column L)
5. Key in the Weekly duty slot quota for each staff (Column M)
6. Key in the workload in terms of duty slot for each staff, Sec 1 ~ 3 (can be generalised to 3 areas of work); (Column P, Q, R)
7. Assuming manpower is to be split to 3 groups, allocate each member to a respective group. (Indicate on Column T,U,V a number, typically based on Column S number.)
8. Save.


<img src= >


## B. Set up Constraints 1 ##

1. Go to TTData tab.
2. For each, staff, put an indication 'X' or any character, to block of the time slot that staff is NOT available.
3. If the staff does not report on a particular day, key in 'X' for the slots for the whole row of the day.
4. Repeat steps 1 to 3 for TTData2 if the work schedule is one that alternates weekly every fortnight. (2-week timetable)
5. Save.


<img src= >

## C. Set up Constraints 2 ##

1. Click on Sec 1 Tab. This tab contains the roster for the level over 10 days.
2. Insert the name of staff to exclude for the particular day due to unavailability (Column S).
3. Key in the examination name, time period and coordinator for each day. (Column L, K and M).
4. Repeat Step 3 for Sec 2, 3, 4 and 5 where necessary. Do this if many levels are having exam
5. Step 4 can be adapted if many departments are having activities.
6. Save

<img src= >

## D. Generate ##

1. Click the macro button to generate the deployment (Button is near column C and D on each tab)
2. Program will start to generate staff that is deployable for cells that have NO colour fill !
3. If no staff is required for a particular time slot and venue, shade the cell grey.
4. If you would like to fix the deployment for a slot, colour the cell a visible colour (eg. yellow)
5. Repeat the generate till deployment is satisfactory.
6. At the end of each generation, the time taken to generate is displayed.
7. The more slots requiring deployment, the longer the generation time.
8. The more constrained the situation, the longer the generation time.
10. Save the planning regularly.
11. As the program is intented on a deep root search algorithm, if no solution is found, the program may need to be terminated manually.
12. It is possible to do the planning first by leveraging Steps 2 to 4 and leave difficult zones for the program to find solutions.


## E. Identifying the Non-optimised deployment ##

1. As the deployment is generated based on a fair distribution of load, the first iteration may not be optimised.
2. For example, there could be a case of "triangular bootstrapping" 
   Eg. current slot: A in venue 1, B in venue 2. C in venue 3. In the next slot, 
                     A --> venue 2, B --> venue 3, C to be relieved when D---> venue 1.
                     If the venue cannot be left unattended, C to be relieved is contingent on D arriving at venue 1 punctually.
                     In the extreme case, if C ---> Venue 1, then all A, B, C are bootstrapped.
                     As such, there could be a cascaded delay effect and the deployment is not optimised.
3. The optimised button is found at near column M and N.
4. Click it to highlight all the optimised regions. The non-optimised regions will remain uncoloured after inspection.


## F. Swapping ##

1. The non optimised scenario can be remedied by doing what is known as "horizontal swap" 
2. This means that if staff A has to on duty for 3 consecutive blocks, ideally, staff A should not be moving from venue to venue during this period.
3. The non optimised scenario also bypasses the optimisation check if the staff has a break in between duties.
4. This means that to ensure consistency in reporting venues, a swap method is in available and in place, if needed.
5. The swap button is found in (Column Z) at each day section.
6. The swapping process only swap slots for cells that are not coloured filled. Uncolour the cells before using swap.
7. Ensure the deployment is saved before attempting swapping.
8. Click it to do horizontal swapping.
9. Note that swapping is an irreversible process as cache memory is not dedicated to store the swap process.

## G. Completion ##

1. At the Index Page tab, the deployment statistics table can be seen.
2. At the deployment frequency tab, the deployment statistics are displayed as bar charts for visual analysis.
3. At the end of the planning, each staff member's duty roster may be obtained by printing the respective tabs.
4. Click the tab to print the duty roster. Do a mass batch print is possible by selecting multiple tabs and batch print.
5. Softcopy PDF version is also available when selecting print properties to PDF.

## H. Troubleshooting ##

1. If clicking the button does not start the program, it might be that Excel macros are blocked by your local version of Excel.
2. Allow macros by going to File>Options>Trust Centre>Trust Centre settings>Macro settings>Enable VBA Macrios
3. As windows suite constantly change their menus with different versions of Excel, pls check for the latest steps to enable macros.