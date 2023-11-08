# VBA-challenge
Module 2 Challenge README

The script begins on line 10 where it is named StockAnanlysis. 

The first grouping of information is where I begin having it function on all sheets within Excel. 

Following that is where I label all my new columns and rows so the final Excel sheets are easily readable. 

Next, beginning on line 25, is where I call out all the strings and doubles I plan to use.

Then the code begins. Here we are on line 37 and I am setting the values of openprice, vol, and start. 

My first for-loop starts on line 41 and ends on 78. This is creating the information for columns I, J, K, and L. Within this is also the conditional formatting for the colour scheme. 

Then on lines 80-82 I am setting the values for greatestincrease, greatestdecrease, and greatestvol so that they may be used in the for-loops ahead. 

The second for-loop in my script starts on line 84. This loop sorts through column L for the greatest volume. Below that is where I have the values determined sent to columns P and Q to be a part of the new table to the right. 

This process is repeated for each row in the rightmost table with their corresponding categories. 

Finally, on row 117 I have the script move to the next sheet in the Excel worksheet until the sub is ended. 


References:

Lines 37-78:
Tutor assisted by Sharon Colson- No code was provided but we collaborated together. 

Lines 87, 98, and 109:
Code provided by classmate David Roth
