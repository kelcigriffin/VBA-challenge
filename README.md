# VBA-challenge
Week 2 homework
I referenced several of our VBA class lessons to help me through this code. I needed additional help for some of the formatting, and for the last section, which asked us to compile the greatest increase, decrease, and largest total volume. I searched for VBA formulas and functions to determine min/max within a column or range. These are the resrouces that helped me form my code:

# Code Sources and Locations
This Stack Overflow page helped me determine the proper formula to format my decimal values into percentages
https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage

These links helped me find out how to determine min/max in a range. They also showed me how to write out the code, and understand the proper syntax to use. Especially for referencing the range effectively. I experienced a lot of trial and error, before landing on "ws.Range("K2:K" & last_row_yearly_change)". 
https://www.wallstreetmojo.com/vba-max/
https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction
https://learn.microsoft.com/en-us/office/vba/api/excel.range.insert
https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/refer-to-named-ranges#worksheet-specific-named-range
Referring to my range as "ws.Range("K2:K" & last_row_yearly_change)" is not my original idea. Variable names are mine, and are consistent with my script, but this way of referencing column and rows came from the following repository:
https://github.com/shrawantee/VBA-Scripting---Stock-Market-Analysis/blob/master/HW2_Challenge_DS.vbs
