SALES REPORT AND DISCOUNTS
This project involved working with large xlsm file. 
The macro enable excel workbook had more than 195 worksheets 
which kept growing over time. 
Each worksheet was for separate customer, 
or for a data dump.
The data dump was pasted in the "DATA" sheet of the document.
The document had numerous LOOKUP and FILTER functions in each sheet. 
This caused the file to respond slow.
From the "DATA" sheet all other sheets were
populated using FILTER and LOOKUP functions. 
Dealer wise, product wise discounts were applied
MANUALLY. Sales Report sheet had to be manually updated each month for new customers.


WHAT I WAS REQUIRED TO DO.


I. Create a sheet in the existing document which would have the following: 
(a) Customer List
(b) Products on which discounts were applicable.
(C) The slab of discount

2. Calculate Customer wise, productwise discount
and place in this sheet;
and in the respectivecells of the respective customers' sheet.

3. Calculate a flat discount for qualifying customers and apply these
   in the respective Customers sheets.
   Also, The flat discount should not apply to the
   discounts calculated earlier (Product wise).

5. Also, for each discount calculation, there should be separate Python Scripts.
   The user would decide whether a particular script needs to be run.

6. In the end, generate a sales report.







