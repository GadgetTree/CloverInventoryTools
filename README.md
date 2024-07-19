# CloverInventoryTools

This is a very basic console application written by not a programmer.
The entire purpose of this program for the time being is to allow you to do an easy inventory count for CloverPOS.
It accomplishes this goal by adding a worksheet to the exported clover inventory excel workbook, called inventory count.
A file called output.xlsx will be saved and you can use the inventory count sheet to count all of your inventory.
Once done, run the program again and point it to your completed inventory sheet.
The program will then take those new inventory numbers and apply them back to the inventory sheet for clover.
This will spit out an inventory_done.xlsx workbook that you can import into clover.

Their are 2 advantages of doing this. 
1). The items are first reorganized by their respective categories. For my shop at least, this is helpful as all of my items tend
to be bunched together by the category they are associated with.
2). Even if you don't use a computer to do inventory, the inventory count sheet is formatted to print nicely on a portrait 
paper, so you can have manual inventory count anyways.



In the future, I may add inventory analysis to monitor suspected shrinkage per product and total cost of shrink.
I also may not.
