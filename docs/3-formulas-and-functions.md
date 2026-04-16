---
title: Formulas and Functions
parent: "Excel: Beyond The Basics - Windows"
layout: default
nav_order: 3
---

## Formulas and Functions

In this section, we’ll start using formulas to explore and analyze our data.

1. **A formula** in Excel is any mathematical equation. It’s made from values that we have entered into cells. You can create a formula from a direct value (ie. 5) ​​or using cell references (ie. A2). All formulas **begin with the symbol =**. This tells Excel you’re entering a calculation, and it needs to evaluate it.

2. You can use formulas to perform tasks such as adding numbers and multiplying values ​​in cells. Excel can understand and interpret a variety of mathematical operators, including:

    **+ (addition)**    
    **- (subtraction)**     
    *** (multiplication)**      
    **/ (division)**.   

    It can also understand comparison operators, such as:       
    **= (equal to)**    
    **> (greater than)**    
    **< (less than**    

    A full list of all operators is available via [Microsoft Support.](https://support.microsoft.com/en-us/office/calculation-operators-and-precedence-in-excel-48be406d-4975-4d31-b2b8-7af9e0e2878a)

3. **A function** is really just a formula - the only difference is that it’s a predefined formula that comes with Excel, not one that you create. **Functions** are built into Excel to perform specific calculations. For example:  

    <img src='{{ '/assets/images/N1_FormulasAndFunctions.png' | relative_url }}' alt='Cell C1 has the text of =A1+A2, then there is a visualization of this' title='' width='322' height='182' />

    **=A1+A2** is a formula, which adds the value of these cells together using the addition operator.

    <img src='{{ '/assets/images/N2.png' | relative_url }}' alt='Cell c1 now contains =sum(a1:a2) ' title='' width='348' height='184' />

    **=SUM(A1:A2)** is a formula in the form of a function. Excel knows that you want to add together a bunch of cells or numbers (this is what =SUM does), without needing to use the + (addition) operator. It also knows you want to add together all cells between A1 and A2, because you used a range operator (the colon!, : ).

4. Functions can increase productivity and make your formula much shorter. For example, if you wanted to add together all 5 cells you could write: **=SUM(A1:A5)**. This is much shorter than **=A1+A2+A3+A4+A5)!**

 

### Using Built in Functions
{: #using-built-in-functions}

1. Rather than type out every calculation by hand, we can use Excel’s built-in functions. Common calculations like averages, medians, sums, and maximums have their own Excel functions. A full list of math and statistical functions in Excel is available from [Microsoft](https://support.microsoft.com/en-us/office/statistical-functions-reference-624dac86-a375-4435-bc25-76d659719ffd).

2. If you think Excel may have the function you want to use you can go to the **Formulas** tab and select **Insert Function**. The functions are organized categorically in the function library (to the right of the **Insert Function** box).

    <img src='{{ '/assets/images/M2.png' | relative_url }}' alt='Insert function button within the ribbon is highlighted' title='' width='428' height='131' />

3. In cell **L1** of the **2024_Bus_Delays** worksheet, type a new column heading “Avg_Jan_Delay”.

4. Select cell **L2**. Let’s use the “Insert Function” button to look up how to calculate the average bus delay in January 2024. Average is found under **Category > Statistical**, but can also be located by searching for a function by name.

    <img src='{{ '/assets/images/M4.png' | relative_url }}' alt='Cell L2, insert function button, category as statistical, function is average' title='' width='1240' height='569' />

5. That is the first of two ways to access Excel’s built-in formulas. When you go this way, a pop-up window appears and gives you information about how to use the formula and what data the formula expects. When the formula asks for a number, you can enter a number, a cell identifier, a range. In this case, we want to calculate the average of all cells between **G2** and **G4765**. Enter **G2:G4765** Press **OK** to execute the formula.

    <img src='{{ '/assets/images/M5.png' | relative_url }}' alt='Function arguments as Number 1 in the range of G2:G4789, then press OK' title='' width='588' height='344' />

6. The cell now displays the average delay for the month of January.

    <img src='{{ '/assets/images/M6.png' | relative_url }}' alt='Displaying column L which now contains the "average January delay of 21.6458858"' title='' width='143' height='75' />

7. The other way to access built in functions is to begin typing the name of the formula  In cell **M2** of the **2024_Bus_Delays** worksheet, type in “Avg_Feb_Delay”, and let’s try this same calculation for the month of February.

8. In this case, once you start typing **“=A”** a dropdown will appear. Formulas can be typed directly in the cell, or in the **formulas bar** above.

    <img src='{{ '/assets/images/M8_0.png' | relative_url }}' alt='Cell M2 has =Average inputted' title='' width='834' height='169' />

9. You can double-click on the formula you want. A shadow explanation will show up, helping you understand the formula.

    <img src='{{ '/assets/images/M8.png' | relative_url }}' alt='There should be text under the cell which shows you how the formatting within the brackets should be' title='' width='300' height='96' />

10. Enter the cell range **G4766:G8966**. Hit **Enter** on your keyboard to execute your formula. We can now see that February had longer delays on average than January in 2024.

    <img src='{{ '/assets/images/I10_0.png' | relative_url }}' alt='Displaying the result within cell M2 as 22.86' title='' width='189' height='83' />

 

| **Formula not calculating?** Sometimes when using a formula, Excel will show you that formula instead of the results. Check to make sure the data type in your cell is not set to text. If it isn’t, you might have accidentally enabled the “Show Formulas” button on the Formulas ribbon - select this again to turn it off. Finally, check your formula for any typos. A single space before the = sign will mean that Excel won’t recognize your entry as a formula! |

 

### Calling Data from Other Worksheets
{: #calling-data-from-other-worksheets}

1. Let’s explore another mathematical function in Excel. Say we want to count how many times each bus route logged a delay in 2024. So, we want to know how many rows have 127 (ie. Bus #127) in column B “Route” of our **2024_Bus_Delays** worksheet, to determine which bus routes suffered from the most delays. We can do this using the **COUNTIF** function. This function counts the number of cells that meet a criteria that we set.

2. Let’s start by creating a new worksheet in our workbook. While still in **2024_Bus_Routes**, **click the + button** beside the worksheet names. A new worksheet will automatically appear after your active worksheet in the workbook.

    <img src='{{ '/assets/images/N2_0.png' | relative_url }}' alt='New sheet button is highlighted, as well as the new sheet mixed in' title='' width='483' height='73' />

3. **Right Click** on your new worksheet and select **Rename**. Rename the worksheet **Total_By_Route**.

    <img src='{{ '/assets/images/J3.PNG' | relative_url }}' alt='Context menu for the sheet is shown, rename is highlighted' title='' width='530' height='339' />

4. Go back to the **2024_Bus_Delays** worksheet. **Select** all of **column B**, “Route”, by clicking on column header B. **Copy** the contents of your selection. Now go back into **Total_Bus_Delays** and paste the copied cells into **A1**.

5. We’ll now want to remove duplicate values,  so that each bus route appears only once. With your pasted cells still selected, navigate to the **Data** ribbon. Under the **Data Tools** toolbox, select **Remove Duplicates**.

    <img src='{{ '/assets/images/N5.png' | relative_url }}' alt='Remove Duplicates button is highlighted' title='' width='462' height='114' />

6. Ensure that the correct column is selected. Make sure to leave “My data has headers” checked, and click **OK**. Excel will confirm the duplicates have been removed, and that 200+ unique values remain.

    <img src='{{ '/assets/images/N6.png' | relative_url }}' alt='Selecting the proper settings and clicking OK' title='' width='436' height='277' />

7. In cell **B1**, create a new column called “Total_Incidents”. Now let’s start typing our **COUNTIF** formula in cell **B2**. Remember, you can type this formula directly in the cell, or in the formulas bar at the top of your worksheet.

8. The **COUNTIF** function requires only two inputs: the **range** and the **criteria**. The range is the cells that you want to count (ie. where to look for matches, each match = 1 count) based on the criteria (ie. what you’re looking up).

    <img src='{{ '/assets/images/N8.png' | relative_url }}' alt='Cell B2 selected with the countif function inputted' title='' width='243' height='134' />

9. In this case, our range is all cells in column B of the **2024_Bus_Delays** worksheet, since each entry represents a reported delay on that route. Since this data is in another worksheet, we’ll need to add “worsheet_name!” immediately before our range. Hint: there are 50,040 rows in our **2024_Bus_Delays** worksheet.

    <img src='{{ '/assets/images/J10.PNG' | relative_url }}' alt='Formula bar has =countif('2024_Bus_Delays'!B2:B50040)' title='' width='467' height='65' />

10. When you’ve finished typing in your range, add the **comma** as shown in the shadow explanation. Our **criteria** in this case is the value of cell **A2** - this is what we want Excel to look for and count in our range. Type **A2** and close the formula using parentheses.

    <img src='{{ '/assets/images/J10_0.PNG' | relative_url }}' alt='Additions to the formula bar, =countif('2024_Bus_Delays'!B2:B50040, A2)' title='' width='538' height='127' />

11. Hit **Enter**. You should now see that route 89 had 515 delays in 2024.

    <img src='{{ '/assets/images/J11.png' | relative_url }}' alt='Highlighting the result in cell B2 of 515' title='' width='543' height='125' />

12. Now we want our calculation to apply to the whole column so that we know the total number of delays for each bus route. To do this, **select cell B2** then place your cursor over the **bottom right corner** of cell **B2**; you will see the cursor become a small black cross. **Click and drag** down the whole column (until the data ends in column A). You should see the formula but the row numbers should refer to that row’s data.

    <img src='{{ '/assets/images/N12.png' | relative_url }}' alt='Displaying the result after the column has been updated' title='' width='534' height='293' />

13. Let’s double check our results. Select any cell after row 2 in Column B. You’ll notice that our criteria has changed - which we want. But so has our range! This will cause issues, as the COUNTIF function is now skipping rows near the top of our worksheet.

    <img src='{{ '/assets/images/N13.png' | relative_url }}' alt='Checking the formula bar of a different cell to verify results are correct' title='' width='551' height='168' />

14. To fix this, we’ll need to **lock our cell values**. If you need a cell reference to stay unchanged when copying down a formula, you can add the '$' symbol to the cell references in the formula. This will lock those cells in place.

15. Let’s change our original formula in **A2.** You’ll need to add the ‘$’ symbol before both the column and row number, to lock both in place.

    <img src='{{ '/assets/images/N15.png' | relative_url }}' alt='Adding $ to the formula to lock parts of the formula while it is being copied' title='' width='569' height='127' />

16. Drag your new formula down. When you drag this formula down, the criteria changes but the range value will be locked in.

    <img src='{{ '/assets/images/N16_0.png' | relative_url }}' alt='Re copying the updated formula into the columns' title='' width='571' height='171' />

 

### Nested Functions
{: #nested-functions}

1. Let’s dive a bit more deeply into Formulas. What if you want to calculate or check multiple things in a single line? In these cases, you’ll need to nest your formula. We’ll explore this using another common formula in Excel, the **IF function**.

2. The **IF function** isn’t a math function, but a **logical function**. Logical formulas return either TRUE or FALSE when their arguments are evaluated - they are extremely useful tools to have at your disposal! A full list of these functions is available from [Microsoft](https://support.microsoft.com/en-us/office/logical-functions-reference-e093c192-278b-43f6-8c3a-b6ce299931f5).

3. The **IF function** is used when we want to test IF a condition is true or false, and return one value if the condition is met (TRUE) and another value if it is not met (FALSE). The IF function takes three inputs: the **logical test**, the value if **True**, and the value if **False**.

    ```
    =IF(A1>30, “High Temperature”, “Low Temperature”)
    ```
4. Let’s say we want to manage several conditions though, with several possible TRUE answers. For example, we may want temperatures equal to or greater than 30 to be defined as “High Temperature”, anything greater than or equal to 20 to be “Warm Temperature”, anything greater than or equal to 10 to be  “Average Temperature”, and anything under 10 to be “Cool”. In this case, we can use a **nested function**. Each subsequent IF formula is incorporated into the “value_if_false” argument of the previous IF. **So, this formula works as follows**:

    ```
    Test condition1, if True return result1, if False:
        Test condition 2, if True return results2, if False:
            Test condition3, if True return result3, if False:
                Return Result 4
    ```
5. In the case of our temperature data, we might write something like:

    ```
    =IF(A1>=30, “High Temperature”, IF(A1>=20, “Warm Temperature”, IF(A1>=10, “Average Temperature”, “Cool”)))
    ```
6. **Formulas are evaluated in order from left to right**, so in this case we don’t need to worry about setting a minimum value - the action on a cell is complete as soon as a condition is met. This means that if a cell has already been given the value “High Temperature”, the formula does not move on to check the next condition!

7. Let’s try this with our TTC data. In the **2024_Bus_Delays** worksheet, let’s add a column that identifies if a delay is short, medium or long using a nested IF function.

8. For this function, we want any **Min_Delay** under “15” to be **Short**. Anything between “15-30” is **Medium**. And anything over “30” is **Long**.

9. First, let’s create a new column. We want it to appear immediately to the right of **Min_Delay**. To create a new column there, we need to select column H **(Min_Gap)**. Then **right-click**, and choose “Insert”. When you use the “insert” option, the new column is always inserted immediately to the left of the selected column. Let’s name our new column **Delay_Type**.

    <img src='{{ '/assets/images/K9_NestedFunctions.PNG' | relative_url }}' alt='Insert button is highlighted on the context menu' title='' width='548' height='494' />

10. Now place your cursor in **cell H2**. We want to use the **IF function** to assign labels based on the length of delay. If none of those values are found (ie. our **Min_Delay** cell was empty), we want to return the value “Unknown”. You can type your formula either in the formulas bar, or directly in cell H2. **Tip:** There are several ways to type this formula correctly!

    <img src='{{ '/assets/images/O10_NestedFunctions_0.png' | relative_url }}' alt='Updating formula to enable a tertiary colour in the formatting' title='' width='520' height='129' />

11. Hit **Enter** once your formula is complete. You should now see “Short” in cell **H2.**

12. Since we have a lot of rows, dragging this formula down 50040 rows might take a while. A quick shortcut in Excel is to **select** the bottom right hand corner of cell H2, but instead of holding and dragging down with your cursor, simply **double click** with your mouse.

13. Our Delay Type column has now been populated with values in all cases. Nicely enough, this information lines up with the conditional formatting we applied earlier!

    <img src='{{ '/assets/images/O13.png' | relative_url }}' alt='Delay type column now contains a phrase, "short, medium, or long"' title='' width='304' height='368' />