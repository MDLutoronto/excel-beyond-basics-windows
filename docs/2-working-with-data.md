---
title: Working with Data
parent: "Excel: Beyond The Basics - Windows"
layout: default
nav_order: 2
staff: 
    - name: Kara Handren
      link: https://library.utoronto.ca/staff/kara-handren
maintainer:
    - name: Kara Handren
      link: https://library.utoronto.ca/staff/kara-handren
created_date: 2023-02-21
---

## Working with Data
* [Understanding Data](#understanding-data)
    + [Number Data](#number-data)
    + [Text Data](#text-data)
    + [Error Data](#error-data)
    + [Logical Data](#logical-data)
* [Viewing and Editing Data Types](#viewing-and-editing-data-types)
* [Viewing and Sorting Data](#viewing-and-sorting-data)
* [Conditional Formatting](#conditional-formatting)
* [Filtering Data](#filtering-data)
* [Transposing Data](#transposing-data)


### Understanding Data Types
{: #understanding-data}

There are four different kinds of data in Microsoft Excel: text, number, error, and logical. **Data types** may change when pasting data into a spreadsheet, so it's important to understand what type your data exists as. Since the type of data determines what functions you can perform on it, it's also important to know which ones to use and when to use them.

#### Number Data
{: #number-data}

**Number or numerical data** can include any type of number, of any size (including fractions and negatives). There are many different subtypes of numerical data in excel, including whole numbers (integers) and real numbers (decimals), dates and times, phone numbers, and monetary totals in varying currencies.

<img src='{{ '/assets/images/number_data.png' | relative_url }}' alt='An example of numerical data - Date: 45292; Route: 113; Time: 03:37.' title='' width='313' height='134' />

 

#### Text Data
{: #text-data}

**Text data** can include letters, numbers, and symbols, and is often used in column names as well as data cells. However, unlike numerical data you can’t perform calculations on text data. Excel will often classify data as text by default if it doesn’t recognize the type, so you may need to manually change the format of your cells before running any calculations! We’ll talk more about this in a minute.

<img src='{{ '/assets/images/text_data.png' | relative_url }}' alt='An example of text data - Day: Monday; Location: MAIN STATION; Incident: Security.' title='' width='311' height='133' />

 

#### Error Data
{: #error-data}

**Error data** will pop into your cell when there is a mistake or missing information when Excel is processing a formula. There are several types of error data in Excel, which are helpful in determining what may have gone wrong. A full list of these errors can be found via [Microsoft support](https://support.microsoft.com/en-us/office/detect-errors-in-formulas-3a8acca5-1d61-4702-80e0-99a36a2822c1)).

| **#DIV/0** | You’re trying to divide two values, but the first number is a zero or an empty cell. |
| **#N/A** | Your formula can’t find the value it’s looking for. |
| **#NAME?** | You’ll get this error if your function contains text that Excel doesn’t recognize (for example, your function contains a typo ie =VLOKUP instead of =VLOOKUP |
| **#NUM!** | This error is caused when a number for an argument in a function is not valid. For example, if you’re trying to find the square root of a negative: =SQRT(-2) |
| **#REF!** | This reference error occurs when we accidentally delete or move the cell we reference in a formula |
| **#VALUE!** | This error might pop up if you use the wrong data type in a function or formula. For example, if you try to add text using the + operator. |

#### Logical Data
{: #logical-data}

**Logical data** values are often shown as either TRUE or FALSE. Like with error data, they cannot be typed in Excel manually by the user and instead will show up as a result of an expression or function. Logical data type is useful in making comparisons, creating conditions, testing these conditions, and checking the contents of a cell location.

<img src='{{ '/assets/images/logical_data.png' | relative_url }}' alt='Result of a formula resulting in a cell being displayed at TRUE' title='' width='451' height='190' />

 

### Viewing and Editing Data Types
{: #viewing-and-editing-data-types}

1. When you select a cell in Excel, the data type is visible by default from the **Numbers section** of the **Home** ribbon. You can change the data type by selecting your cells, and then choosing a new option from this drop-down menu.

    <img src='{{ '/assets/images/C1.PNG' | relative_url }}' alt='Having selected a cell, the Number tab on the ribbon will display the current data type of the cell' title='' width='376' height='266' />

2. Data type can also be reviewed and changed by selecting a cell or group of cells, right clicking, and selecting **Format Cells** from the drop-down menu.

    <img src='{{ '/assets/images/C2.png' | relative_url }}' alt='Format cells is highlighted on the context menu' title='' width='294' height='580' />

 

    | **Data alignment** 
    By default, all numeric data (including dates) are what is known as right-aligned, while all text (string data) are left-aligned. This is a helpful way to see when something you entered wasn’t quite right and hadn’t been recognized as the correct kind of data by Excel. |

 

3. Navigate to the **5Y_Revenue** worksheet. You can see right away that something isn’t right - these don’t look like revenue numbers! **Click on any cell with revenue data** and check the numbers section of the home ribbon - you’ll notice that this value is being displayed as **Text data**. This means that Excel has not recognized the data as numerical data, and so it’s displaying the contents exactly as entered. So, our revenue data is being displayed as strings of text.

4. We need a way to manually tell Excel that these cells are revenue data (numerical / currency data). We also want to make sure that all other numbers and dates are also being considered number data, with the correct subformat types.

5. In the **5Y_Revenue** worksheet, **select only the cells that contain currency information**, not the row/column headers. Navigate to the **numbers section** of the **Home** ribbon, and use the dropdown menu to select **currency**.

    <img src='{{ '/assets/images/C5.PNG' | relative_url }}' alt='Clicking on the number type, change it to currency' title='' width='814' height='446' />

6. Excel is now recognizing and displaying our data as number data in the form of a monetary value (currency) and has included by default a dollar sign. Note that this sign may not be the default if you’ve set the region/language settings for your entire operating system to something other than “English (Canada)”. You can always easily change the symbol of currency values in Excel by selecting the **$** dropdown menu from the numbers tab of the **Home** ribbon. Choosing **More Accounting Formats** will provide you with a fairly complete list of national currencies and symbols, including CAD.

    <img src='{{ '/assets/images/C6.PNG' | relative_url }}' alt='Highlighting the more accounting formats under the extra currency options' title='' width='496' height='287' />

 

    | **Shortened Numbers in Excel** 
    When you punch in long numeric strings into Excel, say, 12345678901234567890 (20 digits), Excel will generally convert it for you, meaning that the 20 digit number you've just tapped in has been cut back to be only about fifteen significant figures. |

 

7. Let’s go back to the **2024_Bus_Delays** worksheet. The values in our “Date” column also look odd, and the culprit is the same! Ensure that the cells under the “Date” column are selected, and then choose data type **Short Date** from the numbers tab of the Home ribbon.

    | **Selecting Multiple Cells** 
    An easy shortcut to selecting all cells below your current selection in a column is to use **CTRL + SHIFT + ↓**. Note that this selection will only extend to the first blank cell. |

    <img src='{{ '/assets/images/C7.PNG' | relative_url }}' alt='Selecting short date for the column that contains the date information' title='' width='728' height='548' />

8. Our dates are now formatted as short dates, which are a lot easier to understand.

    <img src='{{ '/assets/images/C8.png' | relative_url }}' alt='Highlighting that the column has reflected the change made' title='' width='331' height='257' />

    | **Working with Numbers** 
    Don’t worry if your dates look odd after you copy/pasting into a worksheet. Excel stores dates as sequential serial numbers so that they can be used in calculations. These dates start at January 1, 1900 by default. So this date is serial number 1, which means that January 1, 2024 is serial number 45292 because it is 45,291 days later. These mean the same thing to Excel; the number can always be converted to a date format later on without losing any details. |
 

### Viewing & Sorting Data
{: #viewing-and-sorting-data}

1. Let’s move to the **2024_Routes** worksheet in our workbook. In this section, we’ll start to look at strategies to manipulate your data to help you make sense of it.

2. Have a look at the data. Often, you’ll be working with data you received from someone else; for example, it might be provided by an instructor for an assignment, or you might have downloaded it online. Data created by someone else can sometimes be messy and require some cleanup. Or, as in this case, it might contain extra information that you don’t need.

3. The first thing we might notice is that we can’t read all of the column names in full. Hover your cursor over the line between **columns A and B** in the column header. An icon that looks like a vertical line with arrows pointing right and left appears. When you see this icon, you can manipulate the width of the columns. **Click and drag to the right** to make column A wider.

    <img src='{{ '/assets/images/D3_Viewing_and_Sorting_Data.PNG' | relative_url }}' alt='Highlighting the column divider' title='' width='529' height='167' />

4. Another option you can use when you see this icon is to **double-click** the line between columns in the header, which will “auto-fit” the data in that column. Warning: if you have very long column names, you will produce very wide columns using this tool! Try this for column B.

5. You can also auto-fit the data for the entire worksheet at once. Click in the top left-hand corner of the data. This will select your entire worksheet. Then you can double click in the column header between any of the columns, and it will auto-fit all columns at once.

    <img src='{{ '/assets/images/D5.png' | relative_url }}' alt='Top left triangle which selects every cell' title='' width='526' height='169' />

    | ##### 
    If Excel displays ##### in a cell after you apply currency formatting to your data, the cell probably isn't wide enough to display the data. To expand the column width, double-click the right boundary of the column that contains the cells with the ##### error. This automatically resizes the column to fit the number. You can also drag the right boundary until the columns are the size that you want. |

 

6. Before we start sorting our data, let’s freeze the column names at the top of the screen so we can always see them, even as we scroll. Select the **View** tab in the ribbon, then in the Window area, choose **Freeze Panes ->  Freeze Top Row.** It’s also possible to select and freeze multiple rows or columns using **Freeze Pane**, depending on your needs. Do this for both the **2024_Routes** and the **2024_Bus_Delays** worksheets.

    <img src='{{ '/assets/images/D6.PNG' | relative_url }}' alt='Freeze top row which keep the top row visible while scrolling through the worksheet' title='' width='410' height='258' />

7. You’ll notice a column titled **route_type**, which tells us if the route is a subway (1), streetcar (0), or bus (3). Now let’s sort the data based on that value. **Select** any single cell within the **route_type** column. Then go to the **Data** tab on the ribbon, and in the **Sort & Filter** section, choose the **A-Z** button.

    <img src='{{ '/assets/images/D7.PNG' | relative_url }}' alt='Sort the entire worksheet in ascending order based on the column of the selected cell' title='' width='619' height='333' />

8. The entire worksheet has now been sorted A-Z (or smallest to largest in the case of numerical data) based on the value in Column E.

9. Always check sort results carefully to ensure that all relevant columns were affected by the sort. If you have blank rows or columns, this can confuse Excel. For more sorting options (and to verify the sort will impact all of your data), **select the Custom Sort button** from the **Data** ribbon (Note: Depending on how your Excel is configured, this option might also be under “Sort”).

    <img src='{{ '/assets/images/D9_1.PNG' | relative_url }}' alt='variations in the way that the sort button works, highlighted is the sort button, and custom sort optionariations in the way that the sort button works, highlighted is the sort button, and custom sort option' title='' width='795' height='354' />

    <img src='{{ '/assets/images/D9_2.PNG' | relative_url }}' alt='Variations in the way that the sort button works, highlighted is the sort button, and custom sort option' title='' width='221' height='113' />

10. But what happens when two observations have the same value? If we sort by **route_type** and there are multiple routes of each type, what order should those routes be listed in? By adding a second variable, you tell Excel exactly what to do. For all the route types, let’s say that we also want to sort them alphabetically by route name.

11. Open the **Custom Sort** menu. Ensure that “my data has headers” is highlighted, so that your variable names are not sorted like an observation.

12. Click on **Add Level** to add a second variable to sort by ie. **route_long_name**. Excel will sort initially on your first variable, and then those results based on your second variable. Ensure **Sort On** is set to “cell values”, and **Order** is set to “A to Z”. Press **OK**.

    <img src='{{ '/assets/images/D12.PNG' | relative_url }}' alt="GUI that will then open. add level, checkmark 'my data has headers' and edit the column sorting" title='' width='587' height='269' />

13. Our data has been sorted according to the variables we chose.

    <img src='{{ '/assets/images/D13.png' | relative_url }}' alt='Worksheet sorted with two criterion' title='' width='608' height='340' />

 

### Conditional Formatting
{: #conditional-formatting}

Conditional formatting allows you to automatically apply formatting, such as colour, to cells based on the cell value. This can make it easier to quickly identify cells that meet certain conditions (such as outliers in your data or the top 10% of values). This can also help you to better sort your data, since you can sort by colour using the same method we just applied. To do this, you'll need to **create a conditional formatting rule**.

1. Let’s go to our **2024_Bus_Delays** worksheet. Let’s say we want to highlight in yellow all delays longer than **15 minutes**, and in red all delays that lasted longer than **30 minutes**.

2. Select the cells you’d like to format. In this case, let’s select the entire **Min_Delay** column excluding the column name.

    <img src='{{ '/assets/images/E2_Conditional_Formatting.png' | relative_url }}' alt='Column G is highlighted' title='' width='298' height='216' />

3. From the **Home** tab on the ribbon, select the **Conditional Formatting** button. Navigate to **Highlight Cell Rules** and select **Greater Than…**

    <img src='{{ '/assets/images/E3.png' | relative_url }}' alt='Conditional formatting in the ribbon is highlighted' title='' width='432' height='477' />

4. You have the option to select a custom combination of colours, borders or text, or otherwise use default presets to highlight your cells. Let’s use the defaults in this case. Input “15” under **format cells that are GREATER THAN**. Select **Yellow Fill** from the highlight rules menu. Click **OK**.

    <img src='{{ '/assets/images/E4.PNG' | relative_url }}' alt='Cell values of 15, and yellow fill with dark yellow text' title='' width='473' height='133' />

5. Let’s repeat this process with a **Red Fill**, for cells with a value greater than “30”.

    <img src='{{ '/assets/images/E5.PNG' | relative_url }}' alt='cell values of 30, and Light red fill with dark red text' title='' width='472' height='133' />

6. Our cells have now been highlighted to reflect our new conditional formatting rules.

7. You can review and manage conditional formatting rules by clicking on **Manage Rules** under the **Conditional Formatting** options in the Home ribbon. Here you can edit or delete rules. You can also change the order of rules if you would like one rule to take precedence over another.

    <img src='{{ '/assets/images/E7.png' | relative_url }}' alt='Review of the changes that have been done' title='' width='625' height='296' />

    | **Note** 
    If you want to highlight based on multiple conditions at the same time, you can do this by navigating to New Rule and selecting **Use a formula to determine which cells to format**. We'll learn more about formulas later in this workbook. |

 

### Filtering Data
{: #filtering-data}

1. Another key tool to help you explore your data is filtering. On the **Data** tab of the ribbon, in the “Sort & Filter” section, choose **Filter**.

    <img src='{{ '/assets/images/F1_Filtering_Data.PNG' | relative_url }}' alt='selection of the data tab and filter icon' title='' width='754' height='128' />

2. Little arrows should appear on each of your column names. These contain the filter options for each of your columns.

    <img src='{{ '/assets/images/F2.png' | relative_url }}' alt='Visual verification of the filter icons along the top row of every column' title='' width='1022' height='135' />

3. **Click** the arrow to open the filter options for **Incident**. Notice that sorting options are also available here. In the lower part of the menu, **uncheck** “Select all”. Then, manually check off several options. In this case, we are only interested in delays caused by collisions, investigations, operator or mechanical reasons. Click **OK**.

    <img src='{{ '/assets/images/F3_0.PNG' | relative_url }}' alt='An overview of the context menu and how it should appear' title='' width='300' height='550' />

4. Only the selected incident types now appear in your view. Notice that **column F** now has a filter graphic on the button, which indicates that a filter has been applied on that column. You’ll also notice hidden row numbers when scrolling through your worksheet - the data is still there, it’s just been hidden from view. **Note**: When you run calculations, Excel will still calculate these hidden rows unless you ask it to ignore hidden rows!

    <img src='{{ '/assets/images/F4_1.PNG' | relative_url }}' alt='Visual notification showing that the column is being filtered' title='' width='400' height='133' />

  
    <img src='{{ '/assets/images/F4_2_0.png' | relative_url }}' alt='Visual proof of row 7-9 being hidden' title='' width='91' height='43' />

5. Let’s imagine that we want to add a further filter to this dataset. We only want to see those rows where the delay resulted in a vehicle gap (ie. time between vehicles at a given stop) of greater than 10 minutes. Open the filter options for the **Min_Gap** column. Choose Number Filters, then **Greater Than Or Equal To…**

    <img src='{{ '/assets/images/F5.PNG' | relative_url }}' alt='Min_Gap column filter, number filters, followed by greater than filter location' title='' width='383' height='376' />

6. The “Custom AutoFilter” window pops up. Enter the expression “is greater than or equal to 10” and click **OK**.

    <img src='{{ '/assets/images/F6.PNG' | relative_url }}' alt='Selecting the correct settings and clicking ok' title='' width='722' height='219' />

7. We can see that now any values less than 10 have been removed from our view.

8. You can remove filters from specific columns, by opening the filter options for that column and selecting **Clear filter from** [column name]. Let’s **clear** our filters for now.

    <img src='{{ '/assets/images/F8_0.PNG' | relative_url }}' alt='Clicking the clear filter button to remove it' title='' width='395' height='412' />

 

### Transposing Data
{: #transposing-data}


1. Transposing data is another helpful data management task. Switch over to the **5Y_Revenue** tab. Looking at the table, you decide that you wish this data had the years across the top of the table (as column headings) and the months along the side (as row headings).

2. There is a very quick way to fix this. **Select all the cells in the data table**, excluding the table name and descriptive information. You’ll notice this is the first worksheet we have looked at which contains extra information that is not strictly part of the data table. This is something you will commonly find in Excel, and it can make it more challenging as you always need to check to ensure that Excel is able to correctly identify the range of data you want to use when you perform operations (and exclude anything that is extraneous).

    <img src='{{ '/assets/images/K1.png' | relative_url }}' alt='Selecting all the relevant data' title='' width='898' height='214' />

 

3. **Copy** all the cells you have selected. Next, place your cursor in **column A**, at least 2 rows below the existing data Now, right-click in the selected cell, and in the **Paste Options**, choose **Transpose**.

    <img src='{{ '/assets/images/G3.PNG' | relative_url }}' alt='paste using the context menu, so the option to transpose the data is available' title='' width='727' height='528' />

4. The values will be pasted, but with the row and column headings transposed, or reversed: the years will now be across the top, and the months across the rows. This feature is very handy when you find yourself looking at some data that is very wide – i.e. it has a large number of columns, and it would be much easier to read if those headings were in the rows instead.

5. Now **select and delete** the rows containing the original data; you prefer to use your new version of the table and no longer need the original one. Select **Shift cells up** when prompted - this will move all the cells below up to fill the gap.

    <img src='{{ '/assets/images/G5.PNG' | relative_url }}' alt='Highlight the first selection again and delete' title='' width='730' height='410' />

**Tools:** [Excel](https://mdlutoronto.github.io/tutorials-search/?tool=Excel)