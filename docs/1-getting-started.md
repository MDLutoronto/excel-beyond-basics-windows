---
title: Getting Started
parent: "Excel: Beyond The Basics - Windows"
layout: default
nav_order: 1
staff: 
    - name: Kara Handren
      link: https://library.utoronto.ca/staff/kara-handren
maintainer:
    - name: Kara Handren
      link: https://library.utoronto.ca/staff/kara-handren
created_date: 2023-02-21
---

## Getting Started

1. Open the file [**Excel_TTCData_2024.xlsx**](https://uoft.me/ExcelWorkshop2025 ) in Excel (you download it alongside this PDF). You can do this either by browsing to the file on your computer and double-clicking on it, or by opening Excel from the Start menu and choosing **File ->  Open**.

2. Before we begin, let’s briefly review the Excel interface. We are currently looking at the first **worksheet** in this Excel file (Excel files are also known as **workbooks.** A workbook can contain many worksheets, and you can link data across worksheets). You can click on the tabs at the bottom of the screen to switch between worksheets.

    <img src='{{ '/assets/images/A2.png' | relative_url }}' alt='Showing of the worksheet bar across the bottom within Excel' title='' width='408' height='98' />

    | **Worksheet Naming** 
    Choose descriptive names when you name worksheets, so that it’s easy to remember what data they contain. Another good practice is to avoid whitespace - computers struggle to read whitespaces when you automate tasks, so avoiding spaces will even help when you want to refer to a cell in another sheet. Use_underscores or UseCapitalization between words instead. |

3. Across the top of the screen, you have what is known as the **ribbon**. There are a series of tabs across the top, each of which provides access to a range of tools. You’re likely very familiar with this layout, as it is similar across all Microsoft Office applications.

    <img src='{{ '/assets/images/A3_0.png' | relative_url }}' alt='Displaying the Ribbon' title='' width='974' height='135' />

4. Notice at the very top of the screen there is a smaller set of buttons. This is known as the **Quick Access Toolbar**. By default, it contains **Save, Undo** and **Redo**. Click the **drop-down menu** at the end of the toolbar. Notice you can add a range of other buttons to the toolbar.

    <img src='{{ '/assets/images/A4.png' | relative_url }}' alt='Quick access toolbar displaying the more commands button' title='' width='358' height='475' />

5. In addition, you can use the **More Commands** option to add any other tool you wish. This can be very useful if you use some tools all the time and find it frustrating switching between tabs on the ribbon. For example, imagine that you use Bubble Charts frequently, and you don’t wish to go to the Insert tab on the ribbon each time. Let’s use the **More Commands** option to add the bubble chart tool to the quick access toolbar. **Click More Commands**. The “Excel Options” window pops up. Under “Choose commands from” select **Insert Tab**. Scroll down in the list and select **Insert Scatter (X, Y) or Bubble Chart**. Click **Add** to move this to the right side of the Options box. Click **OK**.

    <img src='{{ '/assets/images/A5.png' | relative_url }}' alt='Customizing of the quick access toolbar' title='' width='824' height='678' />

6. You can now access all the scatter and bubble chart options directly from the quick access toolbar! You can do this for any tools you use regularly.

    [<img src='{{ '/assets/images/A6.png' | relative_url }}' alt='Highlighting the newly added scatter plot option to the toolbar' title='' width='233' height='311' />](https://Accessing the quick access toolbar)

7. **Right-click** on any cell in the current worksheet. The **context menu** will appear - this is another great way to access commonly used tools from the Home tab while you are working on another tab and cannot currently see the Home tab tools on the ribbon.

    <img src='{{ '/assets/images/A7.png' | relative_url }}' alt='Displaying the right-click feature known as the context menu' title='' width='346' height='527' />

8. There are also many keyboard shortcuts you can use to access popular tools. You are probably familiar with **Ctrl+C**for copy and **Ctrl+V**for paste, but there are many others as well. We’ll point out a few throughout this tutorial. A full list is available from the [Microsoft Office support website](https://support.microsoft.com/en-us/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f?ui=en-us&rs=en-us&ad=us).

9. The convention in Excel is to organize your data with observations as rows and variables as columns      
    **Rows** are identified by row numbers.     
    **Columns** are identified by column letters.   
    **Cells** are identified by the row-column combination.     
    **Ranges of cells** are identified by a colon (i.e. in the selection below **A2:D13**, means the range of cells starting at A2 and finishing at D13)        
        <img src='{{ '/assets/images/A9.png' | relative_url }}' alt='Showing the highlighted cells A2:D13' title='' width='465' height='270' />

10. Let’s look a bit at the data in the current worksheet. These letters and numbers that identify our rows and columns are indicated in a grey area above and to the left of the data, which are known as the **column headers** and **row headers.** This allows for each cell to have a unique name indicated by its column letter and row number. Click on any cell in the worksheet. The name box immediately above the data will show the cell’s name.

    <img src='{{ '/assets/images/A10.PNG' | relative_url }}' alt='After clicking on a cell, the top left display shows which cell is selected' title='' width='421' height='234' />

    | **Column Naming** 
    Try to keep your names short but descriptive (so you remember what you meant the next time you open the workbook). Excel will let you type anything as a column name, but if your data might eventually be imported into other software (such as a database or a geographic information system), then you will want to follow a few additional rules: Use letters, numbers and underscores only (no spaces or special characters). Start all names with a letter, and do not use more than 64 characters in a name. |
 

11. Immediately beside the name box is the **formula bar.** It will show you the contents of the selected cell (even if the cell does not contain a formula). When you edit data, you can edit it within the cell or in the formula bar.

    <img src='{{ '/assets/images/A11.PNG' | relative_url }}' alt='While the cell name is shown on the left, the bar on the right is displaying what is occuring within the cell to display the value it contains' title='' width='421' height='234' />

12. Next, in the current worksheet, **highlight** a few cells containing only number (numeric data) by **clicking and dragging**.

    <img src='{{ '/assets/images/A12_1.png' | relative_url }}' alt='Result of having highlighted only a few cells' title='' width='264' height='149' />

13. Have a look at the bottom bar on the screen. This is called the **status bar**. It will show you some information about your highlighted cells, including the number of cells as well as the sum and average of the values in the cells.

    <img src='{{ '/assets/images/A13_0.png' | relative_url }}' alt='In the bottom right corner of the screen information appears such as average, count, and sum' title='' width='334' height='66' />

14. You can change the details shown by right-clicking anywhere in the status bar. Notice also that the status bar contains your zoom tool which lets you zoom in and out on your view as needed.

    <img src='{{ '/assets/images/A14.png' | relative_url }}' alt='Displaying the right-click feature known as the customize status bar' title='' width='447' height='368' />

**Tools:** [Excel](https://mdlutoronto.github.io/tutorials-search/?tool=Excel)