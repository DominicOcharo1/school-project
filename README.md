Worked for 1m 18s

We will create a new Excel sheet called Feedback Dashboard that summarizes feedback requests by client, shows how many requests are Completed, In Progress, or Requested, and includes charts to make the information easy to understand.

I will explain every step, from opening your worksheet to creating the final dashboard.

### What your finished Excel dashboard will contain

### FEEDBACK REQUEST DASHBOARD

Summary by Client / Project

Completed

### 125

Illustrative count

In Progress

### 48

Illustrative count

Requested

### 32

Illustrative count

### 1\. Feedback requests by client

A stacked bar chart comparing the three statuses for each client or project.

### 2\. Overall feedback status distribution

A doughnut chart showing the proportion of requests in each status.

The numbers and charts above are examples to show the intended layout. They are not calculated from your workbook.

Your real dashboard will use the figures from your sheet named `HC_Report_Extract_202`.

### Part 1: Understand the columns we will use

Your original sheet has data in columns A to O. We will use the following three columns and add one helper column.

| Column | Header                     | Purpose                                                            |
| ------ | -------------------------- | ------------------------------------------------------------------ |
| H      | Status of Feedback request | Tells us whether a request is Completed, In Progress, or Requested |
| L      | Project Name               | Used when Client Name is null                                      |
| N      | Client Name                | Used as the primary name                                           |
| P      | Client / Project           | A new helper column that decides which name to use                 |

The most important rule is:

* If Client Name contains a real name, use that name.

* If Client Name is blank or contains the word `null`, use Project Name instead.

For example:

| Client Name (N) | Project Name (L) | Client / Project (P) |
| --------------- | ---------------- | -------------------- |
| ABC Limited     | Website Upgrade  | ABC Limited          |
| null            | Payroll System   | Payroll System       |
| XYZ Company     | Network Upgrade  | XYZ Company          |
| (blank)         | Mobile App       | Mobile App           |

We need the helper column because a PivotTable needs one consistent field to group records by the correct name.

### Part 2: Prepare your original data

### Step 1: Open your Excel file

* Open your Excel workbook.

* Look at the sheet tabs at the bottom of Excel.

* Find and click the tab named `HC_Report_Extract_202`.

* You should now see your original data.

* Press Ctrl + S on your keyboard to save the workbook.

For additional protection, you can also use File → Save As to create a backup copy before making changes.

### Step 2: Check that the column headings are correct

Look at the first row of your data.

Check these cells:

* Click cell `H1`. It should say `Status of Feedback request`.

* Click cell `L1`. It should say `Project Name`.

* Click cell `N1`. It should say `Client Name`.

The instructions below assume that row 1 contains your column headings and row 2 contains your first data record.

Do not delete any of your original records.

### Step 3: Convert your data into an Excel Table

An Excel Table helps formulas fill down automatically and makes your PivotTables easier to update when new records are added.

First, click any cell containing data on the `HC_Report_Extract_202` sheet.

Now follow these steps:

* Press Ctrl + T on your keyboard.

* A small window called Create Table should appear.

* Look at the proposed cell range. It should cover your original data from columns A to O.

* Make sure the checkbox labelled My table has headers is ticked.

* Click OK.

Your data should now have filter arrows beside the column headings, and the rows may have alternating colours.

Important: If your data already has these features and is already an Excel Table, do not create another one. Just click inside it and continue.

### Step 4: Give the table a name

This makes it easier to identify the correct source data later.

* Click any cell inside the table.

* Look at the top ribbon in Excel.

* Click the Table Design tab. In some versions, it may simply say Table or Table Design under Table Tools.

* Find the box labelled Table Name, usually near the left side of the ribbon.

* Click inside that box.

* Delete the existing name.

* Type:

`FeedbackData`

8\. Press Enter.

Your source table is now called `FeedbackData`.

### Part 3: Create the Client / Project helper column

### Step 5: Add a new column in P

Go back to the worksheet `HC_Report_Extract_202`, if you are not already there.

* Look at the column letters along the top: A, B, C, and so on.

* Find column P, immediately to the right of column O.

* Click cell `P1`.

* Type the following heading:

`Client / Project`

5\. Press Enter.

If Excel does not automatically extend your existing table to include column P, we will correct that in Step 7.

### Step 6: Enter the formula

Click cell `P2`.

Copy the formula below and paste it into the formula bar or directly into the cell.

Formula for cell P2

=IF(OR(TRIM(N2&"")="",LOWER(TRIM(N2&""))="null"),TRIM(L2&""),TRIM(N2&""))

Copy formula

After pasting the formula:

* Press Enter.

* Look at cell P2.

* Check whether the result is the client name or project name, depending on the contents of columns N and L.

### What does the formula do?

It examines the value in `N2`, which is the Client Name.

If the client name is blank or contains `null`, it takes the project name from `L2`. Otherwise, it uses the client name from `N2`.

It also handles entries such as `NULL`, `Null`, and names with accidental spaces around them.

For example, if N2 contains `null` and L2 contains `Payroll System`, P2 should display `Payroll System`.

Note: This formula assumes your Excel uses commas between formula arguments. If Excel rejects it because of your regional settings, replace the commas separating the arguments with semicolons.

### Step 7: Make sure the formula is applied to every record

This is very important. We do not want to calculate the correct name only for the first row.

### If your data is an Excel Table

After entering the formula in P2, Excel will normally fill the formula down the entire column automatically.

* Scroll down column P.

* Check that the formula has produced a result for the other records.

* If the results are filled down, continue to Step 8.

### If only P2 contains a result

Check the following:

* Click cell P2.

* Look at the Table Design tab.

* If column P is outside the table, choose Table Design → Resize Table.

* In the range box, change the table's final column from O to P, keeping the correct first and last row numbers. For example, if your final row is 528, the range would end with `$P$528`.

* Click OK.

Then enter the formula in the first data cell under `Client / Project` again. Excel should fill the calculated column down the table.

If your data is not a Table, convert the full range, including column P, into a Table using Ctrl + T, and ensure My table has headers is ticked.

### Step 8: Check the three important columns

Before creating PivotTables, confirm that:

* Column H contains the feedback status.

* Column L contains Project Name.

* Column N contains Client Name.

* Column P contains the correct client or project name for every record.

For the rest of this exercise, column P is the field that we will use to group the requests.

### Part 4: Create your new dashboard worksheet

### Step 9: Add a new sheet

At the bottom of Excel, find the + sign next to your worksheet tabs.

* Click the + sign.

* Excel will create a new worksheet.

* Double-click the new sheet's tab.

* Type `Feedback Dashboard`.

* Press Enter.

You should now have two sheets:

* `HC_Report_Extract_202` — your original data.

* `Feedback Dashboard` — your new summary and charts.

We will leave the original data in place and build the report on the new sheet.

### Part 5: Create the first PivotTable — summary by client

This is the main summary table. It will show the number of Completed, In Progress, and Requested feedback requests for each client or project.

### Step 10: Start creating the PivotTable

* Click the `HC_Report_Extract_202` sheet.

* Click any cell inside the `FeedbackData` table.

* On the ribbon at the top, click Insert.

* Click PivotTable or From Table/Range, depending on your Excel version.

* A window called Create PivotTable will open.

* Make sure the table or range shows `FeedbackData` or the corresponding range covering your data, including column P.

* Choose New Worksheet.

* Click OK.

Excel creates a new worksheet containing an empty PivotTable and opens the PivotTable Fields panel.

### Step 11: Rename the new PivotTable sheet

At the bottom of Excel, find the newly created sheet tab.

* Double-click its name.

* Type `Feedback Dashboard`.

* Press Enter.

* If you already created a sheet with this name, use that existing sheet as the destination instead. Excel does not permit two worksheets with identical names.

We want the first PivotTable near the top-left of this dashboard. Excel usually starts it around cell A3 on a new PivotTable sheet.

### Step 12: Understand the PivotTable Fields panel

On the right-hand side, you should see a list of field names from your original data, including:

* Status of Feedback request

* Project Name

* Client Name

* Client / Project

Below those field names, there are four boxes:

| Box     | Purpose                                    |
| ------- | ------------------------------------------ |
| Filters | Filters the whole PivotTable               |
| Columns | Creates separate columns for categories    |
| Rows    | Creates one row for each client or project |
| Values  | Calculates counts or other numbers         |

We will drag fields into these boxes.

### Step 13: Add Client / Project to Rows

* In the PivotTable Fields panel, find `Client / Project`.

* Click and hold it with your mouse.

* Drag it into the box labelled Rows.

* Release your mouse.

Your PivotTable should now list client or project names down the left side.

### Step 14: Add the status field to Columns

* Find `Status of Feedback request`.

* Click and hold it.

* Drag it into the box labelled Columns.

* Release your mouse.

Your PivotTable should now have separate columns for statuses such as Completed, In Progress, and Requested.

### Step 15: Add a count to Values

We now need Excel to count the number of feedback records.

* Find `Client / Project` again in the field list.

* Drag it a second time into the Values box.

* Excel should create a field called something like Count of Client / Project.

It is okay to use the same field twice: once for the rows and once for the count.

The Values box should use Count, not Sum. If Excel has selected a different calculation:

* Click the arrow beside the field in the Values box.

* Choose Value Field Settings.

* Select Count.

* Click OK.

Why are we counting `Client / Project` instead of the status field? It allows Excel to count each record with a client/project name, including records where the status might be blank. A missing status will appear in a separate blank category rather than silently disappearing.

### Step 16: Check the result

Your first PivotTable should look similar to this example:

| Client / Project | Completed | In Progress | Requested | Grand Total |
| ---------------- | --------- | ----------- | --------- | ----------- |
| ABC Limited      | 12        | 4           | 3         | 19          |
| XYZ Company      | 8         | 6           | 2         | 16          |
| Payroll System   | 5         | 2           | 7         | 14          |
| Grand Total      | 25        | 12          | 12        | 49          |

These figures are examples only.

Your actual numbers will depend on the records in your workbook.

Each row represents a client or project, and the counts show how many feedback records are in each status.

If you see a `(blank)` column, some records may have missing status values. Check the original status column rather than assuming the blank records belong to one of the three statuses.

### Part 6: Format the main PivotTable

### Step 17: Make it easy to read

* Click any cell inside the first PivotTable.

* Look for the Design tab under PivotTable Tools.

* Click Design.

* Find the PivotTable Styles section.

* Select a style that has a clear header and alternating row colours.

* Make sure the Grand Total row remains visible.

You can also adjust the column widths by double-clicking the boundary between column letters at the top of Excel.

For example, double-click the boundary between columns A and B to automatically fit column A to the client names.

### Step 18: Add a heading to the dashboard

Click cell `A1` on the `Feedback Dashboard` sheet, above the PivotTable, and type:

FEEDBACK REQUEST DASHBOARD

Then:

* Press Enter.

* Select A1.

* Click Home.

* Make the font bold.

* Increase the font size to approximately 18 or 20.

* Choose a dark blue font colour if you like.

In A2, type:

`Summary of feedback requests by client or project`

This explains what the sheet contains.

Keep the heading above the PivotTable and do not type over any cells containing PivotTable data.

### Part 7: Create the first chart — feedback requests by client

This chart will show how the three feedback statuses are distributed for each client or project.

### Step 19: Insert a PivotChart

* Click any cell inside your first PivotTable.

* Click the PivotTable Analyze tab. In some versions, it may simply say Analyze.

* Click PivotChart.

If you cannot find PivotChart there, click the Insert tab and look for the PivotChart option.

* A chart selection window will appear.

* Select Bar from the chart categories.

* Choose Stacked Bar.

* Click OK.

Excel will create a chart linked to the client-summary PivotTable.

A stacked bar is useful because each client gets one horizontal bar, divided into the three feedback statuses.

### Step 20: Move the chart to the right side

The chart may appear on top of the PivotTable. That is normal.

To move it:

* Click the outer border of the chart, not one of the bars.

* Hold down your mouse button.

* Drag the chart towards the right-hand side of the worksheet, starting around column K.

* Release the mouse button.

* Resize it by dragging one of the corner handles until the chart is large enough to read.

Try to keep the first PivotTable on the left and the chart on the right.

If you have many clients, make the chart taller so the labels are easier to read.

### Step 21: Give the chart a title

* Click the chart once.

* Click the chart title.

* Replace the title with:

`Feedback Requests by Client / Project`

If the chart does not have a title, click the chart, go to Chart Design → Add Chart Element → Chart Title → Above Chart, and then enter the title.

### Step 22: Check the legend

Your chart should identify the three feedback statuses:

* Completed

* In Progress

* Requested

If a status is missing from the legend, check that the original data contains that status and that the field has been placed in the PivotTable's Columns box.

The bar lengths should change according to the number of feedback requests associated with each client or project.

### Part 8: Create a second PivotTable — overall status distribution

The first PivotTable answers this question:

How many feedback requests does each client or project have in each status?

The second PivotTable answers a different question:

How many feedback requests are Completed, In Progress, and Requested across the entire dataset?

### Step 23: Start another PivotTable

* Click the `HC_Report_Extract_202` sheet.

* Click any cell inside the `FeedbackData` table.

* Click Insert → PivotTable.

* Confirm that the source is `FeedbackData`.

* Select Existing Worksheet.

* Click inside the Location box.

* Enter the location `'Feedback Dashboard'!$H$3`.

* Click OK.

This creates the second PivotTable on the dashboard, starting around cell H3, away from the first PivotTable.

If your Excel version does not accept the typed location, click the small range-selection button beside the Location box, choose the `Feedback Dashboard` tab and cell H3, and confirm the selection.

### Step 24: Set up the second PivotTable

In the PivotTable Fields panel:

* Drag `Status of Feedback request` into Rows.

* Drag `Client / Project` into Values.

* Check that the Values field uses Count.

Your second PivotTable should look similar to this:

| Status of Feedback request | Count of Client / Project |
| -------------------------- | ------------------------- |
| Completed                  | 125                       |
| In Progress                | 48                        |
| Requested                  | 32                        |
| Grand Total                | 205                       |

Again, these are illustrative figures rather than your actual data.

This table gives you the totals for the entire dataset, regardless of which client or project the records belong to.

### Step 25: Check for blank categories

Look at the rows of the second PivotTable.

If you see `(blank)` as one of the statuses, it means there are records whose Status of Feedback request is empty.

Since you expect only three statuses, you should investigate those missing values in the original sheet. For the chart, you can filter out `(blank)` to display only the three requested statuses, but do not lose track of the records that need correcting.

### Part 9: Create a doughnut chart showing the overall distribution

### Step 26: Insert the chart

* Click any cell inside the second PivotTable.

* Click PivotTable Analyze → PivotChart.

* Choose Pie from the chart categories.

* Select Doughnut.

* Click OK.

Excel will create a doughnut chart linked to the second PivotTable.

If a blank category appears, filter it from the chart using the status field button, or correct the source data as appropriate.

### Step 27: Move the chart

* Click the outside border of the doughnut chart.

* Drag it to the right side of the sheet, below the first chart, starting around column K.

* Resize it using the corner handles.

Keep the doughnut chart separate from the first chart so that the two visuals do not cover one another.

### Step 28: Change the title

Click the chart title and replace it with:

`Overall Feedback Status Distribution`

If there is no title, use Chart Design → Add Chart Element → Chart Title → Above Chart.

### Step 29: Add labels and percentages

This step makes the chart much easier to understand.

* Click the doughnut chart.

* Click the small + button next to the chart, or use Chart Design → Add Chart Element.

* Select Data Labels.

* Open the data-label options, which may be accessed by clicking the arrow beside Data Labels or choosing More Options.

* In the Format Data Labels panel, select Category Name and Percentage.

* If you prefer a less crowded chart, deselect Value so the chart shows the status name and percentage rather than all three label components.

Now each slice should be identifiable as Completed, In Progress, or Requested, with its share of the total.

The percentage for a status is:

\\text{Percentage}=
\\frac{\\text{Number of requests in that status}}
{\\text{Total requests}}\\times100

Excel calculates the displayed percentages from the chart data.

### Part 10: Arrange and finish your dashboard

### Step 30: Arrange the worksheet

Your dashboard should have a layout similar to this:

### FEEDBACK REQUEST DASHBOARD

Client and project feedback tracking

### PivotTable 1

Client / Project and status counts

Completed | In Progress | Requested | Total

### PivotTable 2

Overall status totals

Completed | In Progress | Requested

### Bar Chart

Feedback requests by client or project

### Doughnut Chart

Overall distribution by status

Suggested arrangement, not a screenshot of your workbook.

To finish the layout:

* Keep the main client-summary PivotTable on the left.

* Keep the overall status PivotTable nearby, under or beside it.

* Position the stacked bar chart and doughnut chart on the right.

* Make sure all chart titles and labels are readable.

* If charts cover the PivotTables, drag them to an empty area.

* Optionally, click View and untick Gridlines for a cleaner dashboard appearance.

* Press Ctrl + S to save the workbook.

### Step 31: Make sure everything is counting correctly

Perform these checks before relying on the report.

| Check                   | What you should see                                                   |
| ----------------------- | --------------------------------------------------------------------- |
| Client Name is not null | The client's name appears in Client / Project                         |
| Client Name is null     | The Project Name appears in Client / Project                          |
| Completed records       | Counted under Completed                                               |
| In Progress records     | Counted under In Progress                                             |
| Requested records       | Counted under Requested                                               |
| Grand Total             | Total of the three statuses, plus any blank-status records if present |

The sum of the three status counts should equal the overall count of feedback requests when every request has one of those three statuses.

If the totals do not match, check for blank statuses, blank client and project names, or other unexpected data values.

### Part 11: How to update the dashboard when your data changes

This is the final and important step. You do not need to rebuild the report every time the source data changes.

When you add new feedback records:

* Go to `HC_Report_Extract_202`.

* Add the new records inside the `FeedbackData` table, at the bottom.

* Enter the Client Name, Project Name and Status of Feedback request.

* Check that the Client / Project formula fills down for the new row.

* Return to `Feedback Dashboard`.

* Click any cell inside one of the PivotTables.

* Click PivotTable Analyze → Refresh. Alternatively, go to Data → Refresh All to refresh both PivotTables.

* Check that the tables and charts reflect the updated data.

* Press Ctrl + S to save.

The charts are linked to the PivotTables, so refreshing the PivotTables updates the chart data as well.

### Your final checklist

### Dashboard completion checklist

0 of 8

Original data is still intact on HC_Report_Extract_202

Column P contains the correct Client / Project values

The source table includes column P and is named FeedbackData

Feedback Dashboard contains the client-summary PivotTable

The overall status PivotTable shows Completed, In Progress and Requested

The stacked bar chart compares statuses by client/project

The doughnut chart shows the overall status distribution

Grand totals have been checked and the workbook has been saved

The key to getting the right result is column P. Once that helper column is correct, both PivotTables and both charts will group requests by the actual Client Name when available, or the Project Name when the client name is null.
