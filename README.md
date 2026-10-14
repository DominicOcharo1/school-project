Worked for 3m 17s

The best way to do this is to create a new worksheet called Q1 Feedback Coverage. It will compare each employee’s total engagement hours against the feedback records in `HC_Report_Extract_202`, identify employees who meet the 37.5-hour threshold, and flag anyone who does not have a feedback record.

You do not need to change your existing `Feedback Dashboard` or the `Client / Project` helper column you added in AS.

### 1\. What the additional analysis will show

Your new worksheet will contain a table like this:

| Employee Name (Q1) | Employee Name (HC) | Total Hours | Feedback Records | Coverage Result  |
| ------------------ | ------------------ | ----------- | ---------------- | ---------------- |
| Sidialo, Sharon    | Sharon Sidialo     | 45.0        | 0                | MISSING FEEDBACK |
| Kakae, Jackson     | Jackson Kakae      | 217.5       | 0                | MISSING FEEDBACK |
| Kinyanjui, Grace   | Grace Kinyanjui    | 248.0       | 1                | Feedback found   |
| Owino, Beatrice    | Beatrice Owino     | 37.5        | 1                | Feedback found   |

These examples illustrate the intended layout. The feedback counts must come from your workbook, not be entered manually.

The important rules are:

* Add up each employee’s hours using the individual project rows in `Q1 Engagement`.

* Exclude the rows labelled, for example, `Sidialo, Sharon Total`, because those are already subtotal rows.

* Include employees whose total is 37.5 hours or more.

* Reverse the employee name format from `Last, First` to `First Last` before comparing it with column B in `HC_Report_Extract_202`.

* If an eligible employee has no matching feedback record, flag them as `MISSING FEEDBACK`.

I will use 37.5 hours or more (≥37.5) because that is the threshold specified in your final requirement. An employee with exactly 37.5 hours must therefore be included.

### 2\. Step 1 — Create the new worksheet

Create a separate sheet

At the bottom of Excel, click the `+` icon beside the sheet tabs. Double-click the new sheet tab and rename it:

`Q1 Feedback Coverage`

Press Enter.

Prepare the heading and table

Enter `Q1 ENGAGEMENT — FEEDBACK COVERAGE ANALYSIS` in cell A1.

In A2, enter:

`Checks whether employees with at least 37.5 engagement hours have a feedback record.`

Then enter the following column headings in row 4.

| Cell | Heading                   |
| ---- | ------------------------- |
| A4   | Employee Name (Q1)        |
| B4   | Employee Name (HC Format) |
| C4   | Total Q1 Hours            |
| D4   | Threshold Met?            |
| E4   | Feedback Record Count     |
| F4   | Coverage Result           |
| G4   | Completed Records         |
| H4   | In Progress Records       |
| I4   | Requested Records         |

The table starts at row 5, leaving rows 1–3 available for the title and explanation.

### 3\. Step 2 — Generate the list of employees without the Total rows

Go to cell A5 in the new sheet.

Copy and paste this formula:

```excel
=SORT(UNIQUE(FILTER('Q1 Engagement'!$A$3:$A$504,('Q1 Engagement'!$A$3:$A$504<>"")*ISERROR(SEARCH(" Total",'Q1 Engagement'!$A$3:$A$504)))))
```

Press Enter.

This formula does three things:

* Reads the employee names in column A of `Q1 Engagement`.

* Excludes blank cells and rows containing ` Total`, so the subtotal lines are not treated as employees.

* Removes duplicate names and sorts the remaining names alphabetically.

Your employee list should now spill down automatically from A5. Do not type into the cells underneath the formula, because Excel needs those cells available for the results.

This formula requires a version of Excel that supports `SORT`, `UNIQUE` and `FILTER`, such as Microsoft 365.

### 4\. Step 3 — Convert the employee names into the HC format

This is one of the most important steps.

Your Q1 sheet contains names such as:

`Sidialo, Sharon`

But column B in `HC_Report_Extract_202` contains:

`Sharon Sidialo`

We need Excel to convert each name before checking whether feedback exists.

Click cell B5 and enter:

```excel
=IF(A5="","",IFERROR(TRIM(MID(A5,FIND(",",A5)+1,LEN(A5))&" "&LEFT(A5,FIND(",",A5)-1)),"CHECK NAME FORMAT"))
```

Press Enter.

The formula takes the text after the comma and puts it first, followed by the text before the comma.

For example:

| Original name    | Converted name  |
| ---------------- | --------------- |
| Sidialo, Sharon  | Sharon Sidialo  |
| Kakae, Jackson   | Jackson Kakae   |
| Kinyanjui, Grace | Grace Kinyanjui |
| Owino, Beatrice  | Beatrice Owino  |

The formula also trims unnecessary spaces and flags any name it cannot convert as `CHECK NAME FORMAT`.

### 5\. Step 4 — Calculate total hours for each employee

Click cell C5 and enter:

```excel
=IF(A5="","",SUMIF('Q1 Engagement'!$A$3:$A$504,A5,'Q1 Engagement'!$E$3:$E$504))
```

Press Enter.

This formula adds the hours in column E of `Q1 Engagement` for every project that matches the employee in column A.

Why this avoids double counting: the formula looks for an exact employee name, such as `Sidialo, Sharon`. The separate row called `Sidialo, Sharon Total` does not match that exact name, so its subtotal is excluded.

For example, Sharon's two projects are 22.5 hours each. The formula returns:

22.5 + 22.5 = 45 hours.

Do not add the existing subtotal row to this calculation. The formula already calculates the total directly from the project records.

### 6\. Step 5 — Check whether the employee meets the threshold

Click D5 and enter:

```excel
=IF(A5="","",IF(C5>=37.5,"Yes","No"))
```

Press Enter.

Employees with 37.5 hours or more will display `Yes`. Employees with fewer than 37.5 hours will display `No`.

### 7\. Step 6 — Check whether a feedback record exists

Now we compare the converted employee names against column B in `HC_Report_Extract_202`.

Click E5 and enter:

```excel
=IF(A5="","",COUNTIF('HC_Report_Extract_202'!$B:$B,B5))
```

Press Enter.

This counts how many records in `HC_Report_Extract_202` contain the employee's converted name.

* A result of `0` means no matching employee-name record was found.

* A result of `1` means one feedback record was found.

* A result greater than `1` means the employee appears in multiple feedback records.

We count records, rather than simply checking whether the employee appears once, because an employee may have received feedback from multiple people or on multiple engagements.

### 8\. Step 7 — Automatically flag employees who need follow-up

Click F5 and enter:

```excel
=IF(A5="","",IF(C5<37.5,"Below threshold",IF(E5>0,"Feedback found","MISSING FEEDBACK")))
```

Press Enter.

This is the main result of your analysis.

| Condition                                                    | Coverage Result  |
| ------------------------------------------------------------ | ---------------- |
| Total hours below 37.5                                       | Below threshold  |
| Total hours at least 37.5, with one or more feedback records | Feedback found   |
| Total hours at least 37.5, with zero feedback records        | MISSING FEEDBACK |

This makes the analysis easy to interpret: every eligible employee without a feedback record is clearly identified.

### 9\. Step 8 — Show the feedback statuses for each employee

These three optional columns add useful detail. They show whether the employee's existing feedback records are Completed, In Progress or Requested.

In G5, enter:

```excel
=IF(A5="","",COUNTIFS('HC_Report_Extract_202'!$B:$B,B5,'HC_Report_Extract_202'!$H:$H,"Completed"))
```

In H5, enter:

```excel
=IF(A5="","",COUNTIFS('HC_Report_Extract_202'!$B:$B,B5,'HC_Report_Extract_202'!$H:$H,"In Progress"))
```

In I5, enter:

```excel
=IF(A5="","",COUNTIFS('HC_Report_Extract_202'!$B:$B,B5,'HC_Report_Extract_202'!$H:$H,"Requested"))
```

Press Enter after each formula.

These columns use the actual feedback status in column H of your original sheet. They count feedback records for the employee, so their totals may exceed one when an employee has multiple records.

### 10\. Step 9 — Fill all the formulas down

You have now entered formulas in B5 through I5.

To apply them to the rest of the employee list:

* Select the range `B5:I506`.

* Press Ctrl + D on your keyboard to fill the formulas down from row 5.

* Check that the formulas return values alongside the employee names in column A.

The range extends beyond the current number of employees so that there is room for all names in your present source range. The formulas will return blank results for rows where column A has no employee name.

Do not fill down over A5, because that cell contains the dynamic employee-list formula.

### 11\. Step 10 — Add a summary showing the overall compliance position

A detailed employee list is useful, but a summary will make the findings easier to communicate.

Enter these headings:

* K4: `Summary Measure`

* L4: `Result`

Then enter the following labels and formulas:

| Cell | Summary measure                     | Formula in column L                      |
| ---- | ----------------------------------- | ---------------------------------------- |
| K5   | Total distinct employees            | =COUNTA(A5#)                             |
| K6   | Employees at or above 37.5 hours    | =COUNTIF(C5:C506,">=37.5")               |
| K7   | Eligible employees with feedback    | =COUNTIFS(C5:C506,">=37.5",E5:E506,">0") |
| K8   | Eligible employees missing feedback | =COUNTIF(F5:F506,"MISSING FEEDBACK")     |
| K9   | Feedback coverage rate              | =IFERROR(L7/(L7+L8),0)                   |

For example, enter the formula for K5's result in L5, the formula for K6's result in L6, and so on.

Format L9 as a percentage using Home → Number → Percentage.

The coverage rate indicates what proportion of employees who meet the hours threshold have at least one matching feedback record.

### 12\. Step 11 — Format the analysis so the exceptions stand out

Make the sheet professional and easy to read.

* Select A1 and make the title bold, with a larger font.

* Format A4:I4 as a coloured header row with bold, white text.

* Format C5:C506 as a number with one decimal place.

* Apply filters by selecting `A4:I506` and choosing Data → Filter.

* Select column F and use Home → Conditional Formatting → Highlight Cells Rules → Text that Contains to highlight `MISSING FEEDBACK` in red.

* Optionally, highlight `Feedback found` in green and `Below threshold` in a neutral colour.

The filter on column F is particularly useful. Choose `MISSING FEEDBACK` to see only the employees who meet the hours threshold but do not have a matching feedback record.

You can also create a simple bar chart using the summary measures for eligible employees with feedback versus eligible employees missing feedback. This provides a quick visual of the compliance gap.

### 13\. What your current workbook shows

I checked the data in your attached workbook and compared the names after converting the Q1 name format to the HC format.

Distinct Q1 employees

### 108

Meet ≥37.5 hours

### 107

Eligible employees with feedback

### 34

Eligible employees with no match

### 73

Current matching coverage

### 31.8%

Based on the currently supplied data, using the ≥37.5-hour threshold and employee-name matching.

This means 73 of the 107 eligible employees currently have no matching employee-name record in `HC_Report_Extract_202`. Those are the records to investigate first.

There is also one employee below the threshold in the current Q1 data: Abel Legesse, with 15 hours. Beatrice Owino has exactly 37.5 hours and is included in the eligible group.

Before treating the 73 unmatched employees as confirmed missing feedback, verify any differences in spelling, name order or extra spaces. A name mismatch can produce a zero count even when a corresponding feedback record exists under a slightly different name.

Final recommendation: keep `Feedback Dashboard` for the existing client/project and status analysis, and use `Q1 Feedback Coverage` for this separate employee-level control. Together, the two sheets will show both the status of feedback requests and whether employees who have reached the required engagement hours have a feedback record.
