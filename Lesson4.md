# 🚀 Lesson 4: Advanced Excel Functions for Analysis

## 🎯 Learning Objectives:

* Master lookup functions (`VLOOKUP`, `HLOOKUP`) to find and retrieve data from tables.
* Gain an introductory understanding of the more flexible `INDEX` and `MATCH` combination.
* Learn to work effectively with dates and times using specialized functions.
* Perform calculations based on single or multiple criteria using conditional functions (`SUMIF`, `COUNTIF`, `SUMIFS`, `COUNTIFS`).

---

## 📚 Topics Covered:

### 1. Lookup & Reference Functions: Finding Data

* **`VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`:** Searches for a value in the **first column** of a table array and returns the value in the same row from a specified column.

  * `lookup_value`: The value you want to search for.
  * `table_array`: The range containing the data table (must include the lookup column and the return column).
  * `col_index_num`: The column number within the `table_array` from which to return a value (1 is the first column, 2 is the second, etc.).
  * `[range_lookup]`: Optional. `TRUE` (or omitted) finds an approximate match (table must be sorted ascending by the first column). `FALSE` finds an exact match.
  * > **Common Use:** Finding a product price based on a product ID.
    >
  * > **Limitation:** Can only look up based on the *first* column and only look *right*.
    >

  ![VLOOKUP Function](./Images/Lesson4/vlookup_example.png)
  *Example of VLOOKUP retrieving product information based on product ID*
* **`HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])`:** Similar to `VLOOKUP`, but searches for a value in the **first row** of a table array and returns the value in the same column from a specified row.

  * Arguments are analogous to `VLOOKUP`, but `row_index_num` specifies the row to return from.
  * > **Common Use:** Finding data based on column headers (e.g., finding sales for a specific month when months are headers).
    >

  ![HLOOKUP Function](./Images/Lesson4/hlookup_example.png)
  *Example of HLOOKUP retrieving data based on column headers*
* **`INDEX(array, row_num, [column_num])` & `MATCH(lookup_value, lookup_array, [match_type])` (Brief Introduction):** A powerful and flexible alternative to `VLOOKUP`/`HLOOKUP`.

  * `MATCH`: Finds the *position* (row or column number) of a lookup value within a range.
    * `match_type`: `0` for exact match, `1` for less than (requires sorted ascending), `-1` for greater than (requires sorted descending).
  * `INDEX`: Returns the value of a cell at a specific row and column intersection within a given array.
  * **Combination:** Use `MATCH` to find the row number (and optionally column number) dynamically, then feed those numbers into `INDEX` to retrieve the value.
  * > **Advantage:** Can look up based on any column/row, can look left/up, more robust if columns/rows are inserted/deleted within the table.
    >

  ![INDEX MATCH Functions](./Images/Lesson4/index_match.png)
  *Example comparing INDEX-MATCH with VLOOKUP for flexible lookup operations*

### 2. Date and Time Functions: Working with Time

* **Getting Current Date/Time:**

  * `TODAY()`: Returns the current date (updates automatically).
  * `NOW()`: Returns the current date and time (updates automatically).

  ![TODAY and NOW Functions](./Images/Lesson4/today_now_functions.png)
  *Using TODAY() and NOW() to get the current date and time*
* **Creating Dates:**

  * `DATE(year, month, day)`: Creates a valid Excel date from year, month, and day numbers. *Ex: `=DATE(2024, 12, 25)`*

  ![DATE Function](./Images/Lesson4/date_function.png)
  *Using the DATE function to create dates from individual components*
* **Calculating Date Differences:**

  * `DATEDIF(start_date, end_date, unit)`: Calculates the difference between two dates in specified units ("Y" for years, "M" for months, "D" for days).
    * > **Note:** `DATEDIF` is a "hidden" function; it might not appear in autocomplete but works.
      >

  ![DATEDIF Function](./Images/Lesson4/datedif_function.png)
  *Using DATEDIF to calculate time periods between dates*
* **Extracting Date Components:**

  * `YEAR(serial_number)`: Returns the year from a date.
  * `MONTH(serial_number)`: Returns the month (1-12) from a date.
  * `DAY(serial_number)`: Returns the day of the month (1-31) from a date.

  ![Date Component Functions](./Images/Lesson4/date_components.png)
  *Using YEAR, MONTH, and DAY functions to extract components from dates*

### 3. Conditional Calculations: Summing and Counting with Criteria

* **Single Criterion:**

  * `COUNTIF(range, criteria)`: Counts the number of cells within a range that meet a given condition. *Ex: `=COUNTIF(A1:A100, ">50")` or `=COUNTIF(B1:B100, "Shipped")`*

  ![COUNTIF Function](./Images/Lesson4/countif_function.png)
  *Using COUNTIF to count values meeting specific criteria*

  * `SUMIF(range, criteria, [sum_range])`: Adds the cells specified by a given condition or criteria.
    * `range`: The range to evaluate the criteria against.
    * `criteria`: The condition (e.g., `">50"`, `"Apples"`).
    * `[sum_range]`: Optional. The actual cells to sum if different from `range`. *Ex: `=SUMIF(A1:A100, "East", B1:B100)`* (Sums values in column B where column A is "East").

  ![SUMIF Function](./Images/Lesson4/sumif_function.png)
  *Using SUMIF to sum values that meet specific criteria*
* **Multiple Criteria:**

  * `COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)`: Counts cells that meet multiple conditions across different ranges (ranges must have the same size/shape).
    * *Ex: `=COUNTIFS(A1:A100, "East", C1:C100, ">1000")`* (Counts rows where Region in A is "East" AND Sales in C are > 1000).

  ![COUNTIFS Function](./Images/Lesson4/countifs_function.png)
  *Using COUNTIFS to count values meeting multiple criteria*

  * `SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`: Adds cells that meet multiple conditions across different ranges.
    * `sum_range`: The range containing the values to sum.
    * *Ex: `=SUMIFS(B1:B100, A1:A100, "East", C1:C100, ">1000")`* (Sums values in B where Region in A is "East" AND Sales in C are > 1000).

  ![SUMIFS Function](./Images/Lesson4/sumifs_function.png)
  *Using SUMIFS to sum values meeting multiple criteria*

---

## ✨ Key Takeaways:

> * **`VLOOKUP` / `HLOOKUP`** are essential for retrieving data based on a key value, but have limitations.
> * **`INDEX` & `MATCH`** offer a more powerful and flexible lookup method.
> * Excel provides specific functions for handling **dates and times** (`TODAY`, `DATE`, `DATEDIF`, `YEAR`, etc.).
> * **`SUMIF(S)` and `COUNTIF(S)`** allow you to perform calculations based on one or multiple conditions.

---

## 🛠️ Activity: Scenario-Based Exercises

* **Goal:** Apply advanced functions to solve practical data analysis problems.
* **Setup:** You'll need two sheets:

  1. `Sales Data`: Columns like `OrderID`, `ProductID`, `SaleDate`, `Region`, `Salesperson`, `Units`, `TotalSale`.
  2. `Product Info`: Columns like `ProductID`, `ProductName`, `Category`, `UnitPrice`.

  * (Populate these sheets with sample data, ensuring some matching `ProductID`s between sheets).
* **Steps:**

  1. **🛒 `VLOOKUP` - Get Product Name:** On the `Sales Data` sheet, add a new column `ProductName`. Use `VLOOKUP` to pull the `ProductName` from the `Product Info` sheet based on the `ProductID` in the `Sales Data` sheet. Remember to use `FALSE` for an exact match and make the `table_array` reference absolute if copying down. *Example: `=VLOOKUP(B2, 'Product Info'!$A$2:$D$10, 2, FALSE)`*
  2. **📅 Date Calculations:** On the `Sales Data` sheet:
     * Add a column `SaleYear` and use the `YEAR` function to extract the year from `SaleDate`.
     * Add a column `DaysSinceSale` and use `TODAY()` and the `SaleDate` to calculate how many days have passed since each sale. *Example: `=TODAY()-C2`* (Format the result as a Number).
     * *(Optional Challenge)* If you have an `OrderDate` and `ShipDate`, use `DATEDIF` to calculate the shipping time in days.
  3. **📊 `SUMIF` - Regional Sales:** In a separate area or sheet, create a small summary table with unique Regions listed. Use `SUMIF` to calculate the `TotalSale` for each `Region` from the `Sales Data` sheet.
  4. **👤 `COUNTIF` - Salesperson Orders:** In the summary area, use `COUNTIF` to count how many orders each `Salesperson` handled.
  5. **📈 `SUMIFS` - Specific Product Sales by Region:** Calculate the total `TotalSale` for a *specific* `ProductName` (e.g., "Laptop") within *each* `Region`. You'll need `SUMIFS` because you have two criteria (ProductName and Region).
  6. **🔢 `COUNTIFS` - High Value Orders by Salesperson:** Count how many orders *each* `Salesperson` had where the `TotalSale` was greater than a certain amount (e.g., $500). Use `COUNTIFS`.
  7. **(Optional) `INDEX/MATCH`:** Repeat step 1 (Get Product Name) using `INDEX` and `MATCH` instead of `VLOOKUP`.
     * *Hint:* `MATCH(B2, 'Product Info'!$A$2:$A$10, 0)` finds the row number. `INDEX('Product Info'!$B$2:$B$10, [result_of_MATCH])` gets the name.
  8. 💾 **Save:** Save your workbook.

---
