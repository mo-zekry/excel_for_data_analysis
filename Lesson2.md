# 🧹 Lesson 2: Data Importing and Preparation

**(⏳ Estimated Duration: 1.5 hours)**

---

## 🎯 Learning Objectives:

*   Learn how to import data into Excel from various common sources.
*   Identify and handle common data issues like missing or incorrect values.
*   Master techniques for sorting and filtering data to isolate relevant information.
*   Apply essential data cleaning tools like Text-to-Columns and Remove Duplicates.

---

## 📚 Topics Covered:

### 1. Importing Data: Getting Data into Excel

*   **Sources:** Learn the methods for bringing external data into your worksheet:
    *   📄 **CSV (Comma Separated Values) & Text Files (.txt):** Use the `Data` tab > `Get & Transform Data` > `From Text/CSV` wizard. Understand delimiters (comma, tab, semicolon) and data type detection.

    ![Importing CSV Files](./Images/Lesson2/csv_import.png)
    *The CSV import process showing the Get & Transform Data options*

    *   🗄️ **Databases (Basic Introduction):** Briefly introduce connecting to simple databases (like Access or SQL Server) via `Data` tab > `Get Data` > `From Database` (mentioning this might require specific drivers or permissions).
    *   🌐 **Web (Mention):** Briefly mention the possibility of importing data from web pages (`Data` tab > `Get Data` > `From Web`).

    ![Database and Web Import](./Images/Lesson2/database_web_import.png)
    *Options for importing data from databases and web sources*

*   **The Import Process:** Understand the steps involved, including previewing data, choosing load destinations (new worksheet vs. existing), and basic transformation options available during import.

    ![Import Wizard](./Images/Lesson2/import_wizard.png)
    *The Excel data import wizard showing preview and transformation options*

### 2. Data Quality: Dealing with Imperfections

*   **Identifying Issues:** Learn visual scanning techniques and tools to spot problems:
    *   **Missing Data:** Recognizing blank cells where data should exist.
    *   **Incorrect Data:** Spotting typos, inconsistent formatting (e.g., dates as text), outliers, or values that don't make sense in context.

    ![Data Quality Issues](./Images/Lesson2/data_quality_issues.png)
    *Examples of common data quality issues like missing values and inconsistent formatting*

*   **Handling Strategies:**
    *   **Missing Values:** Discuss options like deleting rows/columns (use with caution!), filling with a specific value (0, "N/A", mean, median - depending on context), or using more advanced imputation techniques (mention briefly).
    *   **Incorrect Values:** Correcting typos manually, using Find & Replace (`Ctrl+H`), or applying consistent formatting.

    ![Handling Missing Data](./Images/Lesson2/handling_missing_data.png)
    *Strategies for dealing with missing values in Excel*

### 3. Organizing Data: Sorting & Filtering

*   **Sorting:** Arranging data in a meaningful order:
    *   **Single Column Sort:** Use the `Data` tab > `Sort & Filter` > `A-Z` (ascending) or `Z-A` (descending) buttons.
    *   **Multi-Level Sort:** Use the `Sort` dialog box (`Data` tab > `Sort`) to sort by multiple columns sequentially (e.g., sort by Region, then by Sales). Understand sorting options (values, cell color, font color).

    ![Sorting Data](./Images/Lesson2/sorting_data.png)
    *Single-column and multi-level sorting options in Excel*

*   **Filtering:** Displaying only the data that meets specific criteria:
    *   **AutoFilter:** Enable filter dropdown arrows on headers (`Data` tab > `Filter` or `Ctrl+Shift+L`).
    *   **Using Filters:** Filter by specific values, text criteria (contains, begins with), number criteria (greater than, between), date criteria (this month, next year), or cell color/font color.
    *   **Clearing Filters:** Removing filters to show all data again.

    ![Filtering Data](./Images/Lesson2/filtering_data.png)
    *Using AutoFilter to display only specific data based on criteria*

### 4. Basic Data Cleaning Tools

*   **Text-to-Columns:** Splitting data from a single column into multiple columns:
    *   Located on the `Data` tab > `Data Tools` group.
    *   Use cases: Separating full names into first and last names, splitting delimited data pasted into a single column.
    *   Understand `Delimited` (using characters like commas, spaces, tabs) vs. `Fixed Width` options.

    ![Text to Columns](./Images/Lesson2/text_to_columns.png)
    *Using Text to Columns to split full names into first and last names*

*   **Removing Duplicates:** Identifying and deleting entire rows that are identical based on selected columns:
    *   Located on the `Data` tab > `Data Tools` group.
    *   Select the data range, choose which columns to check for duplicates.
    *   > **Caution:** This permanently deletes rows. Consider copying data first if unsure.

    ![Remove Duplicates](./Images/Lesson2/remove_duplicates.png)
    *The Remove Duplicates dialog allowing selection of columns to check*

*   **Find & Replace (`Ctrl+H`):** Useful for correcting consistent errors or standardizing terms (e.g., replacing "USA" with "United States").

    ![Find and Replace](./Images/Lesson2/find_replace.png)
    *Using Find and Replace to standardize text values across a dataset*

---

## ✨ Key Takeaways:

> *   Excel can import data from various sources like **CSV, Text files**, and databases using the **Get & Transform Data** tools.
> *   Data often requires **cleaning** to handle **missing or incorrect** values before analysis.
> *   **Sorting** arranges data logically; **Filtering** isolates specific subsets of data.
> *   Tools like **Text-to-Columns** and **Remove Duplicates** are essential for basic data preparation.

---

## 🛠️ Activity: Import and Clean a Dataset

*   **Goal:** Import a sample dataset (e.g., a provided CSV file) and perform basic cleaning and organization tasks.
*   **Dataset:** (You'll need to provide a sample CSV file for this activity, e.g., `sample_customer_data.csv` with columns like `CustomerID`, `FullName`, `Email`, `JoinDate`, `City State Zip`, `LastPurchaseAmount`, possibly with some blank cells or inconsistent state abbreviations).
*   **Steps:**
    1.  📥 **Import:** Use `Data` > `From Text/CSV` to import the `sample_customer_data.csv` file into a new Excel worksheet. Review the preview and ensure data types look correct before loading.
    2.  👀 **Inspect:** Scan the imported data for obvious issues (blanks, weird values).
    3.  🗑️ **Handle Missing:** Find any rows with missing `Email` addresses. Decide on a strategy (e.g., for this exercise, delete those rows).
    4.  ✂️ **Text-to-Columns:** Select the `City State Zip` column. Use `Data` > `Text to Columns` (Delimited, likely by space) to split it into separate `City`, `State`, and `Zip` columns. Adjust headers accordingly.
    5.  🔍 **Standardize:** Use `Find & Replace` (`Ctrl+H`) in the new `State` column to correct any inconsistencies (e.g., replace "Calif." with "CA").
    6.  🚫 **Remove Duplicates:** Use `Data` > `Remove Duplicates`. Check duplicates based on the `CustomerID` column to ensure each customer appears only once.
    7.  ⇅ **Sort:** Sort the cleaned data alphabetically by `State`, and then by `City` within each state.
    8.  💾 **Save:** Save your workbook.

---
