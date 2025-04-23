# 🧮 Lesson 3: Excel Formulas and Functions Basics

**(⏳ Estimated Duration: 2 hours)**

---

## 🎯 Learning Objectives:

*   Understand the structure and syntax of Excel formulas.
*   Master the difference between relative, absolute, and mixed cell referencing.
*   Learn to use fundamental mathematical, logical, and text functions for data analysis.
*   Recognize common Excel formula errors and understand basic error handling.

---

## 📚 Topics Covered:

### 1. Understanding Excel Formulas

*   **The Basics:**
    *   All formulas start with an equals sign (`=`).
    *   Use standard arithmetic operators: `+` (Add), `-` (Subtract), `*` (Multiply), `/` (Divide), `^` (Exponent).
    *   **Order of Operations (PEMDAS/BODMAS):** Parentheses/Brackets, Exponents/Orders, Multiplication/Division (from left to right), Addition/Subtraction (from left to right).

    ![Basic Excel Formula Syntax](./Images/Lesson3/formula_basics.png)
    *Example of a basic formula in Excel showing the components*

*   **Cell Referencing:** Pointing to other cells to use their values in calculations.
    *   **Single Cell:** `=A1+B1`
    *   **Range:** `=SUM(C1:C10)` (Calculates the sum of cells C1 through C10).

    ![Cell Referencing Example](./Images/Lesson3/cell_referencing.png)
    *Example showing how cell references work in formulas*

### 2. Absolute vs. Relative Referencing: The Power of the `$`

*   **Relative Referencing (Default):** References adjust automatically when you copy/paste or fill a formula. If you copy `=A1+B1` from cell `C1` down to `C2`, it becomes `=A2+B2`.
    *   *Example:* `A1`

    ![Relative Referencing Example](./Images/Lesson3/relative_reference.png)
    *Example showing how relative references change when copied down or across*

*   **Absolute Referencing:** References **do not** change when copied/filled. Use the dollar sign (`$`) before the column letter and row number.
    *   *Example:* `$A$1` (Always refers to cell A1, no matter where the formula is copied).

    ![Absolute Referencing Example](./Images/Lesson3/absolute_reference.png)
    *Example showing how absolute references stay fixed when copied*

*   **Mixed Referencing:** Either the row or the column is absolute, while the other is relative.
    *   *Example:* `$A1` (Column A is fixed, row adjusts when copied vertically).
    *   *Example:* `A$1` (Row 1 is fixed, column adjusts when copied horizontally).
    *   > **Pro Tip:** Use the `F4` key while editing a cell reference in the formula bar to cycle through relative, absolute, and mixed options!

    ![Mixed Referencing Example](./Images/Lesson3/mixed_reference.png)
    *Example showing how mixed references partially adjust when copied*

### 3. Commonly Used Excel Functions for Analysis

*   **Functions** are pre-built formulas that perform specific calculations.
*   **Syntax:** `=FUNCTIONNAME(argument1, argument2, ...)`

    ![Function Syntax](./Images/Lesson3/function_syntax.png)
    *Example showing the structure of an Excel function*

*   **Basic Math Functions:**
    *   `SUM(range)`: Adds all numbers in a range of cells. *Ex: `=SUM(A1:A10)`*
    *   `AVERAGE(range)`: Calculates the average of numbers in a range. *Ex: `=AVERAGE(B1:B10)`*
    *   `COUNT(range)`: Counts how many cells in a range contain numbers. *Ex: `=COUNT(C1:C10)`*
    *   `COUNTA(range)`: Counts how many cells in a range are not empty (contain numbers, text, etc.). *Ex: `=COUNTA(D1:D10)`*
    *   `MIN(range)`: Finds the smallest number in a range. *Ex: `=MIN(E1:E10)`*
    *   `MAX(range)`: Finds the largest number in a range. *Ex: `=MAX(F1:F10)`*

    ![Math Functions Example](./Images/Lesson3/math_functions.png)
    *Example showing how basic math functions calculate values from data*

*   **Logical Functions:** Used for decision-making.
    *   `IF(logical_test, value_if_true, value_if_false)`: Returns one value if a condition is true, and another if it's false. *Ex: `=IF(A1>10, "Pass", "Fail")`*
    *   `AND(logical1, [logical2], ...)`: Returns `TRUE` if ALL arguments are true, otherwise `FALSE`. Often used within `IF`. *Ex: `=IF(AND(A1>10, B1="Active"), "Proceed", "Stop")`*
    *   `OR(logical1, [logical2], ...)`: Returns `TRUE` if ANY argument is true, otherwise `FALSE`. Often used within `IF`. *Ex: `=IF(OR(A1="Shipped", B1="Delivered"), "Complete", "Pending")`*

    ![Logical Functions Example](./Images/Lesson3/logical_functions.png)
    *Example showing how IF, AND, and OR functions work with conditions*

*   **Text Functions:** Used for manipulating text strings.
    *   `CONCATENATE(text1, [text2], ...)` or using the `&` operator: Joins several text strings into one. *Ex: `=CONCATENATE(A1, " ", B1)` or `=A1 & " " & B1` (to join first and last names with a space).
    *   `LEFT(text, num_chars)`: Returns a specified number of characters from the start of a text string. *Ex: `=LEFT(A1, 3)` (gets the first 3 characters).
    *   `RIGHT(text, num_chars)`: Returns a specified number of characters from the end of a text string. *Ex: `=RIGHT(A1, 2)` (gets the last 2 characters).
    *   `LEN(text)`: Returns the number of characters in a text string (length). *Ex: `=LEN(A1)`*

    ![Text Functions Example](./Images/Lesson3/text_functions.png)
    *Example showing how text functions manipulate string data*

### 4. Error Handling in Excel Formulas

*   **Common Errors & Meanings:**
    *   `#DIV/0!`: Trying to divide by zero.
    *   `#N/A`: A value is not available to a function or formula (common with lookups).
    *   `#NAME?`: Excel doesn't recognize text in a formula (e.g., misspelled function name).
    *   `#NULL!`: Specifies an invalid intersection of two areas.
    *   `#NUM!`: Invalid numeric values in a formula or function (e.g., square root of a negative number).
    *   `#REF!`: A cell reference is not valid (e.g., cell was deleted).
    *   `#VALUE!`: Wrong type of argument or operand is used (e.g., adding text to a number).

    ![Common Excel Errors](./Images/Lesson3/common_errors.png)
    *Examples of common Excel formula errors*

*   **Basic Handling with `IFERROR`:**
    *   `IFERROR(value, value_if_error)`: Checks if the first part results in an error. If it does, it returns the `value_if_error`; otherwise, it returns the result of the `value`.
    *   *Ex: `=IFERROR(A1/B1, "Cannot Divide")` (Instead of showing `#DIV/0!`, it will show "Cannot Divide" if B1 is 0).

    ![IFERROR Example](./Images/Lesson3/iferror_example.png)
    *Example showing how IFERROR can provide user-friendly error messages*

---

## ✨ Key Takeaways:

> *   Formulas start with `=` and use **cell references** and **operators**.
> *   Understanding **relative vs. absolute (`$`) referencing** is crucial for copying formulas correctly.
> *   **Functions** (`SUM`, `AVERAGE`, `IF`, `CONCATENATE`, etc.) automate common calculations and tasks.
> *   Recognizing **formula errors** helps in troubleshooting, and **`IFERROR`** can provide cleaner outputs.

---

## 🛠️ Activity: Applying Basic Formulas & Functions

*   **Goal:** Practice using arithmetic operations, cell referencing, and basic functions.
*   **Setup:** Use the `Practice Data` sheet from Lesson 1 (Item Name, Quantity, Price) or create a similar small dataset.
    *   *If creating new:* Columns: `Product`, `Units Sold`, `Unit Price`, `Discount Rate` (enter as a percentage like 5% or 0.05).
*   **Steps:**
    1.  ➕ **Calculate Total Revenue (Basic Formula):** In a new column (e.g., `Total Revenue`), create a formula to multiply `Quantity` by `Price` for the first item. *Example: `=B2*C2`*
    2.  📋 **Copy Formula (Relative Referencing):** Use the fill handle (small square at the bottom-right of the cell) to drag the formula down for all items. Observe how the cell references change.
    3.  💲 **Calculate Discount Amount (Mixed Referencing):** Assume a single discount rate is stored in a separate cell (e.g., `H1`). In a new column (`Discount Amount`), calculate the discount for the first item (`Total Revenue` * `Discount Rate`). Make the reference to the discount rate cell **absolute** (`$H$1`) before copying down. *Example: `=D2*$H$1`*
    4.  ➖ **Calculate Final Price:** In a new column (`Final Price`), subtract the `Discount Amount` from the `Total Revenue`. Copy the formula down.
    5.  ∑ **Use `SUM`:** Below the `Final Price` column, use the `SUM` function to calculate the total final price for all items.
    6.  📊 **Use `AVERAGE`, `MAX`, `MIN`, `COUNT`:** In separate cells, calculate the Average, Highest, Lowest, and Count of `Units Sold` using the respective functions.
    7.  🤔 **Use `IF`:** In a new column (`High Volume?`), use the `IF` function to display "Yes" if `Units Sold` is greater than a certain number (e.g., 50) and "No" otherwise. *Example: `=IF(B2>50, "Yes", "No")`*. Copy down.
    8.  🔗 **Use `CONCATENATE` / `&`:** In a new column (`Product Info`), combine the `Product` name and its `Units Sold` into one text string. *Example: `=A2 & " - Units: " & B2`*. Copy down.
    9.  🚫 **(Optional) Practice `IFERROR`:** Intentionally create a division by zero error in a cell (e.g., `=10/0`). In another cell, use `IFERROR` to wrap that calculation and display a custom message like "Error".
    10. 💾 **Save:** Save your workbook.

---
