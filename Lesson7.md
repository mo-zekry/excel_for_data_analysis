# 🛠️ Lesson 7: Excel Data Analysis Tools

## 🎯 Learning Objectives:

* Utilize Excel's 'What-If Analysis' tools (`Goal Seek`, `Scenario Manager`) for decision-making and exploring outcomes.
* Enable and use basic features of the `Analysis ToolPak` add-in for statistical analysis (Descriptive Statistics, Correlation, Histograms).
* Apply these tools to practical, real-world scenarios.

---

## 📚 Topics Covered:

### 1. What-If Analysis Tools: Exploring Possibilities

* **Location:** `Data` tab > `Forecast` group > `What-If Analysis`
* **Goal Seek:** Finds the required input value needed for a formula to reach a specific target output value.

  * **How it works:** You specify:
    * `Set cell`: The cell containing the formula you want to achieve a target value for.
    * `To value`: The target value you want the formula to reach.
    * `By changing cell`: The single input cell that Excel can change to reach the target.
  * > **Use Case:** Finding the number of units you need to sell to break even (reach a profit of $0), determining the required score on a final exam to achieve a specific overall grade.
    >

  ![Goal Seek](./Images/Lesson7/goal_seek.png)
  *Using Goal Seek to find the input value needed to achieve a specific result*
* **Scenario Manager:** Creates and saves different sets of input values (scenarios) and allows you to compare their effects on formula results.

  * **How it works:**
    * Define `Changing cells`: The input cells that will vary between scenarios.
    * `Add` Scenarios: Give each scenario a name and enter the specific values for the changing cells for that scenario (e.g., "Best Case", "Worst Case", "Most Likely" sales projections).
    * `Show`: Select a scenario to see its impact on your worksheet formulas.
    * `Summary...`: Create a summary report (as a new sheet) comparing the results of all defined scenarios side-by-side.
  * > **Use Case:** Comparing profitability under different sales volume and cost assumptions, evaluating different loan options.
    >

  ![Scenario Manager](./Images/Lesson7/scenario_manager.png)
  *Creating and comparing different scenarios with the Scenario Manager*

### 2. Introduction to the Analysis ToolPak: Statistical Power

* **Enabling the Add-in:**

  * `File` > `Options` > `Add-ins`.
  * At the bottom, manage `Excel Add-ins`, click `Go...`.
  * Check the box for `Analysis ToolPak` and click `OK`.
  * The `Data Analysis` command will now appear on the `Data` tab.

  ![Enabling Analysis ToolPak](./Images/Lesson7/enable_toolpak.png)
  *Steps to enable the Analysis ToolPak add-in in Excel*
* **Overview:** Provides a collection of tools for statistical and engineering analysis.
* **Key Tools (Focus for this lesson):**

  * 📊 **Descriptive Statistics:** Generates a report summarizing key statistical measures for a dataset (mean, median, mode, standard deviation, variance, range, min, max, count, sum, confidence level, etc.).

    * Select `Data Analysis` > `Descriptive Statistics`.
    * Specify the `Input Range` (your data column/row).
    * Choose `Output Range` (where to put the report).
    * Check `Summary statistics`.

    ![Descriptive Statistics](./Images/Lesson7/descriptive_statistics.png)
    *Using the Descriptive Statistics tool to generate statistical summaries*
  * ⚡️ **Correlation:** Calculates the correlation coefficient between two or more numerical variables, indicating the strength and direction of their linear relationship (-1 to +1).

    * Select `Data Analysis` > `Correlation`.
    * Specify the `Input Range` (including all variables to compare).
    * Choose `Output Range`.

    ![Correlation Analysis](./Images/Lesson7/correlation.png)
    *Using the Correlation tool to analyze relationships between variables*
  * 📊 **Histogram:** Creates frequency distributions and histogram charts (similar to the standard chart type but with more options via the ToolPak, like specifying bins).

    * Select `Data Analysis` > `Histogram`.
    * Specify `Input Range` (the data to analyze).
    * Specify `Bin Range` (optional - a range containing the upper bounds for your desired bins/intervals).
    * Choose `Output Range` and check `Chart Output`.

    ![Histogram Tool](./Images/Lesson7/histogram_tool.png)
    *Creating frequency distributions and histograms with the Histogram tool*

### 3. Real-World Scenario Demonstrations

* Briefly demonstrate applying Goal Seek to find a break-even point in a simple sales/cost model.
* Show creating Best Case/Worst Case scenarios with Scenario Manager for a budget projection.
* Illustrate generating Descriptive Statistics for a column of sales figures and interpreting the output.
* Show calculating the Correlation between advertising spend and sales.

---

## ✨ Key Takeaways:

> * **Goal Seek** works backward from a desired result to find the necessary input.
> * **Scenario Manager** helps compare the outcomes of different sets of input assumptions.
> * The **Analysis ToolPak** add-in provides easy access to common **statistical analysis** tools like Descriptive Statistics, Correlation, and Histograms.

---

## 🛠️ Activity: Applying Analysis Tools

* **Goal:** Use Goal Seek and the Analysis ToolPak on a simple dataset.
* **Setup:** Create a small dataset. Examples:
  * **Sheet1 (Profit Calc):**
    * `A1`: Units Sold, `B1`: 100
    * `A2`: Price Per Unit, `B2`: $15
    * `A3`: Fixed Costs, `B3`: $500
    * `A4`: Variable Cost Per Unit, `B4`: $8
    * `A5`: Total Revenue, `B5`: `=B1*B2`
    * `A6`: Total Variable Costs, `B6`: `=B1*B4`
    * `A7`: Total Costs, `B7`: `=B3+B6`
    * `A8`: Profit, `B8`: `=B5-B7`
  * **Sheet2 (Sample Data):**
    * Column A (`X`): 10, 12, 15, 11, 18, 20, 14, 16
    * Column B (`Y`): 25, 30, 35, 28, 40, 45, 32, 38
* **Steps:**
  1. **🎯 Goal Seek - Break-Even Analysis:**
     * Go to `Sheet1`.
     * Open `Data` > `What-If Analysis` > `Goal Seek`.
     * `Set cell`: `B8` (the Profit cell).
     * `To value`: `0` (for break-even).
     * `By changing cell`: `B1` (Units Sold).
     * Click `OK`. Observe the number of units required in cell `B1` to achieve zero profit.
     * *(Optional)* Click `Cancel` to revert, or `OK` to keep the result.
  2. **📊 Analysis ToolPak - Descriptive Statistics:**
     * Ensure the Analysis ToolPak is enabled (`File` > `Options` > `Add-ins`...).
     * Go to `Sheet2`.
     * Go to `Data` > `Data Analysis`.
     * Select `Descriptive Statistics` and click `OK`.
     * `Input Range`: Select the data in Column A (`$A$1:$A$9` or similar, adjust if you have headers).
     * Check `Labels in first row` if your selection includes a header.
     * `Output Range`: Choose an empty cell on the sheet (e.g., `D1`).
     * Check `Summary statistics`.
     * Click `OK`. Review the generated statistics (Mean, Median, Standard Deviation, etc.).
  3. **⚡️ Analysis ToolPak - Correlation:**
     * Go to `Data` > `Data Analysis`.
     * Select `Correlation` and click `OK`.
     * `Input Range`: Select the data in *both* Column A and Column B (`$A$1:$B$9` or similar).
     * Check `Labels in first row` if applicable.
     * `Output Range`: Choose another empty cell (e.g., `G1`).
     * Click `OK`. Review the correlation matrix (it will show the correlation coefficient between X and Y).
  4. **(Optional) Scenario Manager:** On `Sheet1`, use Scenario Manager to create two scenarios: "Low Sales" (Units Sold = 50) and "High Sales" (Units Sold = 200). Generate a Scenario Summary report.
  5. 💾 **Save:** Save your workbook.

---
