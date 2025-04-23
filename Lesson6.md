# 📈 Lesson 6: Data Visualization and Dashboards in Excel

**(⏳ Estimated Duration: 2 hours)**

---

## 🎯 Learning Objectives:

*   Learn to create various standard chart types in Excel for effective data visualization.
*   Understand when to use different chart types (Column, Bar, Line, Pie, Scatter, Histogram).
*   Apply conditional formatting and data bars to highlight trends and patterns directly within cells.
*   Combine multiple elements (charts, Pivot Tables, slicers) to build interactive dashboards.

---

## 📚 Topics Covered:

### 1. Creating Effective Visualizations: Choosing the Right Chart

*   **Standard Chart Types (Insert Tab > Charts Group):**
    *   📊 **Column Charts:** Ideal for comparing values across different categories or showing changes over a short period. Variations include Clustered, Stacked, 100% Stacked.

    ![Column Charts](./Images/Lesson6/column_charts.png)
    *Examples of different column chart types and when to use them*

    *   📊 **Bar Charts:** Similar to column charts but with horizontal bars. Good for comparing many categories or when category labels are long.

    ![Bar Charts](./Images/Lesson6/bar_charts.png)
    *Bar charts showing comparisons across categories with long labels*

    *   📉 **Line Graphs:** Excellent for showing trends over time (days, months, years) or continuous data. Connects data points with lines.

    ![Line Graphs](./Images/Lesson6/line_graphs.png)
    *Line graphs showing data trends over time periods*

    *   🧁 **Pie Charts:** Used to show proportions of a whole (parts of 100%). Best for a small number of categories. Avoid using if categories are numerous or very similar in size. Variations include Doughnut charts.

    ![Pie Charts](./Images/Lesson6/pie_charts.png)
    *Pie and doughnut charts showing proportional data distributions*

    *   ⚛️ **Scatter Plots (XY Charts):** Show the relationship or correlation between two numerical variables. Each point represents an observation based on two values.

    ![Scatter Plots](./Images/Lesson6/scatter_plots.png)
    *Scatter plots showing correlations between two variables*

    *   📊 **Histograms:** Visualize the distribution of a single numerical variable by grouping data into bins (ranges) and showing the frequency (count) within each bin. (Found under Statistical charts).

    ![Histograms](./Images/Lesson6/histograms.png)
    *Histograms showing the distribution of values in different bins*

*   **Chart Elements & Formatting:**
    *   Adding/Modifying Titles, Axis Labels, Data Labels, Legends (Use the `+` icon or `Chart Design` tab).
    *   Changing Colors, Styles, and Layouts (`Chart Design` tab).
    *   Adjusting Axes Scales and Formatting (`Format Axis` pane - double-click axis).

    ![Chart Formatting](./Images/Lesson6/chart_formatting.png)
    *Chart formatting options showing how to add and modify chart elements*

### 2. Conditional Formatting & Data Bars: In-Cell Visualization

*   **Conditional Formatting (Home Tab > Styles Group):** Apply formatting (colors, icons, bars) to cells based on their values or specific rules.
    *   **Highlight Cells Rules:** Greater than, less than, between, equal to, text that contains, duplicate values, etc.

    ![Highlight Cell Rules](./Images/Lesson6/highlight_cell_rules.png)
    *Using highlight cell rules to identify values meeting specific conditions*

    *   **Top/Bottom Rules:** Top 10 items, Bottom 10%, Above/Below Average, etc.

    ![Top-Bottom Rules](./Images/Lesson6/top_bottom_rules.png)
    *Applying top/bottom rules to highlight highest and lowest values*

    *   **Data Bars:** Adds colored bars directly within cells, representing the value relative to others in the selected range. Provides a quick visual comparison.

    ![Data Bars](./Images/Lesson6/data_bars.png)
    *Data bars showing relative values within cells*

    *   **Color Scales:** Applies a color gradient across cells based on their values (e.g., green for high, red for low).

    ![Color Scales](./Images/Lesson6/color_scales.png)
    *Color scales showing value distribution with color gradients*

    *   **Icon Sets:** Adds small icons (arrows, shapes, indicators) to cells based on their values.

    ![Icon Sets](./Images/Lesson6/icon_sets.png)
    *Icon sets visually indicating values with directional or rating symbols*

    *   **Managing Rules:** Edit, delete, or change the order of applied rules.

    ![Managing Rules](./Images/Lesson6/manage_rules.png)
    *Managing multiple conditional formatting rules in a worksheet*

### 3. Introduction to Excel Dashboards: Bringing It All Together

*   **What is a Dashboard?** A visual display (often on a single sheet) of the most important information needed to achieve one or more objectives; consolidated and arranged so the information can be monitored at a glance.

    ![Dashboard Overview](./Images/Lesson6/dashboard_overview.png)
    *Example of a complete Excel dashboard showing multiple data visualizations*

*   **Key Components:**
    *   **Data Source:** Clean, well-structured data (often an Excel Table).
    *   **Analysis Layer:** Pivot Tables or formulas used to summarize the source data.
    *   **Presentation Layer:** Charts, Key Performance Indicators (KPIs), and other visuals derived from the analysis layer.
    *   **Interactivity:** Slicers and Timelines (as seen in Lesson 5) to allow users to filter the dashboard dynamically.

    ![Dashboard Components](./Images/Lesson6/dashboard_components.png)
    *The different layers of a dashboard showing how data flows from source to presentation*

*   **Design Principles (Basic):**
    *   **Audience & Purpose:** Who is it for? What questions should it answer?
    *   **Layout:** Organize charts and elements logically. Place important info prominently (top-left often).
    *   **Clarity:** Use clear titles and labels. Avoid clutter.
    *   **Consistency:** Use consistent colors and formatting.

    ![Dashboard Design](./Images/Lesson6/dashboard_design.png)
    *Examples of good and poor dashboard design practices*

*   **Combining Elements:** Arrange Pivot Charts, standard charts (potentially linked to summary data derived from Pivot Tables or formulas), Slicers, and Timelines on a single worksheet.
    *   Ensure Slicers/Timelines are connected to the relevant Pivot Tables/Charts (`Report Connections...`).

    ![Connecting Dashboard Elements](./Images/Lesson6/connecting_elements.png)
    *Connecting slicers to multiple pivot tables and charts in a dashboard*

---

## ✨ Key Takeaways:

> *   Choosing the **right chart type** is crucial for conveying information effectively.
> *   **Conditional Formatting** and **Data Bars** provide quick, in-cell visual insights.
> *   Dashboards combine **data, analysis, and visualization** into a consolidated, interactive view.
> *   **Slicers and Timelines** are key to making dashboards interactive.

---

## 🛠️ Activity: Creating an Interactive Dashboard

*   **Goal:** Build a simple interactive dashboard using charts, Pivot Tables, and slicers.
*   **Setup:** Use the `Sales Data` and potentially the `Product Info` sheets from previous lessons. Ensure you have Pivot Tables created (like Sales by Region, Sales by Category, Sales over Time).
*   **Steps:**
    1.  **Prepare Analysis:** Ensure you have at least 2-3 Pivot Tables summarizing key aspects of your data (e.g., Total Sales by Region, Total Sales by Category, Units Sold by Salesperson, Sales Trend by Month).
    2.  **Create Visuals:**
        *   Create appropriate Pivot Charts for each of your key Pivot Tables (e.g., Bar chart for Region Sales, Pie chart for Category Sales, Line chart for Sales Trend).
        *   *(Optional)* Create a standard chart (not Pivot) based on a small summary table derived from a Pivot Table or formulas (e.g., a card showing Total Overall Sales).
    3.  **Design Dashboard Sheet:** Create a new, blank worksheet named `Dashboard`.
    4.  **Move/Copy Elements:** Move or copy your created charts onto the `Dashboard` sheet. Arrange them logically.\ (Tip: Cut (`Ctrl+X`) and Paste (`Ctrl+V`) moves charts easily).
    5.  **Add Interactivity:**
        *   Select one of the Pivot Tables that a chart on your dashboard is based on.
        *   Insert `Slicers` for key dimensions you want to filter by (e.g., `Region`, `Salesperson`, `Year` if applicable).
        *   Insert a `Timeline` if you have a `Date` field.
        *   Move the Slicers/Timeline to the `Dashboard` sheet and arrange them (often at the top or side).
    6.  **Connect Slicers/Timeline:** For *each* Slicer and Timeline, right-click > `Report Connections...` and ensure it is connected to *all* the relevant Pivot Tables that feed into your dashboard charts.
    7.  **Apply Conditional Formatting (Example):** Go back to your main `Sales Data` sheet. Select the `TotalSale` column. Apply `Conditional Formatting` > `Data Bars` (choose a style). Observe how it adds visual context to the raw data.
    8.  **Test Interactivity:** Click various options on your Slicers and Timeline. Watch how all connected charts update dynamically.
    9.  **Refine Layout:** Adjust chart sizes, positions, and formatting for clarity and visual appeal. Hide gridlines (`View` tab > `Show` group > uncheck `Gridlines`) for a cleaner look.
    10. 💾 **Save:** Save your workbook.

---
