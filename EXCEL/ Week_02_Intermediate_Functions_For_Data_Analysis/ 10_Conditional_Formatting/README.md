This is **Section 10: Conditional Formatting**, so I’ll organize it in the same Markdown style as Sections 8 and 9. I’ve also cleaned up the wording while keeping what you learned.


# Section 10: Conditional Formatting

## Overview

Conditional Formatting is an Excel feature that automatically changes the appearance of cells based on specific conditions or rules.

It helps make important values, patterns, trends, and differences in a dataset easier to identify visually.

---

## 1. Highlight Cells Rules

Highlight Cells Rules allow me to format cells that meet specific conditions.

Examples include:

- Greater Than
- Less Than
- Between
- Equal To
- Text That Contains
- A Date Occurring
- Duplicate Values

### Steps

1. Select the data range.
2. Go to **Home**.
3. Click **Conditional Formatting**.
4. Select **Highlight Cells Rules**.
5. Choose the condition.
6. Enter the required value.
7. Choose the formatting.
8. Click **OK**.

---

## 2. Top/Bottom Rules

Top/Bottom Rules help identify the highest or lowest values within a selected range.

### Top 10 Items

**Top 10 Items** highlights a specific number of cells containing the highest values in the selected range.

The number does not have to remain 10. It can be changed depending on what I want to identify.

For example, I could highlight the:

- Top 5 values
- Top 10 values
- Bottom 5 values

### Top 10%

**Top 10%** highlights cells whose values fall within the highest 10% of the selected range.

Other available rules include:

- Bottom 10 Items
- Bottom 10%
- Above Average
- Below Average

---

## 3. Data Bars

Data Bars add horizontal bars inside cells.

The length of each bar is based on the value in the cell.

- Larger values have longer bars.
- Smaller values have shorter bars.

This makes it easier to visually compare values without creating a separate chart.

### Steps

1. Select the cells.
2. Go to **Conditional Formatting**.
3. Select **Data Bars**.
4. Choose the preferred style.

---

## 4. Color Scales

Color Scales apply a color gradient to selected cells based on their values.

This creates a heat-map effect that makes it easier to identify:

- High values
- Medium values
- Low values
- Patterns and differences

### Steps

1. Select the data range.
2. Go to **Conditional Formatting**.
3. Select **Color Scales**.
4. Choose the preferred color scale.

---

## 5. Icon Sets

Icon Sets add visual icons to cells based on their values.

Examples of icons include:

- Arrows
- Symbols
- Flags
- Ratings
- Traffic lights

They provide a quick visual way to compare or categorize values.

### Steps

1. Select the data range.
2. Go to **Conditional Formatting**.
3. Select **Icon Sets**.
4. Choose the preferred icon set.

---

## 6. Formula-Based Rules

Formula-based Conditional Formatting allows me to format cells or rows based on a formula.

### Steps

1. Select the range I want to format.
2. Go to **Conditional Formatting**.
3. Select **New Rule**.
4. Choose **Use a formula to determine which cells to format**.
5. Enter the formula.
6. Choose the formatting.
7. Click **OK**.

### Example

```excel
=$K2="Active"
````

This checks whether the value in column `K` for each row is `"Active"`.

The `$` locks **column K**, while the row number remains relative so Excel can check each row.

This is especially useful when I want to format an entire row based on the value in one particular column.

---

## 7. Managing & Editing Rules

Existing Conditional Formatting rules can be viewed, edited, or deleted using **Manage Rules**.

### Steps

1. Select the relevant cells or range.
2. Go to **Conditional Formatting**.
3. Click **Manage Rules**.
4. Select the rule I want to modify.
5. Edit or delete the rule as needed.

The Rules Manager also allows me to review which range each rule applies to.

---

## 8. Clearing Rules

Conditional Formatting can be removed without deleting the actual data.

### Steps

1. Select the cells containing the Conditional Formatting.
2. Go to **Conditional Formatting**.
3. Select **Clear Rules**.
4. Choose either:

   * **Clear Rules from Selected Cells**
   * **Clear Rules from Entire Sheet**

This removes the Conditional Formatting rules while keeping the original cell values.

---

## Key Differences

| Feature               | Purpose                                                            |
| --------------------- | ------------------------------------------------------------------ |
| Highlight Cells Rules | Highlights values that meet specific conditions                    |
| Top/Bottom Rules      | Identifies highest, lowest, above-average, or below-average values |
| Data Bars             | Compares values using horizontal bars                              |
| Color Scales          | Uses color gradients to show differences in values                 |
| Icon Sets             | Uses icons to visually categorize values                           |
| Formula-Based Rules   | Applies formatting based on a custom formula                       |
| Manage Rules          | Allows existing rules to be viewed and edited                      |
| Clear Rules           | Removes Conditional Formatting                                     |

---

## Key Takeaways

* Conditional Formatting makes patterns and important values easier to identify.
* Highlight Cells Rules format cells based on specific conditions.
* Top/Bottom Rules identify high and low values.
* Data Bars provide a visual comparison of numerical values.
* Color Scales create a heat-map effect.
* Icon Sets visually categorize values.
* Formula-based rules provide more control over how formatting is applied.
* Conditional Formatting rules can be edited through **Manage Rules**.
* Rules can be removed from selected cells or the entire worksheet.

---

## Practice & Evidence

* [x] Practiced Highlight Cells Rules
* [x] Practiced Top/Bottom Rules
* [x] Created Data Bars
* [x] Applied Color Scales
* [x] Used Icon Sets
* [x] Created a formula-based rule
* [x] Managed and edited Conditional Formatting rules
* [x] Cleared Conditional Formatting rules

## Reflection

I learned how Conditional Formatting can make a dataset easier to understand by visually highlighting important values, patterns, and differences. I practiced using Highlight Cells Rules, Top/Bottom Rules, Data Bars, Color Scales, and Icon Sets.

I also learned how to create formula-based rules for more specific conditions and how to manage, edit, and clear existing rules. This helped me understand how Conditional Formatting can make data analysis more visual and help important information stand out quickly.

## Status

**Completed**




