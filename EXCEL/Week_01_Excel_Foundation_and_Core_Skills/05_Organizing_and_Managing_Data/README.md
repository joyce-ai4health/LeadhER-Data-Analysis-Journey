# Section 5: Organizing & Managing Data

## LeadHER Foundation – Data Analysis Training

**Phase:** Phase 1 – Excel

**Week:** Week 1 – Excel Foundations & Core Skills

**Section:** Section 5

**Lesson:** Organizing & Managing Data

**Status:** 🟢 Completed

---

## 1. What I Will Learn

This section focuses on how to organize, clean, and manage data in Excel.

The main skills covered are:

* Sorting data
* Filtering data
* Cleaning data
* Validating data
* Removing duplicates
* Splitting data using Text to Columns
* Finding and replacing values
* Creating named ranges
* Merging and unmerging cells
* Using Data Validation

---

## 2. Structured Data

**Structured data** is the foundation of every data analysis.

A well-structured dataset makes it easier to:

* Sort data
* Filter data
* Calculate values
* Create PivotTables
* Create charts
* Analyze information

### Rules for Structured Data

A good dataset should follow these rules:

* **Row 1 = Headers**
* **Every row below = One record**
* Avoid blank rows within the dataset
* Avoid merged cells within the dataset
* Avoid placing totals inside the main dataset

### Example

| Staff Name | Department | Salary | Region |
| ---------- | ---------- | -----: | ------ |
| Joyce      | Finance    | 250000 | Lagos  |
| Ada        | Sales      | 300000 | Abuja  |
| John       | Marketing  | 280000 | Kano   |

Each row represents **one record**, while each column represents a **field/variable**.

---

# 3. Sorting Data

**Sorting** changes the order of data based on a selected column.

For example, I can sort sales data:

* From smallest to largest
* From largest to smallest
* Alphabetically A–Z
* Alphabetically Z–A
* By date

### How to Sort Data

1. Select a cell inside the dataset.
2. Go to **Home**.
3. Select **Sort & Filter**.
4. Select **Custom Sort**.
5. Select the column I want to sort.
6. Select the desired order.
7. Click **OK**.

### Example

If I have a **Total Sales** column, I can sort it from:

```text
Largest → Smallest
```

This allows me to quickly identify the highest sales.

---

# 4. Filtering Data

**Filtering** allows me to display only the records that meet specific conditions.

For example, I could filter a sales dataset to show:

* Sales from Lagos
* Sales by a particular salesperson
* Products in a specific category
* Sales above ₦100,000

### How to Apply a Filter

1. Click inside the dataset.
2. Go to **Home**.
3. Select **Sort & Filter**.
4. Select **Filter**.

### Keyboard Shortcut

```text
Ctrl + Shift + L
```

Filter dropdown arrows will appear on the column headers.

I can then select the values I want to display.

---

# 5. Removing Duplicates

Duplicate records can affect the accuracy of an analysis.

Excel provides a tool for identifying and removing duplicate rows.

### How to Remove Duplicates

1. Highlight the entire dataset.
2. Go to the **Data** tab.
3. Find the **Data Tools** section.
4. Select **Remove Duplicates**.
5. Select the columns Excel should check.
6. Click **OK**.

### Important

I should understand the dataset before removing duplicates.

Not every repeated value is a duplicate record.

For example, two customers may legitimately have the same name.

---

# 6. Text to Columns

**Text to Columns** allows me to split information from one column into multiple columns.

For example:

```text
Full Name
Joyce Adele
John Smith
Mary Johnson
```

I can split the names into:

| First Name | Last Name |
| ---------- | --------- |
| Joyce      | Adele     |
| John       | Smith     |
| Mary       | Johnson   |

### Using Text to Columns

1. Select the column containing the data.
2. Go to the **Data** tab.
3. Select **Text to Columns**.
4. Choose the appropriate option.
5. Select the delimiter.
6. Specify the data format.
7. Complete the process.

### Delimiter

A **delimiter** is a character that separates pieces of information.

Common delimiters include:

* Space
* Comma
* Tab
* Semicolon

For example:

```text
Joyce,Adele
```

The comma is the delimiter.

### Fixed Width

Text to Columns can also split data using **Fixed Width**, where the data is divided based on specific character positions rather than a delimiter.

---

# 7. Find & Replace

**Find & Replace** allows me to search for a particular value and replace it with another value.

### How to Use Find & Replace

1. Select the relevant column or dataset.
2. Go to **Home**.
3. Select **Find & Select**.
4. Select **Replace**.
5. Enter the value I want to find.
6. Enter the replacement value.
7. Choose **Replace** or **Replace All**.

### Example

If a dataset contains:

```text
Lagos
Lagos
Lagos
```

and I need to change them to:

```text
Lagos State
```

Find & Replace can do this quickly.

---

# 8. Named Ranges

A **named range** allows me to give a range of cells a meaningful name.

Instead of referring to:

```text
A2:A100
```

I could give the range a name such as:

```text
Sales
```

This can make formulas easier to understand and use.

### How to Create a Named Range

1. Select the range.
2. Go to the **Formulas** tab.
3. Go to **Defined Names**.
4. Select **Name Manager**.
5. Select **New**.
6. Enter the name for the range.
7. Click **OK**.

---

# 9. Merging & Unmerging Cells

**Merge & Center** combines multiple selected cells into one larger cell.

### How to Merge Cells

1. Select the cells I want to merge.
2. Go to the **Home** tab.
3. Find the **Alignment** section.
4. Select **Merge & Center**.

### How to Unmerge Cells

Select the merged cell and choose:

**Home → Alignment → Merge & Center → Unmerge Cells**

### Important Data Analysis Note

Merged cells should generally be **avoided inside structured datasets** because they can interfere with sorting, filtering, and other data-analysis operations.

They are more appropriate for titles and presentation areas.

---

# 10. Data Validation

**Data Validation** controls what type of information users can enter into a cell.

It can help prevent incorrect data entry.

For example, I can create a dropdown list containing:

```text
Lagos
Abuja
Kano
Port Harcourt
```

Instead of allowing users to type anything they want.

### How to Access Data Validation

1. Select the relevant cells or column.
2. Go to the **Data** tab.
3. Go to **Data Tools**.
4. Select **Data Validation**.
5. Choose the validation criteria.
6. Set the allowed values.
7. Click **OK**.

### Why Data Validation Is Useful

Data Validation helps:

* Reduce data-entry errors
* Keep information consistent
* Standardize categories
* Improve data quality

---

# 11. Key Takeaways

* Structured data is important for accurate analysis.
* Each row should represent one record.
* Headers should clearly describe each column.
* Blank rows and unnecessary merged cells should be avoided within datasets.
* Sorting changes the order of records.
* Filtering displays only records that meet selected conditions.
* `Ctrl + Shift + L` activates filters.
* Remove Duplicates helps eliminate duplicate records.
* Text to Columns can split information into separate columns.
* Find & Replace can quickly correct or change repeated values.
* Named ranges give meaningful names to cell ranges.
* Merging cells is useful for presentation but should generally be avoided inside datasets.
* Data Validation helps control what users can enter.

---

# 12. Practical Application

* [ ] Create a structured sales dataset.
* [ ] Sort sales from largest to smallest.
* [ ] Sort sales from smallest to largest.
* [ ] Sort a text column alphabetically.
* [ ] Filter the dataset by region.
* [ ] Filter the dataset by product.
* [ ] Filter the dataset by salesperson.
* [ ] Remove duplicate records.
* [ ] Use Text to Columns.
* [ ] Use Find & Replace.
* [ ] Create a named range.
* [ ] Merge and unmerge cells.
* [ ] Create a dropdown list using Data Validation.

---

# 13. Reflection

## What I Learned

I learned how to organize and manage datasets in Excel using sorting, filtering, data cleaning, and data validation tools.

I also learned how to remove duplicates, split data using Text to Columns, find and replace values, create named ranges, and control data entry using Data Validation.

## What I Need to Practice

I need to practice these tools using a real-world sales dataset so that I can become comfortable cleaning, organizing, and preparing data for analysis.

