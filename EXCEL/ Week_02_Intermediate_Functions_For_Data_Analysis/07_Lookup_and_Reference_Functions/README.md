# Section 7: Lookup & Reference Functions

## Overview

Lookup and reference functions allow me to find information in a dataset and return related information from another column or row.

These functions are especially useful when working with **multiple tables that share a common column**, such as `Department ID`.

---

## 1. VLOOKUP – Vertical Lookup

**VLOOKUP** is one of the most commonly used Excel lookup functions.

It searches for a value in the **leftmost column** of a table and returns related information from another column in the same row.

### VLOOKUP Syntax

```excel
=VLOOKUP(lookup_value, table_array, col_index_num, FALSE)
```

### Arguments

* **lookup_value** – The value I am searching for.
* **table_array** – The table containing the information I want to search.
* **col_index_num** – The column number containing the result I want.
* **FALSE** – Tells Excel to look for an exact match.

### Important Rule

For an exact match, I should use:

```excel
FALSE
```

---

## 2. Example Using the Practice Dataset

The employee table contains:

```text
Department ID
DT101
DT102
DT103
```

The department reference table also contains:

```text
Department ID | Department | Bonus % | Manager | Budget
DT101         | HR         | 8%      | Lisa Park | $50,000
DT102         | IT         | 10%     | Mike Ross | $120,000
DT103         | Sales      | 12%     | Sarah Chen| $95,000
```

The **Department ID** is the common column between the two tables.

For example, to find the department associated with `DT103`:

```excel
=VLOOKUP(B2,$J$2:$N$5,2,FALSE)
```

Excel searches for the Department ID in the first column of `J2:N5` and returns the department from column **2**.

---

## 3. Locking the Table Array

When using AutoFill to copy a VLOOKUP formula down, I need to **lock the table range**.

For example:

```excel
=VLOOKUP(B2,$J$2:$N$5,2,FALSE)
```

The `$` signs prevent the lookup table from moving when I drag the formula down.

### Shortcut

Select the range in the formula and press:

```text
F4
```

This creates an absolute reference.

---

## 4. VLOOKUP Limitation

VLOOKUP has an important limitation:

> The lookup value must be in the **leftmost column** of the lookup table.

For example, if I want to search for a value in column `L` but return something from column `K`, VLOOKUP cannot easily do this because `L` is to the right of `K`.

Alternatives include:

* `INDEX` + `MATCH`
* `XLOOKUP`

---

# 5. HLOOKUP – Horizontal Lookup

**HLOOKUP** searches for a value across the **top row** of a table and returns information from a specified row.

### Syntax

```excel
=HLOOKUP(lookup_value, table_array, row_index_num, FALSE)
```

Unlike VLOOKUP, which searches vertically, HLOOKUP searches horizontally.

---

# 6. INDEX Function

The **INDEX** function returns a value from a specific position within a range or array.

### Syntax

```excel
=INDEX(array, row_num, column_num)
```

For example:

```excel
=INDEX(A2:E11,3,2)
```

This returns the value located at:

* Row 3
* Column 2

within the selected range.

---

# 7. MATCH Function

The **MATCH** function searches for a value and returns its **position** within a range.

### Syntax

```excel
=MATCH(lookup_value, lookup_array, 0)
```

The `0` means I want an **exact match**.

### Example

```excel
=MATCH("Chicago",D2:D11,0)
```

Excel searches for `"Chicago"` and returns the position where it is found.

### Key Difference

**INDEX** tells me:

> "What value is at this position?"

**MATCH** tells me:

> "What position is this value in?"

---

# 8. INDEX + MATCH

I can combine `INDEX` and `MATCH` to create a flexible lookup.

`MATCH` first finds the position of the lookup value.

Then `INDEX` uses that position to return the corresponding value.

### Example

```excel
=INDEX($L$2:$L$5,MATCH("DT103",$J$2:$J$5,0))
```

Here:

* `MATCH` finds where `DT103` appears in the Department ID column.
* `INDEX` returns the corresponding department from column L.

This combination can overcome some of the limitations of VLOOKUP.

---

# 9. XLOOKUP

**XLOOKUP** is a modern Excel lookup function.

It is generally easier and more flexible than VLOOKUP.

### Syntax

```excel
=XLOOKUP(lookup_value, lookup_array, return_array)
```

### Example

```excel
=XLOOKUP("DT103",$J$2:$J$5,$K$2:$K$5)
```

This searches for `DT103` in the Department ID column and returns the corresponding department.

### Advantages of XLOOKUP

Unlike VLOOKUP, XLOOKUP:

* Can look to the left or right.
* Does not require a column number.
* Uses separate lookup and return ranges.
* Is easier to understand.
* Can handle missing results more conveniently.

---

# 10. CHOOSE Function

The **CHOOSE** function returns a value from a list based on a position number.

### Syntax

```excel
=CHOOSE(index_num,value1,value2,...)
```

### Example

```excel
=CHOOSE(D2,"Good","Fair","Excellent")
```

The value in `D2` determines which item Excel returns.

For example:

```text
D2 = 1 → Good
D2 = 2 → Fair
D2 = 3 → Excellent
```

---

# 11. INDIRECT Function

The **INDIRECT** function converts text into a live cell reference.

### Example

```excel
=INDIRECT("A2")
```

Excel interprets `"A2"` as a cell reference and returns the value in cell A2.

### Practical Use

INDIRECT can be useful when creating dynamic spreadsheets, especially when working with:

* Dropdown lists
* Dynamic references
* Dynamic worksheets

---

# 12. OFFSET Function

The **OFFSET** function returns a reference that is shifted by a specified number of rows and columns.

### Syntax

```excel
=OFFSET(reference,rows,cols)
```

### Example

```excel
=OFFSET(A2,4,1)
```

Starting from `A2`, Excel moves:

* **4 rows down**
* **1 column to the right**

and returns the resulting reference.

---

# 13. Key Differences

| Function        | Main Purpose                        |
| --------------- | ----------------------------------- |
| `VLOOKUP`       | Vertical lookup                     |
| `HLOOKUP`       | Horizontal lookup                   |
| `INDEX`         | Returns a value from a position     |
| `MATCH`         | Finds the position of a value       |
| `INDEX + MATCH` | Flexible lookup combination         |
| `XLOOKUP`       | Modern and flexible lookup          |
| `CHOOSE`        | Selects a value from a list         |
| `INDIRECT`      | Converts text into a cell reference |
| `OFFSET`        | Returns a shifted reference         |

---

# 14. Key Takeaways

* Lookup functions help me retrieve related information from datasets.
* VLOOKUP searches vertically.
* HLOOKUP searches horizontally.
* VLOOKUP requires the lookup column to be the leftmost column.
* `FALSE` is used for an exact VLOOKUP match.
* `$` locks a range when copying formulas.
* INDEX returns a value from a position.
* MATCH returns the position of a value.
* INDEX and MATCH can be combined for flexible lookups.
* XLOOKUP is a modern alternative to VLOOKUP.
* CHOOSE selects a value based on a position number.
* INDIRECT converts text into a cell reference.
* OFFSET creates a reference shifted by rows and columns.

---

## Practice & Evidence

* [ ] Practice VLOOKUP with the employee and department tables.
* [ ] Use VLOOKUP to return the Department.
* [ ] Use VLOOKUP to return the Bonus %.
* [ ] Use VLOOKUP to return the Manager.
* [ ] Lock the table array and use AutoFill.
* [ ] Practice HLOOKUP.
* [ ] Practice INDEX.
* [ ] Practice MATCH.
* [ ] Combine INDEX + MATCH.
* [ ] Practice XLOOKUP.
* [ ] Practice CHOOSE.
* [ ] Practice INDIRECT.
* [ ] Practice OFFSET.

## Reflection

To be updated after completing the practical exercises for Section 7.

