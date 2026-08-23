# Section 6: Intermediate Functions & Data Analysis

## LeadHER Foundation – Data Analysis Training

**Phase:** Phase 1 – Excel

**Week:** Week 2 – Intermediate Functions for Data Analysis

**Section:** Section 6

**Lesson:** Intermediate Functions & Data Analysis

**Status:** 🟢 Completed

---

## 1. What We Will Learn

This section focuses on using Excel functions to make decisions, summarize data, handle errors, rank values, and work with dates.

The main topics covered are:

* Conditional logic with `IF`, `AND`, `OR`, and `NOT`
* Nested `IF`
* `IFS`
* Conditional aggregation functions
* `COUNTIF` and `COUNTIFS`
* `SUMIF` and `SUMIFS`
* `AVERAGEIF` and `AVERAGEIFS`
* Error handling with `IFERROR`
* `LARGE` and `SMALL`
* Ranking with `RANK.EQ`
* Date functions
* `DATEDIF`
* `NETWORKDAYS`
* `EDATE`

---

# 2. IF Function

The `IF` function allows me to build **decision-making into a formula**.

It checks whether a condition is `TRUE` or `FALSE`.

* If the condition is **TRUE**, Excel returns one result.
* If the condition is **FALSE**, Excel returns another result.

### Syntax

```text
=IF(condition, value_if_true, value_if_false)
```

### Example

```text
=IF(H2>85000,"High Pay","Regular Pay")
```

If `H2` is greater than `85,000`, Excel returns:

```text
High Pay
```

Otherwise, it returns:

```text
Regular Pay
```

---

# 3. Nested IF

A **Nested IF** is an `IF` function placed inside another `IF` function.

It is useful when I have **multiple conditions and multiple possible results**.

### Example

Suppose:

* `4.5` and above = Excellent
* `3.5` and above = Good
* Below `3.5` = Fair

The formula can be:

```text
=IF(H2>=4.5,"Excellent",IF(H2>=3.5,"Good","Fair"))
```

### How It Works

Excel checks the conditions from left to right:

1. Is `H2 >= 4.5`?
2. If not, is `H2 >= 3.5`?
3. If neither condition is true, return `Fair`.

---

# 4. IFS Function

The `IFS` function provides a more modern approach to handling multiple conditions.

It takes **conditions and their corresponding results as pairs**.

### Example

```text
=IFS(H2>=4.5,"Excellent",H2>=3.5,"Good",H2<3.5,"Fair")
```

### How It Works

Excel checks each condition in order.

| Condition | Result    |
| --------- | --------- |
| `H2>=4.5` | Excellent |
| `H2>=3.5` | Good      |
| `H2<3.5`  | Fair      |

---

# 5. AND Function

The `AND` function is used when **multiple conditions must be TRUE**.

It returns `TRUE` only when all specified conditions are true.

### Example

```text
=IF(AND(M2>=4.5,N2="Yes"),"High","Regular")
```

In this example, both conditions must be true:

* `M2` must be greater than or equal to `4.5`
* `N2` must equal `"Yes"`

If both are true, the result is:

```text
High
```

Otherwise:

```text
Regular
```

---

# 6. OR Function

The `OR` function is used when **at least one condition needs to be TRUE**.

### Example

```text
=IF(OR(M2>=4.5,N2="Yes"),"High","Regular")
```

If either condition is true, Excel returns:

```text
High
```

If neither condition is true, Excel returns:

```text
Regular
```

---

# 7. NOT Function

The `NOT` function **reverses or negates a condition**.

If a condition is `TRUE`, `NOT` makes it `FALSE`.

If a condition is `FALSE`, `NOT` makes it `TRUE`.

### Example

```text
=IF(NOT(Q2="Terminated"),"Send Renewal","Do Not Send")
```

This means:

* If the employee is **not terminated** → `Send Renewal`
* If the employee **is terminated** → `Do Not Send`

---

# 8. Conditional Aggregation Functions

Conditional aggregation functions allow me to calculate values based on specific conditions.

These include:

* `COUNTIF`
* `COUNTIFS`
* `SUMIF`
* `SUMIFS`
* `AVERAGEIF`
* `AVERAGEIFS`

---

# 9. COUNTIF

`COUNTIF` counts cells that meet **one condition**.

### Syntax

```text
=COUNTIF(range,criteria)
```

### Example

```text
=COUNTIF(F:F,"New York")
```

This counts the number of cells in column `F` containing:

```text
New York
```

---

# 10. COUNTIFS

`COUNTIFS` counts cells that meet **multiple conditions**.

### Example

```text
=COUNTIFS(E:E,"Female",F:F,"New York")
```

This counts records where:

* Gender = Female
* Location = New York

Both conditions must be satisfied.

---

# 11. SUMIF

`SUMIF` adds values based on **one condition**.

### Syntax

```text
=SUMIF(range,criteria,sum_range)
```

### Example

```text
=SUMIF(C:C,"IT",H:H)
```

This adds the values in column `H` for records where column `C` contains:

```text
IT
```

---

# 12. SUMIFS

`SUMIFS` adds values based on **multiple conditions**.

### Example

```text
=SUMIFS(H:H,E:E,"Female",F:F,"New York")
```

This calculates the total in column `H` where:

* Column `E` = Female
* Column `F` = New York

---

# 13. AVERAGEIF

`AVERAGEIF` calculates the average of values that meet **one condition**.

### Syntax

```text
=AVERAGEIF(range,criteria,average_range)
```

### Example

```text
=AVERAGEIF(C:C,"IT",H:H)
```

This calculates the average of the values in column `H` for records where column `C` contains `IT`.

---

# 14. AVERAGEIFS

`AVERAGEIFS` calculates the average based on **multiple conditions**.

### Example

```text
=AVERAGEIFS(H:H,E:E,"Male",C:C,"IT")
```

This calculates the average of column `H` where:

* Column `E` = Male
* Column `C` = IT

---

# 15. IFERROR

`IFERROR` allows me to handle errors and display a more meaningful result instead.

### Syntax

```text
=IFERROR(value,value_if_error)
```

### Example

```text
=IFERROR(C4/0,"Illegal math operation")
```

Instead of displaying:

```text
#DIV/0!
```

Excel displays:

```text
Illegal math operation
```

### Why It Is Useful

`IFERROR` can make reports cleaner and easier for users to understand.

---

# 16. LARGE Function

The `LARGE` function returns the **largest value based on a specified position**.

### Example

```text
=LARGE(H:H,5)
```

This returns the **5th largest value** in column `H`.

### Example

If the values are:

```text
100
90
80
70
60
```

Then:

```text
=LARGE(A1:A5,2)
```

returns:

```text
90
```

---

# 17. SMALL Function

The `SMALL` function returns the **smallest value based on a specified position**.

### Example

```text
=SMALL(H:H,1)
```

This returns the **smallest value** in column `H`.

For example:

```text
=SMALL(A1:A5,2)
```

returns the second-smallest value.

---

# 18. RANK.EQ Function

`RANK.EQ` returns the rank of a specific number within a range.

### Syntax

```text
=RANK.EQ(number,ref,order)
```

### Example

```text
=RANK.EQ(H10,H2:H12,0)
```

The `0` means ranking from **largest to smallest**.

Therefore:

```text
1 = Highest value
2 = Second highest
3 = Third highest
```

If I use `1` instead, Excel ranks from **smallest to largest**.

---

# 19. Date Functions

Excel provides functions that allow me to extract information from dates.

### YEAR

Extracts the year from a date.

```text
=YEAR(H2)
```

### MONTH

Extracts the month number from a date.

```text
=MONTH(H2)
```

### DAY

Extracts the day from a date.

```text
=DAY(H2)
```

### Example

If:

```text
H2 = 15/08/2026
```

Then:

```text
=YEAR(H2)
```

returns:

```text
2026
```

```text
=MONTH(H2)
```

returns:

```text
8
```

```text
=DAY(H2)
```

returns:

```text
15
```

---

# 20. DATEDIF Function

`DATEDIF` calculates the difference between two dates.

### Syntax

```text
=DATEDIF(start_date,end_date,"unit")
```

### Example

```text
=DATEDIF(A2,B2,"y")
```

The `"y"` calculates the difference in **complete years**.

Other units can be used to calculate months or days.

---

# 21. NETWORKDAYS Function

`NETWORKDAYS` calculates the number of **working days between two dates**.

By default, it excludes weekends.

### Syntax

```text
=NETWORKDAYS(start_date,end_date)
```

### Example

```text
=NETWORKDAYS(X1,X2)
```

This calculates the working days between the two dates.

### Excluding Public Holidays

I can provide a range containing holidays:

```text
=NETWORKDAYS(X1,X2,A3:A5)
```

Excel will exclude:

* Weekends
* The dates listed in `A3:A5`

This is useful for calculating working days, project durations, and employee work periods.

---

# 22. EDATE Function

`EDATE` shifts a date forward or backward by a specified number of months.

### Syntax

```text
=EDATE(start_date,months)
```

### Move a Date Forward

```text
=EDATE(X1,5)
```

Moves the date in `X1` **5 months forward**.

### Move a Date Backward

```text
=EDATE(X1,-3)
```

Moves the date in `X1` **3 months backward**.

---

# 23. Key Takeaways

* `IF` allows me to make decisions using formulas.
* Nested `IF` handles multiple conditions.
* `IFS` provides another way to handle multiple conditions.
* `AND` requires all conditions to be true.
* `OR` requires at least one condition to be true.
* `NOT` reverses a condition.
* `COUNTIF` counts based on one condition.
* `COUNTIFS` counts based on multiple conditions.
* `SUMIF` adds values based on one condition.
* `SUMIFS` adds values based on multiple conditions.
* `AVERAGEIF` calculates an average based on one condition.
* `AVERAGEIFS` calculates an average based on multiple conditions.
* `IFERROR` helps handle formula errors.
* `LARGE` finds values based on their position from the top.
* `SMALL` finds values based on their position from the bottom.
* `RANK.EQ` ranks values within a range.
* `YEAR`, `MONTH`, and `DAY` extract information from dates.
* `DATEDIF` calculates the difference between dates.
* `NETWORKDAYS` calculates working days between dates.
* `EDATE` moves a date forward or backward by months.

---

# 24. Practical Application

* [ ] Practice `IF`
* [ ] Practice Nested `IF`
* [ ] Practice `IFS`
* [ ] Practice `AND`
* [ ] Practice `OR`
* [ ] Practice `NOT`
* [ ] Practice `COUNTIF`
* [ ] Practice `COUNTIFS`
* [ ] Practice `SUMIF`
* [ ] Practice `SUMIFS`
* [ ] Practice `AVERAGEIF`
* [ ] Practice `AVERAGEIFS`
* [ ] Practice `IFERROR`
* [ ] Practice `LARGE`
* [ ] Practice `SMALL`
* [ ] Practice `RANK.EQ`
* [ ] Practice `YEAR`
* [ ] Practice `MONTH`
* [ ] Practice `DAY`
* [ ] Practice `DATEDIF`
* [ ] Practice `NETWORKDAYS`
* [ ] Practice `EDATE`

---

# 25. Reflection

## What I Learned

I learned how to use intermediate Excel functions to make decisions, summarize data based on conditions, handle errors, rank values, and work with dates.

## What I Need to Practice

I need to practice combining these functions with real-world datasets so that I can use them confidently for data analysis and reporting.

