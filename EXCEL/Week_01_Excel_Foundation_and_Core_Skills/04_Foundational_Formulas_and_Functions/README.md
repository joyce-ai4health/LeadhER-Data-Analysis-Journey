# Section 4: Foundational Formulas & Functions

## LeadHER Foundation – Data Analysis Training

**Phase:** Phase 1 – Excel

**Week:** Week 1 – Excel Foundations & Core Skills

**Section:** Section 4

**Lesson:** Foundational Formulas & Functions

**Status:** 🟢 Completed

---

## 1. The Anatomy of an Excel Formula

An Excel formula is an instruction that starts with an **equal sign (`=`)** and calculates a new value using mathematical operations, cell references, or built-in functions.

### Important Point

All Excel formulas start with:

```text
=
```

The `=` tells Excel that I want it to **calculate something** rather than display the formula as ordinary text.

### Formulas Can Contain

* Constants
* Cell references
* Operators
* Functions

### Formula Bar vs Cell

* The **Formula Bar** shows the formula used.
* The **cell** normally shows the calculated result.

---

## 2. Arithmetic Operations in Excel

Excel can perform basic mathematical operations such as:

* Addition
* Subtraction
* Multiplication
* Division

Examples:

```text
=A1+B1
=A1-B1
=A1*B1
=A1/B1
```

---

## 3. Cell References

A cell reference tells Excel which cell contains the value I want to use.

There are three main types of cell references:

* Relative reference
* Absolute reference
* Mixed reference

---

## 4. Relative References

A **relative reference** changes when a formula is copied or dragged to another cell.

For example:

```text
=B2*C2
```

If I drag the formula down one row, Excel can automatically change it to:

```text
=B3*C3
```

This is useful when I want the formula to adjust according to its position.

---

## 5. Absolute References

An **absolute reference** does not change when the formula is copied or dragged.

The dollar sign `$` is used to lock the reference.

Example:

```text
=$B$2
```

This locks:

* Column `B`
* Row `2`

Therefore, the reference remains **B2** even when the formula is copied elsewhere.

### Shortcut

I can press:

```text
Fn + F4
```

on some keyboards to cycle through reference types.

---

## 6. Mixed References

A **mixed reference** locks either the column or the row while allowing the other part to change.

### Lock the Column

```text
=$B2
```

The column `B` is locked, but the row can change.

### Lock the Row

```text
=B$2
```

The row `2` is locked, but the column can change.

### Summary

| Reference | Example | What is Locked? |
| --------- | ------- | --------------- |
| Relative  | `B2`    | Nothing         |
| Absolute  | `$B$2`  | Column and row  |
| Mixed     | `$B2`   | Column          |
| Mixed     | `B$2`   | Row             |

---

## 7. How Excel Functions Work

A function is a **built-in instruction** that performs a specific calculation.

For example:

```text
=SUM(A2:A5)
```

### Function Structure

A function can contain different types of arguments.

Example:

```text
=SUM(A2:A5,B2,100)
```

In this example:

* `SUM` = Function name
* `A2:A5` = Range
* `B2` = Cell reference
* `100` = Constant value
* The values inside the parentheses are the function's **arguments**

---

## 8. Common Excel Functions

### SUM

Adds numbers together.

```text
=SUM(A1:A100)
```

**Purpose:** Calculates the total.

---

### AVERAGE

Calculates the mean.

```text
=AVERAGE(A1:A100)
```

**Purpose:** Finds the average value.

---

### MIN

Finds the smallest value.

```text
=MIN(A1:A100)
```

**Purpose:** Returns the minimum value.

---

### MAX

Finds the largest value.

```text
=MAX(A1:A100)
```

**Purpose:** Returns the maximum value.

---

### COUNT

Counts cells containing numbers.

```text
=COUNT(A1:A100)
```

**Purpose:** Counts numeric values only.

---

### COUNTA

Counts all non-empty cells.

```text
=COUNTA(A1:A100)
```

**Purpose:** Counts cells containing any type of data.

---

## 9. Rounding Functions

### ROUND

Rounds a number to a specified number of decimal places.

```text
=ROUND(3.567,2)
```

Result:

```text
3.57
```

---

### ROUNDUP

Rounds a number away from zero.

```text
=ROUNDUP(3.501,0)
```

Result:

```text
4
```

---

### ROUNDDOWN

Rounds a number toward zero.

```text
=ROUNDDOWN(3.999,0)
```

Result:

```text
3
```

---

### Negative Decimal Places

Excel can also round to the left of the decimal point.

Example:

```text
=ROUND(1234,-2)
```

This rounds to the nearest hundred.

---

## 10. Date Functions

### TODAY

Returns the current date.

```text
=TODAY()
```

The `TODAY` function does not require an argument.

---

### NOW

Returns the current date **and time**.

```text
=NOW()
```

### Difference

| Function   | Returns               |
| ---------- | --------------------- |
| `=TODAY()` | Current date          |
| `=NOW()`   | Current date and time |

---

## 11. Basic Text Functions

Excel also provides functions for working with text.

### UPPER

Converts text to uppercase.

```text
=UPPER(B2)
```

Example:

```text
joyce
```

becomes:

```text
JOYCE
```

---

### LOWER

Converts text to lowercase.

```text
=LOWER(B2)
```

Example:

```text
JOYCE
```

becomes:

```text
joyce
```

---

### PROPER

Changes text so that the first letter of each word is uppercase.

```text
=PROPER(B2)
```

Example:

```text
joyce adele
```

becomes:

```text
Joyce Adele
```

---

### LEN

Counts the number of characters in a cell.

```text
=LEN(B2)
```

---

### TRIM

Removes unnecessary spaces from text.

```text
=TRIM(B2)
```

This is particularly useful when cleaning data.

---

## 12. Concatenation

**Concatenation** means joining text together.

The `&` operator can be used to combine text from different cells.

Example:

```text
=E2&" "&F2
```

This joins the contents of `E2` and `F2` with a space between them.

For example:

```text
E2 = Joyce
F2 = Adele
```

Result:

```text
Joyce Adele
```

---

## 13. Understanding Excel Error Messages

Excel uses error messages to tell me that something is wrong with a formula or the data being used.

### `#DIV/0!`

Occurs when a number is divided by zero or an empty cell.

Example:

```text
=A2/B2
```

If `B2` contains zero, Excel may return:

```text
#DIV/0!
```

---

### `#VALUE!`

Usually occurs when Excel receives an inappropriate data type for the calculation.

---

### `#REF!`

Occurs when a formula refers to a cell or range that has been deleted or is no longer valid.

---

### `#N/A`

Usually means that a lookup or formula could not find the value it was looking for.

---

## 14. Key Takeaways

* Excel formulas begin with `=`.
* Formulas can contain constants, references, operators, and functions.
* The Formula Bar displays the formula while the cell displays the result.
* Relative references change when copied.
* Absolute references remain fixed.
* Mixed references lock either the row or column.
* The `$` symbol is used to lock references.
* Functions are built-in Excel instructions.
* `SUM` calculates totals.
* `AVERAGE` calculates the mean.
* `MIN` finds the smallest value.
* `MAX` finds the largest value.
* `COUNT` counts numeric cells.
* `COUNTA` counts non-empty cells.
* `TODAY()` returns the current date.
* `NOW()` returns the current date and time.
* Text functions help clean and transform data.
* Concatenation joins text together.
* Excel error messages help identify problems in formulas or data.

---

## 15. Practical Application

* [ ] Create basic arithmetic formulas.
* [ ] Use relative cell references.
* [ ] Use absolute cell references.
* [ ] Use mixed cell references.
* [ ] Use `SUM()`.
* [ ] Use `AVERAGE()`.
* [ ] Use `MIN()`.
* [ ] Use `MAX()`.
* [ ] Use `COUNT()`.
* [ ] Use `COUNTA()`.
* [ ] Use `ROUND()`.
* [ ] Use `ROUNDUP()`.
* [ ] Use `ROUNDDOWN()`.
* [ ] Use `TODAY()`.
* [ ] Use `NOW()`.
* [ ] Use `UPPER()`.
* [ ] Use `LOWER()`.
* [ ] Use `PROPER()`.
* [ ] Use `LEN()`.
* [ ] Use `TRIM()`.
* [ ] Combine text using `&`.
* [ ] Identify and understand common Excel errors.

---

## 16. Reflection

### What I Learned

I learned how Excel formulas and functions work and how to use cell references in calculations.

I also learned the difference between **relative, absolute, and mixed references**, as well as basic mathematical, date, and text functions.

### What I Need to Practice

I need to practice writing formulas, using different types of cell references, applying common functions, and understanding Excel error messages.

