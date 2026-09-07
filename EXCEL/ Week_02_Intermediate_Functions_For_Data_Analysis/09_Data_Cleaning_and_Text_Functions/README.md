
# Section 9: Data Cleaning & Text Functions

## Overview

Data cleaning is the process of fixing or removing inconsistent, corrupted, incorrectly formatted, duplicate, or incomplete data within a dataset.

Real-world data is almost never completely clean, so data cleaning is an important part of the data analysis process.

---

## 1. Why Data Cleaning Matters

Data cleaning is where analysis actually begins.

Common data quality problems include:

- Extra spaces
- Inconsistent text casing
- Merged fields
- Hidden or non-printable characters
- Duplicate records
- Missing values
- Incorrectly formatted data
- Inconsistent values

The quality of an analysis depends heavily on the quality of the data being analyzed.

> **Key Lesson:** Clean data leads to more reliable analysis.

---

## 2. TRIM Function

The `TRIM` function removes unnecessary spaces from text.

### Syntax

```excel
=TRIM(text)
````

### Example

```excel
=TRIM(A2)
```

### Cleaning a Column with TRIM

1. Create a new column beside the messy column.
2. Enter the `TRIM` formula.
3. Fill the formula down the column.
4. Highlight and copy the cleaned column.
5. Paste it back as **Values** using the paste option with `123`.
6. The cleaned values are now detached from the original formula.

This is useful when I want to remove the original messy column without affecting my cleaned data.

---

## 3. CLEAN Function

The `CLEAN` function removes non-printable or hidden characters from text.

### Syntax

```excel
=CLEAN(text)
```

### Example

```excel
=CLEAN(A2)
```

`CLEAN` can also be nested inside `TRIM`:

```excel
=TRIM(CLEAN(A2))
```

This allows me to remove hidden characters and unnecessary spaces at the same time.

---

## 4. LEFT Function

The `LEFT` function extracts a specified number of characters starting from the **left side** of a text value.

### Syntax

```excel
=LEFT(text,num_chars)
```

### Example

```excel
=LEFT(B2,4)
```

This returns the first **4 characters** from the value in `B2`.

---

## 5. RIGHT Function

The `RIGHT` function extracts a specified number of characters starting from the **right side** of a text value.

### Syntax

```excel
=RIGHT(text,num_chars)
```

### Example

```excel
=RIGHT(B2,4)
```

This returns the last **4 characters** from the value in `B2`.

---

## 6. MID Function

The `MID` function extracts characters from the middle of a text value.

### Syntax

```excel
=MID(text,start_num,num_chars)
```

### Example

```excel
=MID(B2,10,4)
```

This tells Excel to:

* Start at character 10.
* Extract 4 characters.

---

## 7. FIND & SEARCH Functions

`FIND` and `SEARCH` are used to identify the position of specific text or characters inside another text value.

### FIND

```excel
=FIND(find_text,within_text)
```

Example:

```excel
=FIND("e",J2)
```

`FIND` is **case-sensitive**.

This means:

```text
e ≠ E
```

### SEARCH

```excel
=SEARCH(find_text,within_text)
```

Example:

```excel
=SEARCH("e",J2)
```

`SEARCH` is **not case-sensitive**.

This means it does not distinguish between:

```text
e and E
```

### Difference

| Function | Case-Sensitive? |
| -------- | --------------- |
| `FIND`   | Yes             |
| `SEARCH` | No              |

---

## 8. SUBSTITUTE Function

The `SUBSTITUTE` function replaces specific existing text with new text.

### Syntax

```excel
=SUBSTITUTE(text,old_text,new_text)
```

### Example

```excel
=SUBSTITUTE(B2," ","-")
```

If `B2` contains:

```text
Data Analysis
```

the result becomes:

```text
Data-Analysis
```

### When to Use SUBSTITUTE

Use `SUBSTITUTE` when I **know the specific text or character I want to replace**.

For example:

* Replacing spaces with dashes
* Changing abbreviations
* Removing specific characters
* Standardizing text values

---

## 9. REPLACE Function

The `REPLACE` function replaces characters based on their **position** within a text value.

### Syntax

```excel
=REPLACE(old_text,start_num,num_chars,new_text)
```

### Example

```excel
=REPLACE(B2,10,4,"google")
```

This tells Excel to:

1. Look at the value in `B2`.
2. Start from character 10.
3. Replace 4 characters.
4. Insert `"google"`.

### SUBSTITUTE vs REPLACE

A useful rule to remember:

**Use `SUBSTITUTE` when I know what text I want to replace.**

**Use `REPLACE` when I know where the text is located.**

---

## 10. VALUE Function

The `VALUE` function converts a number stored as **text into a numeric value**.

### Syntax

```excel
=VALUE(text)
```

### Example

```excel
=VALUE(J9)
```

This is useful when numbers have been imported or stored as text and need to be used for calculations.

---

## 11. TEXT Function

The `TEXT` function converts a number or date into **formatted text**.

### Syntax

```excel
=TEXT(value,format_text)
```

For example, if `J10` contains a date:

```excel
=TEXT(J10,"MMMM")
```

returns the full month name, such as:

```text
September
```

Using:

```excel
=TEXT(J10,"mmm")
```

returns the abbreviated month:

```text
Sep
```

### VALUE vs TEXT

| Function | Conversion                   |
| -------- | ---------------------------- |
| `VALUE`  | Text → Number                |
| `TEXT`   | Number/Date → Formatted Text |

---

## 12. TEXTJOIN Function

`TEXTJOIN` combines text from multiple cells and allows me to specify what should separate the values.

### Syntax

```excel
=TEXTJOIN(delimiter,ignore_empty,text1,...)
```

### Example

```excel
=TEXTJOIN(" ",TRUE,J11:L11)
```

Here:

* `" "` adds a space between each value.
* `TRUE` tells Excel to ignore empty cells.
* `J11:L11` contains the values to join.

For example:

```text
Joyce | James | Etata
```

can become:

```text
Joyce James Etata
```

---

## 13. TEXTSPLIT Function

`TEXTSPLIT` performs the opposite type of operation to joining text.

It splits text into separate cells using a specified delimiter.

### Split Across Columns

```excel
=TEXTSPLIT(J12," ")
```

If `J12` contains:

```text
Joyce James Etata
```

Excel can split the words across separate columns.

### Split Across Rows

`TEXTSPLIT` can also split values vertically by supplying the delimiter as the row delimiter.

For example:

```excel
=TEXTSPLIT(A12,," ")
```

This can split space-separated text down different rows.

---

## 14. TEXTBEFORE & TEXTAFTER

`TEXTBEFORE` and `TEXTAFTER` are modern Excel functions that extract text based on a specified delimiter.

### TEXTBEFORE

Returns the text that appears **before** a delimiter.

```excel
=TEXTBEFORE(B2,"@")
```

For example:

```text
joyce@gmail.com
```

returns:

```text
joyce
```

### TEXTAFTER

Returns the text that appears **after** a delimiter.

```excel
=TEXTAFTER(B2,"@")
```

Using:

```text
joyce@gmail.com
```

returns:

```text
gmail.com
```

---

## 15. Flash Fill

**Flash Fill** automatically fills values based on a pattern that Excel detects from examples I provide.

For example, if a column contains:

```text
Joyce James
John Smith
Mary Brown
```

and I manually type:

```text
Joyce
```

in another column, Excel may recognize that I am trying to extract the first names.

### Flash Fill Shortcut

```text
Ctrl + E
```

Flash Fill is useful for:

* Separating names
* Combining information
* Reformatting text
* Extracting parts of text
* Creating values based on recognizable patterns

---

## 16. Key Differences Between the Text Functions

| Function     | Purpose                                       |
| ------------ | --------------------------------------------- |
| `TRIM`       | Removes unnecessary spaces                    |
| `CLEAN`      | Removes non-printable characters              |
| `LEFT`       | Extracts characters from the left             |
| `RIGHT`      | Extracts characters from the right            |
| `MID`        | Extracts characters from the middle           |
| `FIND`       | Finds text position and is case-sensitive     |
| `SEARCH`     | Finds text position and is not case-sensitive |
| `SUBSTITUTE` | Replaces specific text                        |
| `REPLACE`    | Replaces text based on position               |
| `VALUE`      | Converts text to a number                     |
| `TEXT`       | Converts a value to formatted text            |
| `TEXTJOIN`   | Joins multiple text values                    |
| `TEXTSPLIT`  | Splits text into separate cells               |
| `TEXTBEFORE` | Returns text before a delimiter               |
| `TEXTAFTER`  | Returns text after a delimiter                |
| `Flash Fill` | Automatically recognizes and applies patterns |

---

## 17. Key Takeaways

* Data cleaning is an important part of data analysis.
* Real-world datasets often contain inconsistencies that need to be corrected.
* `TRIM` removes unnecessary spaces.
* `CLEAN` removes non-printable characters.
* `TRIM` and `CLEAN` can be combined for better text cleaning.
* `LEFT`, `RIGHT`, and `MID` extract parts of text.
* `FIND` is case-sensitive while `SEARCH` is not.
* `SUBSTITUTE` replaces known text.
* `REPLACE` replaces characters based on their position.
* `VALUE` converts text into numbers.
* `TEXT` converts values into formatted text.
* `TEXTJOIN` combines text.
* `TEXTSPLIT` separates text.
* `TEXTBEFORE` and `TEXTAFTER` extract text around a delimiter.
* Flash Fill automatically applies patterns it recognizes.

---

## Practice & Evidence

* [x] Practiced `TRIM`
* [x] Practiced `CLEAN`
* [x] Combined `TRIM` and `CLEAN`
* [x] Practiced `LEFT`
* [x] Practiced `RIGHT`
* [x] Practiced `MID`
* [x] Practiced `FIND`
* [x] Practiced `SEARCH`
* [x] Practiced `SUBSTITUTE`
* [x] Practiced `REPLACE`
* [x] Practiced `VALUE`
* [x] Practiced `TEXT`
* [x] Practiced `TEXTJOIN`
* [x] Practiced `TEXTSPLIT`
* [x] Practiced `TEXTBEFORE`
* [x] Practiced `TEXTAFTER`
* [x] Practiced Flash Fill

## Reflection

I learned that data cleaning is an important step before analyzing a dataset because errors and inconsistencies in the data can affect the accuracy of the analysis.

Through the practical exercises, I learned how to use different Excel text functions to remove unwanted spaces and characters, extract parts of text, replace values, join and split text, and convert between text and numeric values. I also practiced using Flash Fill to perform repetitive text transformations more efficiently.



