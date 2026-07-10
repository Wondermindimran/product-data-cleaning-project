# Product Data Cleaning & SEO Title Optimization

## Project Overview
This project focuses on preparing raw e-commerce product data for marketing analysis and search engine optimization. The dataset contained inconsistent formats, missing values, duplicate records, and overly long product titles, all of which were cleaned and standardized using **Microsoft Excel**.

The key deliverable is a cleaned dataset with a new **`short_title`** column; a concise (30–50 character) version of each product title, designed to improve SEO and readability.


## Key Improvements

| Issue | Before | After |
|-------|--------|-------|
| Product length units | Mixed (inches, cm, mm) | All converted to **inches** |
| Missing bullet points | ~20% empty | Filled with *"No features listed"* |
| Duplicate product IDs | 159 duplicates | **Removed** |
| Average title length | 142 characters | **38 characters** (`short_title`) |
| Special characters in titles | `|`, `[`, `]`, `{`, `}` | Replaced or removed |

---

## Files in This Repository

| File | Description |
|------|-------------|
| `productdata_original.xlsx` | The raw with dataset with `long_title` and `duplicated_columns` before cleaning (in the **raw** folder) |
| `productdata_cleaned.xlsx` | Final cleaned dataset with `short_title` column added (in the **cleaned** folder) |
| `Data Cleaning & Title Optimization Report.pdf` | Full technical documentation (14 pages) |
| `README.md` | This project overview |


## Tools Used
- **Microsoft Excel** (no external add-ons or coding required)


## How the Cleaning Was Done

### 1. Missing Values
Blank cells in the `BULLET_POINTS` column were replaced with `"No features listed"` using Excel's **Find & Replace**.

### 2. Duplicate Removal
Duplicate rows were identified and removed based on the `PRODUCTID` column using Excel's **Remove Duplicates** feature.

### 3. Unit Conversion (Most Critical Step)
The `ProductLength` column had a mix of inches, centimeters, and millimeters. All values were converted to a single unit (**inches**) using helper columns and `IF` formulas.

**Formula used to detect unit:**
```excel
=IF(AND(F2>=1, F2<=500), "inches", IF(AND(F2>=501, F2<=2000), "cm", IF(F2>2000, "mm", "unknown")))
