# RAP Data Transformer

**RAP — Reproducible Analytical Pipelines | Forest Research**

The pipeline takes raw CSV files as input and produces formatted Excel workbooks containing summary tables and proportion charts for each metric.

---

## Scripts

| Script | Purpose |
|---|---|
| `Automator_v3.R` | Finds all CSV files in a folder and runs the transformer once per file, in a clean R session. |
| `Transformer_General_Automated_v4.R` | Processes a single CSV file: cleans the data, calculates statistics, and writes a formatted Excel output with tables and charts. |

> **You only ever need to edit and run `Automator_v3.R`.** The transformer runs automatically in the background for each file.

---

## Files You Need

Before running, make sure you have the following in the same folder:

- `Automator_v3.R`
- `Transformer_General_Automated_v4.R`
- A folder containing the CSV files you want to process
- A folder where the processed Excel outputs will be saved (created automatically if it does not exist)

---

## Setup — Edit the Paths

Open `Automator_v3.R` in RStudio. There are three lines near the top of the script (Section 3) that you must edit to match your own folder locations.

### input_dir
The folder containing your CSV files. Only CSV files will be picked up — other file types are ignored.
```r
input_dir <- convert_path("C:\\Your\\Folder\\Files to process")
```

### output_dir
The folder where the processed Excel files will be saved. Created automatically if it does not exist.
```r
output_dir <- convert_path("C:\\Your\\Folder\\Processed")
```

### setwd()
The folder where both R scripts are saved. This must be the folder containing `Transformer_General_Automated_v3.R`.
```r
setwd(convert_path("C:\\Your\\Folder\\R Scripts"))
```

> **Tip:** You can copy a path directly from Windows File Explorer. The `convert_path()` function automatically converts backslashes to forward slashes.

---

## Running the Pipeline

Once the three paths are set, run the entire `Automator_v3.R` script in RStudio (`Ctrl+Shift+Enter`, or click **Source** in the top-right of the editor pane).

The console will print a summary like the following:

```
========================================
Batch processing started
Input folder : C:/Your/Folder/Files to process
Output folder: C:/Your/Folder/Processed
Files found  : 3
Files to process:
  - habitat_deadwoodvol_v_band_B_ALL.csv
  - habitat_disease_v_band_B_ALL.csv
  - habitat_invasive_v_band_B_ALL.csv
========================================

[1 /3 ] habitat_deadwoodvol_v_band_B_ALL.csv  ->  habitat_deadwoodvol_v_band_B_ALL_processed.xlsx
  Status: OK

[2 /3 ] habitat_disease_v_band_B_ALL.csv  ->  habitat_disease_v_band_B_ALL_processed.xlsx
  Status: OK
```

Each file runs in its own clean R session, so errors in one file do not affect the others.

---

## Outputs

For each input CSV, one Excel file is written to the output folder containing two sheets:

| Sheet | Contents |
|---|---|
| `Tables` | Formatted summary tables showing area (ha) and standard error (SE%) for each habitat type and metric category, broken down by GB, Scotland, England, Wales, and all sub-regions. |
| `Graphs` | Proportions bar charts for GB, Scotland, England, and Wales, showing the percentage of area in each category by habitat type. |

> Cells highlighted in **orange** indicate that the standard error exceeds 25%, meaning the estimate for that cell should be interpreted with caution.

---

## Supported Data Types

The transformer automatically detects the type of data in each CSV by inspecting its contents. No manual configuration is needed.

| Data Type | Detected By |
|---|---|
| Vertical structure (storeys) | group2 contains "Storey" or "Storeys" |
| Number of native tree species | group2 contains "native species" |
| Proportion of favourable habitat | group2 contains "good quality land cover" |
| Proportion of woodland | group2 contains "land cover as woodland" |
| Veteran trees | group2 contains "trees per ha" |
| Herbivore / grazing pressure | group2 contains "damage" |
| Regeneration — square level | group2 contains Seedlings/Saplings + filename contains "regensection" |
| Regeneration — stand level | group2 contains Seedlings/Saplings + filename contains "regenpop" |
| All other metrics (standard) | Everything else: deadwood, disease, invasive, open space, veg layers, tree age, canopy nativeness, overall score |

> **Note:** The two regeneration file types (`regenpop` and `regensection`) have identical data content and can only be distinguished by their filename. All other types are detected purely from the data.

---

## Troubleshooting

### Status: ERROR (exit code 1)
The transformer encountered an R error. The full error message is printed below the status line. Common causes:
- A path contains a typo or the folder does not exist
- A required R package is not installed (see [Required Packages](#required-r-packages))
- The CSV file has an unexpected structure or column names

### Status: ERROR (exit code 5)
Usually a file access or path issue. Check that the paths in the Setup section are correct and that you have read/write permission to those folders.

### No CSV files found
The `input_dir` path is incorrect, or the folder contains no `.csv` files. Check the path and confirm the files are present.

### Category label shows NULL in the output
The data type was not recognised by the detection logic. Check that the `group2` column values in the CSV match the patterns listed in [Supported Data Types](#supported-data-types).

---

## Required R Packages

The following packages must be installed before running. If any are missing, install them from the R console:

```r
install.packages(c("tidyr", "dplyr", "tidyverse", "readxl", "writexl",
                   "hash", "openxlsx", "openxlsx2", "rlang"))
```

| Package | Used For |
|---|---|
| `tidyr` | Pivoting data between long and wide formats |
| `dplyr` | Data manipulation and summarisation |
| `tidyverse` | Collection including ggplot2, stringr, purrr, forcats |
| `readxl` | Reading Excel files |
| `writexl` | Writing Excel files |
| `hash` | Hash map lookups for labels and habitat names |
| `openxlsx` | Excel workbook creation and formatting |
| `openxlsx2` | Supplementary Excel functions |
| `rlang` | Used in the automator for the `%||%` operator |

---

## How the Transformer Works

### 1. Data detection
After reading the CSV, the transformer inspects the unique values in the `group2` column to determine which RAP metric the file contains. This sets a `data_type` variable used throughout the rest of the script to select the correct labels and processing logic.

### 2. Pre-processing
- `group1` values are mapped to readable habitat or nativeness category names via a lookup table (`habitat_map`)
- `group2` values have the `"Group2: "` prefix stripped to create the `Variable` column
- Rows where `group2` contains only a dot (unclassified records) are removed
- Rows with habitat values of `"None"` or `"Missing value"` are removed
- Separate data frames are created for Scotland and England; Wales is handled within the full dataset

### 3. Pivoting and aggregation
The data is pivoted from long to wide format so that each category becomes its own column. Where multiple rows share the same region, habitat, and category combination, values are summed (`values_fn = sum`). Results are then grouped and summed by habitat type within each region.

### 4. Standard error calculations
SE% is calculated for each category column as:

```
SE% = 100 × √z3 / bulk_up_forecast
```

A `Total` column is added (sum of all category areas per row), and a corresponding `SE% Total` column is derived from the individual SE values.

### 5. Regional tables
Separate summary tables are produced for GB, Scotland, England, and all individual sub-regions. These are stacked into a single data frame with separator rows between each region and formatted as a single Excel sheet.

### 6. Proportions and charts
A proportions version of each country-level table is calculated by expressing each category's area as a percentage of the row total. These are used to produce horizontal bar charts (one per habitat type) for GB, Scotland, England, and Wales, saved as PNG files and inserted into the Graphs sheet of the Excel output.

### 7. Excel formatting
The Tables sheet is formatted with:
- Dark green header rows with white text
- Alternating grey fill for numeric cells
- Bold grey fill for Total rows and columns
- Orange font applied via conditional formatting to any SE% value exceeding 25%, and to the corresponding area value in the adjacent column
- Column widths set to 50 for the habitat column and 10 for all numeric columns