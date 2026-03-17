# =============================================================================
# Batch runner for Transformer_General_Automated_v3.R
#
# What this script does
# - Finds all CSV files in input_dir
# - Runs Transformer_General_Automated_v3.R once per input file (fresh R session via Rscript)
# - Writes one output Excel file per input file into output_dir
#
# What you should edit
# - input_dir: folder containing ONLY the CSVs you want to process
# - output_dir: folder where processed .xlsx files should be saved
# - setwd(): folder where Transformer_General_Automated_v3.R is located
#
# Notes on console output
# - Prints a batch summary including number of files detected
# - If >1 file, lists all filenames so you can confirm the batch contents
# - Prints per-file progress in the form [i/n] input -> output and a Status line
#
# Exit codes (from system2 / Rscript)
# - 0: success
# - 1: script error
# - 5: often indicates a path or file access issue (e.g., bad path / permissions)
# =============================================================================

# -----------------------------------------------------------------------------
# 1) Required packages (driver script only)
# -----------------------------------------------------------------------------
# In the driver script, we typically want to FAIL FAST rather than install
# packages automatically (installing can spam the console and can fail on locked
# corporate machines). This keeps batch runs deterministic and audit-friendly.
required_packages <- c("rlang")

missing <- required_packages[!vapply(required_packages, requireNamespace, logical(1), quietly = TRUE)]
if (length(missing) > 0) {
  stop(
    "Missing required package(s): ",
    paste(missing, collapse = ", "),
    "\nPlease install them and re-run."
  )
}

# Load only what we actually use in the driver script
suppressPackageStartupMessages({
  library(rlang)
})

# Operators used in the driver script
`%||%` <- rlang::`%||%`

# -----------------------------------------------------------------------------
# 2) Path helper
# -----------------------------------------------------------------------------
# File Explorer gives Windows paths with backslashes. R accepts forward slashes.
# This helper normalises Windows paths so they work reliably.
convert_path <- function(path) {
  gsub("\\\\", "/", path)
}

# -----------------------------------------------------------------------------
# 3) FIELDS TO EDIT
# -----------------------------------------------------------------------------
# Copy/paste Windows paths from File Explorer into the quotes "..."
# convert_path() will turn backslashes into forward slashes.

input_dir  <- convert_path("C:\\Users\\Llilians.Calvo\\OneDrive - Forest Research\\Documents\\2026.01.12_Lukes_code_review\\Results RAP\\R Scripts\\Files to process")
output_dir <- convert_path("C:\\Users\\Llilians.Calvo\\OneDrive - Forest Research\\Documents\\2026.01.12_Lukes_code_review\\Results RAP\\R Scripts\\Processed")

# IMPORTANT:
# This must be the folder containing the script to run, in this case
# Transformer_General_Automated_v3.R
setwd(convert_path("C:\\Users\\Llilians.Calvo\\OneDrive - Forest Research\\Documents\\2026.01.12_Lukes_code_review\\Results RAP\\R Scripts"))

# -----------------------------------------------------------------------------
# 4) Discover input files
# -----------------------------------------------------------------------------
# We intentionally filter to CSV files only to avoid accidentally processing
# other files (Excel, temporary files, etc.).
files <- list.files(
  input_dir,
  full.names = TRUE,
  pattern = "\\.csv$",
  ignore.case = TRUE
)

# If nothing is found, stop immediately with a clear message.
if (length(files) == 0) {
  stop(
    "No CSV files found in input_dir:\n  ",
    input_dir,
    "\nCheck the folder path and try again."
  )
}

# -----------------------------------------------------------------------------
# 5) Batch summary (printed once)
# -----------------------------------------------------------------------------
# This prints the number of files detected. If multiple, it lists them explicitly
# so the user can confirm what will be processed before anything runs.
cat("\n========================================\n")
cat("Batch processing started\n")
cat("Input folder : ", input_dir, "\n", sep = "")
cat("Output folder: ", output_dir, "\n", sep = "")
cat("Files found  : ", length(files), "\n", sep = "")

if (length(files) > 1) {
  cat("Files to process:\n")
  cat(paste0("  - ", basename(files)), sep = "\n")
}

cat("========================================\n\n")

# -----------------------------------------------------------------------------
# 6) Ensure output directory exists
# -----------------------------------------------------------------------------
# If the output folder is missing, create it (including any intermediate folders).
if (!dir.exists(output_dir)) {
  dir.create(output_dir, recursive = TRUE)
}

# -----------------------------------------------------------------------------
# 7) Process each file
# -----------------------------------------------------------------------------
# We run the transformer using Rscript so each file runs in a clean R session.
# That avoids cross-file contamination (objects lingering in memory) and mimics
# how the transformer would run in a production/batch environment.
for (i in seq_along(files)) {
  
  # Current file path
  file <- files[i]
  
  # Create an output name based on the input filename (same basename + _processed.xlsx)
  base_name <- tools::file_path_sans_ext(basename(file))
  output_name <- paste0(base_name, "_processed.xlsx")
  output_path <- file.path(output_dir, output_name)
  
  # Progress message: makes it obvious we are processing multiple files
  cat(sprintf("[%-2d/%-2d] %s  ->  %s\n", i, length(files), basename(file), output_name))
  
  # Call the transformer script:
  # - arg1: script name
  # - arg2: input csv path
  # - arg3: output xlsx path
  #
  # stdout/stderr are captured so we can print them only on error (clean console).
  res <- system2(
    "Rscript",
    args = c(
      shQuote("Transformer_General_Automated_v4.R"),
      shQuote(file),
      shQuote(output_path)
    ),
    stdout = TRUE,
    stderr = TRUE
  )
  
  # Exit code handling:
  # - If Rscript exits successfully, system2 often returns NULL status -> treat as 0.
  # - Otherwise, use the returned status.
  exit_code <- attr(res, "status") %||% 0
  
  # Report success vs error.
  # We keep normal runs quiet (just "OK"). If an error happens, print the captured output.
  if (exit_code == 0) {
    cat("  Status: OK\n\n")
  } else {
    cat("  Status: ERROR (exit code ", exit_code, ")\n", sep = "")
    if (length(res)) {
      cat("  Output:\n", paste(res, collapse = "\n"), "\n\n", sep = "")
    }
  }
}

# -----------------------------------------------------------------------------
# 8) End-of-batch summary 
# -----------------------------------------------------------------------------
cat("Done. Processed ", length(files), " file(s).\n", sep = "")