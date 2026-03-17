library(tidyr)
library(readxl)
library(hash)
library(dplyr)
library(tidyverse)
library(writexl)
`%||%` <- rlang::`%||%`
syms <- rlang::syms

# Get command-line arguments
args <- commandArgs(trailingOnly = TRUE)

# Assign arguments
file_path <- args[1]         # Input file
name_of_output <- args[2]    # Output file name or full path
file_id <- tools::file_path_sans_ext(basename(file_path)) # to get the suffix in the plot names and avoid rewritting


# Get directory from the output path
output_dir <- dirname(name_of_output)

# Ensure the output directory exists
if (!dir.exists(output_dir)) {
  dir.create(output_dir, recursive = TRUE)
}

# Set working directory to the output folder (optional)
setwd(output_dir)

# Read input file
raw <- read.csv(file_path)

# ---- Detect data type from group2 content ----

# Rather than relying on filenames, we inspect the actual group2 values in the raw data
# to determine which metric this file contains. This makes the script filename-independent.
# Each data type has a unique signature in its group2 values.

g2_vals <- unique(raw$group2)

detect_data_type <- function(g2_vals, file_path) {
  if (any(grepl("Storey|Storeys|Complex|No storeys", g2_vals, ignore.case = TRUE))) {
    return("vertical_structure")
  } else if (any(grepl("native species", g2_vals, ignore.case = TRUE))) {
    return("native_tree_species")       # NumberofNativeTreeSpecies / NoNativeRaw
  } else if (any(grepl("good quality land cover|quality land cover", g2_vals, ignore.case = TRUE))) {
    return("favourable_habitat")        # ProportionFavourableLandcover (PropAsHab)
  } else if (any(grepl("land cover as woodland", g2_vals, ignore.case = TRUE))) {
    return("proportion_woodland")       # ProportionFavourableLandcover (PropAsWood)
  } else if (any(grepl("trees per ha", g2_vals, ignore.case = TRUE))) {
    return("veteran_trees")             # habitat_veterantree (uses "X trees per ha" not range format)
  } else if (any(grepl("damage", g2_vals, ignore.case = TRUE))) {
    return("grazing")                   # All herbivore/grazing files (browsing, squirrel, combined)
  } else if (any(grepl("Seedlings|Saplings|<7 cm", g2_vals, ignore.case = TRUE))) {
    # regenpop and regensection have identical group2 values - only filename distinguishes them.
    # Check for section/square-level keywords first, then fall back to stand/population.
    if (grepl("regensection|square.level|RegenSquare", file_path, ignore.case = TRUE)) {
      return("regen_square")            # Square-level / component group regeneration
    } else {
      return("regen_stand")             # Stand-level / population regeneration (regenpop)
    }
  } else {
    return("standard")                  # All other RAP metrics
  }
}

data_type <- detect_data_type(g2_vals, file_path)

# ---- Hash maps ----

# Maps data_type to the human-readable axis/table label for that metric.
# For "standard" files the label is resolved via name_hash after processing,
# since many standard files are distinguished by their group2 first-word alone.
data_type_label <- hash(
  "vertical_structure"  = "Number of storeys present",
  "native_tree_species" = "Number of native tree and/or shrub species",
  "favourable_habitat"  = "Proportion of favourable habitat",
  "proportion_woodland" = "Proportion of woodland",
  "regen_square"        = "Regeneration at component group level",
  "regen_stand"         = "Regeneration at population level",
  "veteran_trees"       = "Number of veteran trees per hectare",
  "grazing"             = "Herbivores / grazing pressure"
)

# Used to resolve the category label for standard files, keyed on the first word
# of the group2 variable (after "Group2: " prefix is stripped). Only used when
# data_type == "standard"; all other types are resolved directly via data_type_label.
name_hash <- hash(
  "Field"        = "Proportion of field layer vegetation in the component group",
  "Ground"       = "Proportion of ground layer vegetation in the component group",
  "Bare"         = "Proportion of bare soil in component group",
  "<5"           = "Size of woodland parcel",
  ">=5"          = "Size of woodland parcel",
  "No"           = "Tree health and disease measure",
  ">0"           = "Volume of deadwood m3 per ha",
  "Invasive"     = "Presence of invasive species",
  ">=1"          = "Number of veteran trees per hectare",
  "0"            = "Number of veteran trees per hectare",
  ">1"           = "Number of veteran trees per hectare",
  ">10"          = "Number of veteran trees per hectare",
  ">2"           = "Number of veteran trees per hectare",
  ">20"          = "Number of veteran trees per hectare",
  ">5"           = "Number of veteran trees per hectare",
  "Young"        = "Age distribution of tree species",
  "Intermediate" = "Age distribution of tree species",
  "Old"          = "Age distribution of tree species",
  ">=0"          = "Canopy nativeness",
  "Woodland"     = "Proportion of open space",
  "Squirrel"     = "Herbivores / grazing pressure",
  "Browsing"     = "Herbivores / grazing pressure",
  "Cat"          = "Tree health and disease measure",
  "Total"        = "Overall ecological condition score",
  "None"         = "Volume of deadwood m3 per ha"
)


# Used to rename habitat columns 
habitat_map <- hash(
  "Group1: Lowland beech/yew woodland" = "Lowland beech/yew woodland",
  "Group1: Lowland Mixed Deciduous Woodland" = "Lowland Mixed Deciduous Woodland",
  "Group1: Native pine woodlands" = "Native pine woodlands",
  "Group1: Non HAP native pinewood" = "Non-HAP native pinewood",
  "Group1: Upland birchwoods" = "Upland birchwoods (Scot); birch dominated upland woods",
  "Group1: Upland mixed ashwoods" = "Upland mixed ashwoods",
  "Group1: Upland oakwood" = "Upland oakwood",
  "Group1: Wet woodland" = "Wet woodland",
  "Group1: Wood Pasture & Parkland" = "Wood Pasture & Parkland",
  "Group1: BROADLEAVED, MIXED/YEW WOODLANDS" = "Broadleaf habitat NOT classified as priority",
  "Group1: CONIFEROUS WOODLANDS" = "Non-native coniferous woodland",
  "Group1: Transition or felled" = "Transition or felled",
  "Group1: ." = "None",
  "Group1: missing value" = "Missing value",
  "Group1:                ." = "None",
  # Vertical structure nativeness categories
  "Group1: Native" = "Native",
  "Group1: Near native" = "Near native",
  "Group1: Non-native" = "Non-native",
  "Group1: Not determinable" = "Not determinable"
)

# Used to replace region codes with full names
region_map <- hash(
  "GB" = "GB",
  "Scotland" = "Scotland",
  "England" = "England", 
  "NWE" = "North West England",
  "NEE" = "North East England",
  "YHE" = "Yorkshire and the Humber",
  "EME" = "East Midlands",
  "EE" = "East England",
  "SELE" = "South East England", # Sometimes London is in there somewhere
  "SWE" = "South West England",
  "WME" = "West Midlands",
  "NS" = "North Scotland",
  "NES" = "North East Scotland",
  "ES" = "East Scotland",
  "SS" = "South Scotland",
  "WS" = "West Scotland",
  "W" = "Wales",
  "SEE" = "South East London"
)

# ---- Data Pre-Processing ----

df <- tryCatch(
  {
    # Rename habitat column to get rid of "Group 1:" as well as deal with "Group1: ." and missing values
    raw |>
      mutate(Habitat = map_chr(group1, ~ habitat_map[[.x]]))
  },
  error = function(e) {
    # If error occurs, do this fallback:  remove  "Group: " from habitat (or similar data)
    raw |>
      mutate(Habitat = str_replace(group1, "Group1: ", "")) |>
      filter(!str_detect(group1, "\\."))
  }
)

# Remove "Group: " from category Variable (which is deadwood volume, invasive cover etc)
df$Variable <- gsub("Group2: ", "", df$group2)
df <- df[!grepl("\\.", df$group2), ]


# Remove unneeded columns 
df$group2 <- NULL
df$group1 <- NULL
df$group3 <- NULL

# Create data frames for Scotland and England, note since Wales is a single region it is handled within the full df 
df_Scotland <- df[df$country == "Scotland", ]
df_England <- df[df$country == "England", ]

# ---- Functions ----

# The pipe functions in this section take care of the data processing of the data they are ran sequentially
# in the full pipe for df_Scotland and df_England but the full df "df_wide" requires a slightly
# different pipe so the pipes are called individually for that case 

# Pivot wider pipe, converts the data into a wide format so have one column per category 
pivot_pipe <- function(df){
  df <- df |>
    pivot_wider(
      names_from = Variable,
      values_from = c(bulk_up_forecast, z3),
      values_fill = 0,
      values_fn = sum
    )
  return(df) 
}

# Summing pipe, sums value by 'Habitat' groups
# This pipe also removes missing values and incoherent data, depending on the users preference the final two lines
# in this function could be moved forward
sum_pipe <- function(df, groups = "Habitat"){
  start <- grep("^bulk", names(df))[1]
  
  cols_to_sum <- names(df)[start:ncol(df)]
  df <- df |>
    group_by(!!!syms(groups)) |>
    summarise(
      across(all_of(cols_to_sum), ~sum(.x, na.rm = TRUE), .names = "{.col}"),
      .groups = "drop"
    )
  
  df <- df[!(df$Habitat %in% c("Missing value", "None")),]
  df <- df |> select(-any_of("z3_Group2:                        ."))
  return(df)
}

# Total row pipe, creates a row which is the region specific total value across habitats has "Total" in its Habitat column
total_row_pipe <- function(df, from = 2) {
  sum_row <- df |>
    summarise(across(from:ncol(df), \(x) sum(x, na.rm = TRUE)))
  
  sum_row <- sum_row |>
    mutate(Habitat = "Total") |>
    select(Habitat, everything())
  
  df <- bind_rows(df, sum_row)
  return(df)
}

# SE pipe, handles the standard error calculations 
se_pipe <- function(df){
  cols_to_sum <- names(df)[2:ncol(df)]
  
  bulk_cols <- cols_to_sum[startsWith(cols_to_sum, "bulk_")]
  
  for (col in bulk_cols) {
    total_colname <- col
    se_colname <- sub("^bulk_up_forecast_", "z3_", col)
    
    df <- df |>
      mutate(
        !!se_colname := if_else(
          .data[[total_colname]] == 0,
          0,
          100 * sqrt(.data[[se_colname]]) / .data[[total_colname]]
        )
      )
  }
  return(df)
}

# Total column pipe, creates a column "total "Total" which is the region specific total value across categories 
total_col_pipe <- function(df){
  bulk_cols <- grep("^bulk_", names(df), value = TRUE)
  df$Total = rowSums(df[, bulk_cols])
  
  return(df)
}


# Total SE column pipe, creates a column "total "SE% Total" which is standard error of the total column
total_se_pipe <- function(df){
  bulk_cols <- grep("^bulk_up_forecast_", names(df), value = TRUE)
  Ranges <- str_replace(bulk_cols, "^bulk_up_forecast_", "")
  
  df <- df |>
    rowwise() |>
    mutate(
      `SE% Total` = {
        # For each range, get the values from bulk_up_forecast_<range> and z3_<range>
        vals <- sapply(Ranges, function(r) {
          bulk_val <- get(paste0("bulk_up_forecast_", r))
          z3_val <- get(paste0("z3_", r))
          bulk_val * z3_val
        })
        ifelse(Total == 0, 0, sqrt(sum(vals^2, na.rm = TRUE)) / Total)
      }
    ) |>
    ungroup()
  return(df)
}

# Column reordering and renaming pipe, renames and reorders the columns to be what we want in the final excel table
column_fix_pipe <- function(df, all_regions = FALSE){
  
  bulk_cols = grep("bulk_up", names(df), value = TRUE)
  
  colnames(df) <- sub("^bulk_up_forecast_", "", colnames(df))
  colnames(df) <- sub("^z3_", "SE% ", colnames(df))
  
  if (all_regions == TRUE) {
    start <- 3
    shift = 2
  }
  else {
    start <- 2
    shift = 1
  }
  
  # Interlacing columns
  cols_order <- names(df)[start:(length(bulk_cols) + shift)]
  cols_order <- c(cols_order, "Total")
  
  cols_order <- as.vector(rbind(cols_order, paste0("SE% ", cols_order)))
  cols_order <- c("Habitat", cols_order)
  
  if (all_regions == TRUE) {
    cols_order <- c("region", cols_order)
  }
  
  df <- df[, cols_order]
  
  return(df)
}

# Full pipeline, combines the pipes defined above into a single function that acts on a dataframe
full_pipe <- function(df){
  df <- pivot_pipe(df)
  df <- sum_pipe(df)
  df <- total_row_pipe(df)
  df <- se_pipe(df)
  df <- total_col_pipe(df)
  df <- total_se_pipe(df)
  df <- column_fix_pipe(df)
  return(df)
}

# Proportions pipe, creates the proportions table by dividing the value for each category by the total value, on a per habitat basis
prop_pipe <- function(df){
  df <- df |>
    mutate(
      across(
        names(df)[!(startsWith(names(df), "Habitat") | startsWith(names(df), "region"))], ~ ifelse(Total != 0, 100 * .x / Total, 0)
      )
    ) 
  
  df <- df[, !grepl("SE%", names(df), fixed = TRUE)]
  return(df)
}

# ---- Excel operations functions ----

# These functions are responsible for applying the excel formatting to the data such as colours, conditional formatting, and headers

# A helper function to streamline creating specific workbooks
make_wb <- function(df, name = "Tables"){
  wb <- createWorkbook()
  addWorksheet(wb, "Tables")
  
  # Write the data
  writeData(wb, name, df, startRow = 3, colNames = FALSE)
  return(wb)
} 

# Adding formatting such as colors, fonts, conditional formatting, and headers to the workbook
beautify_Tables <- function(df, wb, spacer = 0, Tables_length = 13, n_cols = 34, header_length = 2){
  
  # Tables Data
  n_cols <- ncol(df)
  n_rows = Tables_length + header_length
  
  
  # Create styles
  header_style <- createStyle(
    fontColour = "white", fgFill = "#004d26", halign = "center", valign = "center",
    border = "TopBottomLeftRight", borderColour = "white", wrapText = TRUE
  )
  
  label_style <- createStyle(
    fontColour = "black", halign = "center", valign = "center", wrapText = TRUE, textDecoration = "bold"
  )
  
  number_style <- createStyle(fgFill =  "#E6E6E6", numFmt = "#,##0",  borderColour = "white", border = "TopBottomLeftRight")
  
  
  # Prepare header information 
  classes <- names(df_wide)[!grepl("SE%|Habitat|region", names(df_wide))]
  
  col1 <- c(as.character(df[(1+spacer),1]))
  row1 <- c("Habitat Type", rep(classes, each = 2))
  row2 <- c("", rep(c("Area (ha)", "SE%"), length(classes)))
  
  # Create header
  writeData(wb, "Tables", t(row1), startCol = 2, startRow = 1 + spacer, colNames = FALSE)
  writeData(wb, "Tables", t(row2), startCol = 2, startRow = 2 + spacer, colNames = FALSE)
  
  # Merge first column ("Habitat Type") vertically
  mergeCells(wb, "Tables", cols = 2, rows = (1 + spacer):(2 + spacer))
  
  # Merge every pair of columns for classes
  start_col <- 3
  for (i in seq_along(classes)) {
    mergeCells(wb, "Tables", cols = start_col:(start_col + 1), rows = 1 + spacer)
    start_col <- start_col + 2
  }
  
  # Format header rows
  addStyle(wb, "Tables", header_style, rows = (1 + spacer):(2 + spacer), cols = 2:ncol(df_GB_wide), gridExpand = TRUE)
  
  # Format Habitat column
  addStyle(wb, "Tables", header_style, rows = (2 + spacer):(n_rows + spacer) , cols = 2, gridExpand = TRUE)
  
  
  # Format numeric columns
  num_cols <- which(sapply(df_GB_wide, is.numeric))
  addStyle(wb, "Tables", number_style, rows = (3 + spacer):(n_rows + spacer), cols = num_cols, gridExpand = TRUE)
  
  # Create a grey bold style 
  grey_fill <- createStyle(textDecoration = "bold", fgFill = "#CCCCCC", borderColour = "white", border = "TopBottomLeftRight", numFmt = "#,##0")  
  
  # Apply style to bottom row and two rightmost columns
  addStyle(wb, "Tables", style = grey_fill,
           rows = (n_rows + spacer):(n_rows + spacer), cols = 3:n_cols, gridExpand = TRUE)
  
  addStyle(wb, "Tables", style = grey_fill,
           rows = (3 + spacer):(n_rows + spacer), cols = (n_cols-1):n_cols, gridExpand = TRUE)
  
  # Adjust column widths
  setColWidths(wb, "Tables", cols = 2, widths = 50)
  setColWidths(wb, "Tables", cols = 3:n_cols, widths = 10)
  setRowHeights(wb, "Tables", rows = (1 + spacer):(2 + spacer), heights = 30)
  setRowHeights(wb, "Tables", rows = (1 + spacer), heights = 60) 
  
  # Apply orange colour to standard error values >25
  for (col in seq(4, n_cols, by = 2)) {
    conditionalFormatting(
      wb,
      sheet = "Tables",
      cols = col,
      rows = (3 + spacer):(n_rows + spacer),
      rule = ">25",
      style = createStyle(fontColour = "orange", numFmt = "#,##0", borderColour = "white", border = "TopBottomLeftRight")
    )
  }
  
  # Apply orange colour to the values which have a related standard error value >25
  for (col in seq(3, n_cols, by = 2)) {
    conditionalFormatting(
      wb,
      sheet = "Tables",
      cols = col,
      rows = (3 + spacer):(n_rows + spacer),
      rule = "OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())), 0, 1) > 25",
      style = createStyle(fontColour = "orange", numFmt = "#,##0", borderColour = "white", border = "TopBottomLeftRight"),
      type = "expression"
    )
  }
  
  # Adds blank rows to worksheet
  writeData(wb, "Tables", x = rep("", n_rows + 1 + spacer), startRow = 2 + spacer, startCol = 1)
  
  # Label Tables
  writeData(wb, "Tables", t(col1), startCol = 1, startRow = 1 + spacer, colNames = FALSE)
  
  # Format Table labels
  addStyle(wb, "Tables", label_style, rows = 1+spacer, cols = 1, gridExpand = TRUE)
  
  return(wb)
}

# Applying formatting iteratively to each table (one table being the data for a particular country or region)
apply_styles <- function(df, wb, Tables_length, n_cols = 34) {
  for (i in 1:17) {
    spacer <- (i-1)*(Tables_length + 4) + 2 # To create empty space between tables and prevent overlapping
    wb <- beautify_Tables(df, wb, spacer, Tables_length, n_cols)
  }
  return(wb)
}

# ---- Transforming tables ----

# All regions
df_wide <- pivot_pipe(df)
df_wide <- sum_pipe(df_wide, groups = c("region", "Habitat"))

# Add total row for each region, this method differs to our Total row pipe function
region_totals <- df_wide |>
  group_by(region) |> 
  summarise(across(where(is.numeric), sum)) |>
  mutate(Habitat = "Total")

df_wide <- bind_rows(df_wide, region_totals)
df_wide$country = NULL 

#Do se calculations for all regions
df_wide <- se_pipe(df_wide)
df_wide <- total_col_pipe(df_wide)
df_wide <- total_se_pipe(df_wide)

# Re-order and rename columns 
df_wide <- column_fix_pipe(df_wide, all_regions = TRUE)

# Whole of GB
df_GB_wide <- full_pipe(df)

# Scotland
df_Scotland_wide <- full_pipe(df_Scotland)

# England
df_England_wide <- full_pipe(df_England)

# Make proportions tables
df_GB_prop <- prop_pipe(df_GB_wide)
df_Scotland_prop <- prop_pipe(df_Scotland_wide)
df_England_prop <- prop_pipe(df_England_wide)
df_wide_prop <- prop_pipe(df_wide) 
df_Wales_prop <- df_wide_prop[df_wide_prop$region == "W", ]

# Preparing for stacking 
df_GB_wide <- cbind(region = "GB", df_GB_wide)
df_Scotland_wide <- cbind(region = "Scotland", df_Scotland_wide)
df_England_wide <- cbind(region = "England", df_England_wide)

# Combining into a single data frame
df_final = bind_rows(list(df_GB_wide, df_Scotland_wide, df_England_wide, df_wide))

# Creating a row of NAs for use in formatting of excel sheet. NOTE: THESE NA VALUES ARE NOT AN ISSUE WITH THE DATA
na_row <- as.data.frame(as.list(rep(NA, ncol(df_final))))
names(na_row) <- names(df_final)

# In the next few lines we add 4 rows of NAs between each region to separate the tables from one another in the final formatted excel sheet
regions = unique(df_final$region)
regions <- c(regions, regions, regions, regions)

for (reg in regions) {
  # Add NA row
  df_final <- rbind(df_final, na_row)
  
  # Assign region value to the newly added row
  new_row_index <- nrow(df_final)
  df_final$region[new_row_index] <- reg
}

# Ordering by GB, country then Region 
priority_order <- c("GB", "Scotland", "England", "W")
df_final <- df_final |>
  mutate(region = fct_relevel(region, priority_order)) |>
  arrange(region)

# Change the rest of the regions to full names
df_final <- df_final |>
  mutate(region = as.character(region)) |>
  mutate(
    region = map_chr(region, ~ region_map[[.x]])
  )

df_final <- rbind(na_row, na_row, df_final)

# ---- Formatting to produce display tables ---- 

library(openxlsx)
library(openxlsx2)

Tables_length = nrow(df_GB_wide) # The length of each individual table (not the length of the full df)

# Full workbook with tables
wb_final <- make_wb(df_final) # Create workbook with data 
wb_final <- apply_styles(df_final, wb_final, Tables_length) # Add formatting

# Resolve the category label using the data_type detected from group2 content at the top of the script.
# For non-standard types, the label is looked up directly from data_type_label.
# For standard types, the first word of the 5th column name is used as a key into name_hash.
if (data_type != "standard") {
  category_label <- data_type_label[[data_type]]
} else {
  identifier     <- sub(" .*", "", names(df_final)[5]) # First word of the variable column name
  category_label <- name_hash[[identifier]]
}

# Add a label describing the condition metric in the tables sheet
label_style <- createStyle(
  fontColour = "black", halign = "center", valign = "center", wrapText = TRUE, textDecoration = "bold", fontSize = 20
)

writeData(wb_final, "Tables",  x = category_label, startRow = 1, startCol = 3)
mergeCells(wb_final, "Tables", cols = 3:(length(names(df_final))-2), rows = 1:2)
addStyle(wb_final, "Tables", label_style, rows = 1:2, cols = 3:(length(names(df_final)) +1), gridExpand = TRUE)

# ---- Plots ----

# This section is responsible for generating the proportions plots

# Create copies of data frames with countries as names
GB <- df_GB_prop
Scotland <- df_Scotland_prop
England <- df_England_prop
Wales <- df_Wales_prop
Wales$region <- NULL

# Plot each of these tables in a loop
tables_to_graph = list(GB, Scotland, England, Wales)
colours = c("#004d26", "#00008b", "#008000", "#ff0000") # Added by Gilly so each country gets their own colour
n = 1

for (i in 1:4){
  
  df <- tables_to_graph[[i]]
  name_suffix = c("GB", "Scotland", "England", "Wales")[i]
  
  
  df_t <- as.data.frame(t(df))                       # Transpose the data in order to work with ggplot
  colnames(df_t) <- df_t[1, ]                        # Creates column names according to the first row of the data (i.e. your habitat column becomes the columnames header)
  df_t <- df_t[-1, ]                                 # Removes row which contains the column names so we can convert to numeric data
  df_t[] <- lapply(df_t, as.numeric)                 # Convert to numeric data to allow plotting
  df_t <- df_t[-nrow(df_t), ]                        # Remove the total row since this will always be 100 (or 0)
  df_t <- rownames_to_column(df_t, var = "Category") # Turns row names into a column called category to allow labeling of plots
  
  # Convert to long format for compatability with ggplot
  df_long <- pivot_longer(
    df_t,
    cols = -Category,              # All columns except Category
    names_to = "Habitat",          # New column to store former column names
    values_to = "Value"            # New column to store values
  )
  
  df_long$Category <- factor(df_long$Category, levels = unique(df_long$Category)) # Preserve the order of categories 
  
  p <- ggplot(df_long, aes(x = Category, y = Value)) +
    geom_bar(stat = "identity", fill = colours[i]) + # Added by Gilly so each country gets their own colour
    geom_text(aes(label = round(Value, 0)), hjust = -0.1, color = "black", size = 3) + 
    facet_wrap(~ Habitat, scales = "fixed", nrow =1, labeller = label_wrap_gen(width = 30)) + # Makes an individual plot per habitat 
    coord_flip() +
    theme_minimal() +
    scale_y_continuous(expand = expansion(mult = c(0, 0.11))) +
    theme(
      panel.background = element_rect(fill = "white"),  # plot panel background
      plot.background = element_rect(fill = "white")    # overall plot background
    ) +
    labs(
      x = category_label,  
      y = "Proportion of area",
      title = name_suffix
    )
  
  
  name = paste0(file_id, "_proportions_", name_suffix, ".png") # I (Llilians Calvo) have changed this line to attempt to write initial file name
  ggsave(name, plot = p, width = 36, height = 4, dpi = 300) # Save as png
}

# Set wb to a copy of wb_final
wb <- unserialize(serialize(wb_final, NULL))

addWorksheet(wb, sheetName = "Graphs")

# Add images of plots to the sheet "Graphs"
insertImage(wb, "Graphs", file = paste0(file_id, "_proportions_GB.png"), startRow = 1,  startCol = 1, width = 36, height = 4)
insertImage(wb, "Graphs", file = paste0(file_id, "_proportions_Scotland.png"), startRow = 23, startCol = 1, width = 36, height = 4)
insertImage(wb, "Graphs", file = paste0(file_id, "_proportions_England.png"), startRow = 45, startCol = 1, width = 36, height = 4)
insertImage(wb, "Graphs", file = paste0(file_id, "_proportions_Wales.png"), startRow = 67, startCol = 1, width = 36, height = 4)


saveWorkbook(wb, name_of_output, overwrite = TRUE) # Save the excel file