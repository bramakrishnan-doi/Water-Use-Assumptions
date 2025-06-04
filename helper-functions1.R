
## Function to read Projected State Use excel data and process it ########
process_excel_data <- function(file_path_arg, sheet_name_arg, data_range_arg) {
  
  excel_data <- tryCatch({
    readxl::read_excel(path = file_path_arg, 
                       sheet = sheet_name_arg, 
                       range = data_range_arg)
  }, 
  error = function(e) {
    stop(paste("Error reading Excel file:", e$message))
    return(NULL) # Should not reach here due to stop()
  })
  
  if (is.null(excel_data) || nrow(excel_data) == 0) {
    warning("Excel data is NULL or empty after reading. Returning NULL.")
    return(NULL)
  }
  # Store original number of columns for checks
  original_ncol <- ncol(excel_data)
  
  # Rename column 2 to "value"
  if (original_ncol >= 2) {
    names(excel_data)[2] <- "value"
  } else {
    warning("Not enough columns to rename column 2 to 'value'. Expected at least 2 columns.")
    # Depending on desired behavior, you might stop or return excel_data as is, or NULL.
    # For now, we'll proceed but subsequent operations might fail or be incorrect.
  }
  
  # Rename column 4 to "notes"
  if (original_ncol >= 4) {
    names(excel_data)[4] <- "notes"
  } else {
    warning("Not enough columns to rename column 4 to 'notes'. Expected at least 4 columns.")
  }
  
  # Remove the original 3rd column
  if (original_ncol >= 3) {
    excel_data <- excel_data[, -3, drop = FALSE] # drop = FALSE ensures it remains a data.frame/tibble
  } else {
    warning("Not enough columns to remove the 3rd column. Expected at least 3 columns.")
  }
  
  # Rename the first column
  if (ncol(excel_data) >= 1 && !is.null(names(excel_data)[1]) && !is.na(names(excel_data)[1])) {
    current_name_col1 <- names(excel_data)[1]
    words <- strsplit(current_name_col1, "\\s+")[[1]]
    words_to_use <- words[words != "" & !is.na(words)] # Remove empty strings and NAs
    if (length(words_to_use) > 0) {
      num_words_to_take <- min(3, length(words_to_use))
      selected_words <- words_to_use[1:num_words_to_take]
      new_name_col1 <- paste(selected_words, collapse = "_")
      names(excel_data)[1] <- new_name_col1
    } else {
      warning("First column name was empty or only whitespace after splitting. Not renaming.")
    }
  } else {
    warning("Could not rename the first column (it may not exist or its name is problematic).")
  }
  
  excel_data = remove_all_na_rows(excel_data)
  
  # # Remove Rows Where All Values are NA ---
  # if (nrow(excel_data) > 0) {
  #   # apply works row-wise (MARGIN = 1)
  #   # For each row, function(row_values) checks if all elements in that row are NA
  #   all_na_rows_logical <- apply(excel_data, 1, function(row_values) all(is.na(row_values)))
  #   excel_data <- excel_data[!all_na_rows_logical, , drop = FALSE]
  # }
  
  # Add Row Numbers ---
  # Add this step only if there are rows remaining after cleaning
  if (nrow(excel_data) > 0) {
    excel_data$Row_Num <- 1:nrow(excel_data)
  }
  
  return(excel_data)
}



## Function to find rows containing a string ####################

find_rows_containing_string <- function(data_tibble, search_term) {
  # --- Input Validations ---
  if (!is.data.frame(data_tibble)) {
    stop("Input 'data_tibble' must be a data frame or tibble.")
  }
  if (!is.character(search_term) || length(search_term) != 1) {
    stop("Input 'search_term' must be a single string.")
  }
  
  # Ensure dplyr is available
  if (!requireNamespace("dplyr", quietly = TRUE)) {
    stop("The 'dplyr' package is required for this function. Please install it using install.packages('dplyr').")
  }
  
  # Handle empty data_tibble
  if (nrow(data_tibble) == 0) {
    # message("Input 'data_tibble' is empty. Returning an empty tibble.")
    return(data_tibble[FALSE, ]) # Return empty tibble with same columns
  }
  
  # Handle empty search_term (optional: could also return all rows or error)
  if (nchar(search_term) == 0) {
    # message("Warning: 'search_term' is an empty string. This may lead to unexpected matches. Returning an empty tibble.")
    return(data_tibble[FALSE, ]) # Or handle as per desired behavior
  }
  
  # --- Identify Character Columns ---
  # We'll try to search in columns that are character or can be reasonably coerced to character (like factors)
  # For simplicity, we'll focus on columns dplyr::if_any can process with as.character()
  # Get names of columns that are character or factor
  target_col_names <- names(data_tibble)[sapply(data_tibble, function(col) is.character(col) || is.factor(col))]
  
  
  if (length(target_col_names) == 0) {
    # message("No character or factor columns found in 'data_tibble' to search within. Returning an empty tibble.")
    return(data_tibble[FALSE, ])
  }
  
  # --- Perform Search using dplyr ---
  # The dplyr::if_any function allows checking a condition across multiple columns.
  # For each row, it will be TRUE if the lambda function returns TRUE for any of the specified columns.
  result_tibble <- data_tibble %>%
    dplyr::filter(
      dplyr::if_any(dplyr::all_of(target_col_names), function(column_content) {
        # Inside this function, 'column_content' refers to the vector of values for the current column being processed.
        
        # Convert column to character (handles factors as well) and trim whitespace
        cleaned_content <- trimws(as.character(column_content))
        
        # Perform case-insensitive literal search by converting both to lower case first
        lower_cleaned_content <- tolower(cleaned_content)
        lower_search_term <- tolower(search_term)
        
        # Now perform a fixed (literal) search on the lowercased strings.
        # ignore.case is no longer needed here and fixed = TRUE can be safely used.
        matches <- grepl(lower_search_term, lower_cleaned_content, fixed = TRUE)
        
        # Important: grepl returns NA if the content is NA.
        # For filtering, we want NA content to result in FALSE (no match).
        matches[is.na(matches)] <- FALSE
        
        matches # Return the logical vector of matches for this column
      })
    )
  
  return(result_tibble)
}

## Remove NA rows in dataframe ################################################
remove_all_na_rows <- function(input_tibble) {
  # Check if the input is a tibble or data frame
  if (!is.data.frame(input_tibble)) {
    stop("Input must be a tibble or data frame.")
  }
  
  # Use filter with across() and a helper function to check if all values in a row are NA
  # The `!.all(is.na(.))` part means "not all values are NA" for the selected columns.
  # `across(everything())` applies this check to all columns.
  cleaned_tibble <- input_tibble %>%
    filter(!rowSums(is.na(across(everything()))) == ncol(input_tibble))
  # An alternative using if_all (requires dplyr 1.0.0 or later):
  # filter(!if_all(everything(), is.na))
  
  
  return(cleaned_tibble)
}

## Function to condition replace items #########################################

replace_items_conditional <- function(data_tibble,
                                      lookup_table,
                                      col_to_replace = 1,
                                      key_col_in_lookup = 1,
                                      value_col_in_lookup = 2) {
  
  # Validate col_to_replace
  if (is.character(col_to_replace)) {
    if (!(col_to_replace %in% names(data_tibble))) {
      stop(paste("Error: Column '", col_to_replace, "' not found in data_tibble.", sep=""))
    }
    # Convert column name to index for internal use
    col_idx_to_replace <- match(col_to_replace, names(data_tibble))
  } else if (is.numeric(col_to_replace)) {
    if (col_to_replace < 1 || col_to_replace > ncol(data_tibble)) {
      stop("Error: 'col_to_replace' index is out of bounds for data_tibble.")
    }
    col_idx_to_replace <- as.integer(col_to_replace)
  } else {
    stop("Error: 'col_to_replace' must be a column name (character) or index (numeric).")
  }
  
  # --- Prepare Lookup Table ---
  # Convert lookup_table to a standardized named vector for efficient lookup
  if (is.data.frame(lookup_table) || is_tibble(lookup_table)) {
    if (ncol(lookup_table) < 2 && (is.numeric(key_col_in_lookup) && is.numeric(value_col_in_lookup) && (key_col_in_lookup > ncol(lookup_table) || value_col_in_lookup > ncol(lookup_table)))) {
      stop("Error: If 'lookup_table' is a data frame/tibble, it must have at least two columns, or valid key/value column specifiers.")
    }
    # Handle key_col_in_lookup
    if (is.character(key_col_in_lookup)) {
      if (!(key_col_in_lookup %in% names(lookup_table))) stop(paste("Key column '", key_col_in_lookup, "' not found in lookup_table."))
      keys <- lookup_table[[key_col_in_lookup]]
    } else if (is.numeric(key_col_in_lookup)) {
      if (key_col_in_lookup < 1 || key_col_in_lookup > ncol(lookup_table)) stop("Key column index out of bounds for lookup_table.")
      keys <- lookup_table[[as.integer(key_col_in_lookup)]]
    } else {
      stop("Error: 'key_col_in_lookup' must be a column name or index.")
    }
    
    # Handle value_col_in_lookup
    if (is.character(value_col_in_lookup)) {
      if (!(value_col_in_lookup %in% names(lookup_table))) stop(paste("Value column '", value_col_in_lookup, "' not found in lookup_table."))
      values <- lookup_table[[value_col_in_lookup]]
    } else if (is.numeric(value_col_in_lookup)) {
      if (value_col_in_lookup < 1 || value_col_in_lookup > ncol(lookup_table)) stop("Value column index out of bounds for lookup_table.")
      values <- lookup_table[[as.integer(value_col_in_lookup)]]
    } else {
      stop("Error: 'value_col_in_lookup' must be a column name or index.")
    }
    
    # Ensure keys and values are of the same length
    if(length(keys) != length(values)) {
      stop("Error: Key and value columns in lookup_table must have the same length.")
    }
    
    # Create the named vector, converting keys to character for robust matching
    internal_lookup_vec <- stats::setNames(as.character(values), as.character(keys))
    
  } else if (is.vector(lookup_table) && !is.null(names(lookup_table))) {
    # It's already a named vector, ensure values and names are character for consistency
    internal_lookup_vec <- stats::setNames(as.character(lookup_table), as.character(names(lookup_table)))
  } else {
    stop("Error: 'lookup_table' must be a named vector, data frame, or tibble.")
  }
  
  # --- Perform Replacement ---
  # Make a copy to avoid modifying the original tibble by reference (tibbles are copy-on-modify)
  # but explicit copy is clearer.
  output_tibble <- data_tibble
  
  # Get the actual column data to be modified
  original_column_data <- output_tibble[[col_idx_to_replace]]
  
  # Convert original data to character for matching, if it's not already.
  # This ensures that numeric keys in lookup_table (if provided as characters)
  # can match numeric values in the data_tibble column.
  items_to_check <- as.character(original_column_data)
  
  # Get the names (keys) from the lookup vector
  lookup_keys <- names(internal_lookup_vec)
  
  # Perform the replacement
  # Iterate through each item in the target column
  new_column_data <- sapply(items_to_check, function(item) {
    if (item %in% lookup_keys) {
      # If a match is found, return the replacement value
      # Use match to get the first corresponding value if there are duplicate keys in lookup_vec
      # (though setNames typically handles this by taking the last)
      return(internal_lookup_vec[item])
    } else {
      # If no match, return the original item
      # We need to ensure the type is preserved if no replacement happens.
      # This is tricky because we converted items_to_check to character.
      # We'll find the original item's position and return it.
      # A safer way is to build the new column based on original_column_data's type.
      return(NA) # Placeholder for items not found
    }
  }, USE.NAMES = FALSE)
  
  # Coalesce: take new_column_data where it's not NA (i.e., where a replacement occurred),
  # otherwise take the original data.
  # This also helps in preserving the original data type if no replacement happened.
  final_column_values <- ifelse(is.na(new_column_data), original_column_data, new_column_data)
  
  # Try to coerce back to original type if possible, or a common compatible type
  # This part can be complex if types are very different.
  # For simplicity, if original was factor, try to make new one factor.
  # If original was numeric and new values are all numeric-coercible, convert.
  if (is.factor(original_column_data)) {
    output_tibble[[col_idx_to_replace]] <- factor(final_column_values, levels = unique(c(levels(original_column_data), final_column_values)))
  } else if (is.numeric(original_column_data) && !is.character(original_column_data) && all(!is.na(suppressWarnings(as.numeric(final_column_values))))) {
    # Check if all final values can be numbers (even if they came from character lookup values)
    # And ensure original was not already character (which can look numeric)
    if (all(final_column_values == suppressWarnings(as.character(as.numeric(final_column_values))))) { # Check if they are "clean" numbers
      output_tibble[[col_idx_to_replace]] <- as.numeric(final_column_values)
    } else {
      output_tibble[[col_idx_to_replace]] <- as.character(final_column_values) # Fallback to character if mixed
    }
  } else if (is.integer(original_column_data) && all(!is.na(suppressWarnings(as.integer(final_column_values))))) {
    if (all(final_column_values == suppressWarnings(as.character(as.integer(final_column_values))))) {
      output_tibble[[col_idx_to_replace]] <- as.integer(final_column_values)
    } else {
      output_tibble[[col_idx_to_replace]] <- as.character(final_column_values) # Fallback to character
    }
  }
  else {
    # Default to character if types are mixed or complex, or if lookup values introduce non-numeric strings
    # This is the safest bet if the replacement values are strings.
    output_tibble[[col_idx_to_replace]] <- as.character(final_column_values)
  }
  
  return(output_tibble)
}


####### Function to compute AZ reduction scenarios #########

calculate_AZ_reductions <- function(year, sct_data) {
  # Assuming 'find_rows_containing_string' is a pre-defined function
  # and 'sct_data' is available in the environment or passed as an argument.
  
  year = year +1 
  
  x <- find_rows_containing_string(sct_data, "CAP Annual Shortage Volume")[year]
  y <- find_rows_containing_string(sct_data, "AnnualDCPContribution_AZ")[year]
  z <- find_rows_containing_string(sct_data, "AnnualDeliveryEC_AZ")[year]
  
  # Ensure x, y, z are numeric and handle potential NA values by coercing to 0
  x <- ifelse(is.na(x), 0, as.numeric(x))
  y <- ifelse(is.na(y), 0, as.numeric(y))
  z <- ifelse(is.na(z), 0, as.numeric(z))
  
  condition1_met <- x > 0
  condition2_met <- y > 0
  condition3_met <- z > 0
  
  if (x > 0 | y > 0 | z > 0) {
    addreductions <- TRUE
  } else {
    addreductions <- FALSE
  }
  
  # --- Define the phrases for each condition ---
  phrase1_text <- paste0("Shortage volume of ", round(x / 1000, 0), " kaf")
  phrase2_text <- paste0("DCP contribution of ", round(y / 1000, 0), " kaf by CAWCD")
  phrase3_text <- paste0("ICS delivery of ", round(z / 1000, 1), " kaf")
  
  active_phrases <- character(0)
  if (condition1_met) {
    active_phrases <- c(active_phrases, phrase1_text)
  }
  if (condition2_met) {
    active_phrases <- c(active_phrases, phrase2_text)
  }
  if (condition3_met) {
    active_phrases <- c(active_phrases, phrase3_text)
  }
  
  num_active_conditions <- length(active_phrases)
  reductions_scen_all <- ""
  
  if (num_active_conditions == 1) {
    reductions_scen_all <- active_phrases[1]
  } else if (num_active_conditions == 2) {
    reductions_scen_all <- paste0(active_phrases[1], ", and ", active_phrases[2])
  } else if (num_active_conditions == 3) {
    reductions_scen_all <- paste0(active_phrases[1], ", ", active_phrases[2], ", and ", active_phrases[3])
  }
  
  # Return the two values as a list
  return(list(addreductions = addreductions, reductions_scen_all = reductions_scen_all))
}

## Get SCT data ###################################################
if (!file.exists(file_path2)) {
  stop(paste("Error: The file was not found at the specified path:", file_path2))
}

sct_data <- read_excel(path = file_path2, sheet = sheet_name4, range = data_range4)

sct_data <- sct_data[, -2]

if (ncol(sct_data) == 4) {
  # Get the current year
  current_year <- year(Sys.Date()) # Using lubridate's year() function
  
  # Define the new column names
  new_column_names <- c(
    "Item",
    paste0("CY", current_year),
    paste0("CY", current_year + 1),
    paste0("CY", current_year + 2)
  )
  
  # Assign the new names to excel_data2
  names(sct_data) <- new_column_names
  
} else {
  warning(paste("Warning: sct_data does not have the expected 4 columns after deletion. It has", 
                ncol(sct_data), "columns. Column renaming might be incorrect."))
}

apportionment <- sct_data %>% 
  filter(
    str_detect(Item,
               regex(
                 "California_Apportionment|.AzTotalAnnual|.NvTotalAnnual"
               )
    )
  ) 


mwdICSY1 = ifelse(find_rows_containing_string(sct_data,"MWD_Default")[1,2],
                  paste0("creation of ", 
                         round(find_rows_containing_string(sct_data,"MWD_Default")[1,2]/1000,1)),
                  paste0("delivery of ", 
                         round(find_rows_containing_string(sct_data,"MWD_Default")[2,2]/1000,1)))

mwdICSY2 = ifelse(find_rows_containing_string(sct_data,"MWD_Default")[1,3],
                  paste0("creation of ", 
                         round(find_rows_containing_string(sct_data,"MWD_Default")[1,3]/1000,1)),
                  paste0("delivery of ", 
                         round(find_rows_containing_string(sct_data,"MWD_Default")[2,3]/1000,1)))

mwdICSY3 = ifelse(find_rows_containing_string(sct_data,"MWD_Default")[1,4],
                  paste0("creation of ", 
                         round(find_rows_containing_string(sct_data,"MWD_Default")[1,4]/1000,1)),
                  paste0("delivery of ", 
                         round(find_rows_containing_string(sct_data,"MWD_Default")[2,4]/1000,1)))

cons_CA <- tryCatch({
  readxl::read_excel(path = file_path, sheet = sheet_cons, range = rangeCA)
}, error = function(e) {
  stop(paste("Error reading Excel file:", e$message))
  return(NULL) # Should not reach here due to stop()
})

## Get Conservation Data ################################################
cons_CA <- remove_all_na_rows(cons_CA)

cons_NonCAWCD <- tryCatch({
  readxl::read_excel(path = file_path, sheet = sheet_cons, range = rangeNonCAWCD)
}, error = function(e) {
  stop(paste("Error reading Excel file:", e$message))
  return(NULL) # Should not reach here due to stop()
})

cons_NonCAWCD <- remove_all_na_rows(cons_NonCAWCD)
#remove 242
cons_NonCAWCD_filt <- cons_NonCAWCD %>% 
  slice_head(n = -1)
cons_NonCAWCD_tot <- cons_NonCAWCD %>% 
  slice_head(n = -1) %>% 
  summarise(across(where(is.numeric), \(x) sum(x, na.rm = TRUE))) 

cons_CAWCD <- tryCatch({
  readxl::read_excel(path = file_path, sheet = sheet_cons, range = rangeCAWCD)
  }, error = function(e) {
  stop(paste("Error reading Excel file:", e$message))
  return(NULL) # Should not reach here due to stop()
})

cons_CAWCD_tot <- cons_CAWCD %>% 
  summarise(across(where(is.numeric), \(x) sum(x, na.rm = TRUE)))

current_year <- format(Sys.Date(), "%Y")
column_number <- which(grepl(current_year, colnames(cons_NonCAWCD)))



text_to_exclude = "CVWD"
cons_CA_filt <- cons_CA %>%
  filter(is.na(Contractor) | !str_detect(Contractor, fixed(text_to_exclude)))

cvwd <- find_rows_containing_string(cons_CA, text_to_exclude)
new_item <- c("Contractor" = "CVWD")
new_item <- c(new_item, colSums(cvwd[,2:ncol(cvwd)]))

cons_CA_filt <- cons_CA %>%
  filter(is.na(Contractor) | !str_detect(Contractor, fixed(text_to_exclude)))

cons_CA_filt <- rbind(tibble::as_tibble(as.list(new_item)), cons_CA_filt)

name_lookup <- c(
  "IID 2023-2026" = "IID",
  "GM Gabrych/Matador" = "Gabrych",
  "Metro Water Dist" = "Metro Water District",
  "Spanish Trail Water Company" = "Spanish Trail Water Co."
)

updated_cons_CA_filt <- replace_items_conditional(cons_CA_filt , name_lookup)

cons_CA_CY <- updated_cons_CA_filt %>% 
  filter(.[[column_number]] > 0) %>% 
  select(1,all_of(column_number))

cons_CA_CY2 <- updated_cons_CA_filt %>% 
  filter(.[[column_number+1]] > 0) %>% 
  select(1,all_of(column_number+1))

cons_CA_CY3 <- updated_cons_CA_filt %>% 
  filter(.[[column_number+2]] > 0) %>% 
  select(1,all_of(column_number+2))

cons_CAWCD <- remove_all_na_rows(cons_CAWCD)

cons_CAWCD_CY <- replace_items_conditional(cons_CAWCD, name_lookup) %>% 
  filter(.[[column_number]] > 0) %>% 
  select(1,all_of(column_number))

cons_NonCAWCD_CY <- cons_NonCAWCD_filt %>% 
  replace_items_conditional(name_lookup) %>% 
  filter(.[[column_number]] > 0) %>% 
  select(1,all_of(column_number))

cons_CAWCD_CY2 <- replace_items_conditional(cons_CAWCD, name_lookup) %>% 
  filter(.[[column_number+1]] > 0) %>% 
  select(1,all_of(column_number+1))

cons_NonCAWCD_CY2 <- cons_NonCAWCD_filt %>% 
  replace_items_conditional(name_lookup) %>% 
  filter(.[[column_number+1]] > 0) %>% 
  select(1,all_of(column_number+1))

cons_CAWCD_CY3 <- replace_items_conditional(cons_CAWCD, name_lookup) %>% 
  filter(.[[column_number+2]] > 0) %>% 
  select(1,all_of(column_number+2))

cons_NonCAWCD_CY3 <- cons_NonCAWCD_filt %>% 
  replace_items_conditional(name_lookup) %>% 
  filter(.[[column_number+2]] > 0) %>% 
  select(1,all_of(column_number+2))

## Create ICS Table ################################################
ics_data <- find_rows_containing_string(sct_data, "Bank_")

transformed_data <- ics_data %>%
  mutate(Item = stringr::str_sub(Item, -2)) %>%
  rename_with(~ifelse(. == "Item", "State (volumes in AF)", sub("^CY", "", .))) %>%
  janitor::adorn_totals(where = "row", fill = "-", na.rm = TRUE, name = "Total")

# --- Create and Style the flextable for DOCX ---
# Identify numeric columns by name (excluding the 'State' column)
numeric_cols <- setdiff(names(transformed_data), "State (volumes in AF)")

ftics <- flextable(transformed_data)

ftics <- colformat_num(
  x = ftics,
  j = numeric_cols,
  big.mark = ",",
  digits = 0,
  na_str = ""
)

ftics <- bold(ftics, part = "header", bold = TRUE) # Bolds the entire header
# Bolds the total row
ftics <- bold(ftics, i = nrow(transformed_data), part = "body", bold = TRUE) 

border_style <- fp_border(color = "black", width = 1)
ftics <- border_remove(x = ftics)
ftics <- border_outer(x = ftics, part = "all", border = border_style)
ftics <- border_inner_h(x = ftics, part = "all", border = border_style)
ftics <- border_inner_v(x = ftics, part = "all", border = border_style)

ftics <- padding(ftics, padding = 2, part = "all") # Reduced cell padding

# Center align header text for numeric columns
ftics <- align(
  x = ftics,
  j = numeric_cols,   # Target only the numeric columns
  align = "center",   # Set alignment to center
  part = "all"     # Apply to the header part only
)

# Set font size for all cells in the table to 10pt
ftics <- fontsize(
  x = ftics,
  size = 10,        # Font size in points
  part = "all"      # Apply to header and body
)


ftics <- autofit(ftics) # Autofit after all styling is applied


## Conservation Table #######################################

# --- 2. Define File Path and Import Data ---
file_path <- "data/Projected State Use -APR25.xlsx"
sheet_name <- "SysCon SummaryTable"
excel_range <- "C79:I103"

raw_excel_data <- readxl::read_excel(
  path = file_path,
  sheet = sheet_name,
  range = excel_range,
  col_names = FALSE # Read without initial column names
)

# --- 3. Prepare Data ---
selected_raw_data_indices <- c(2, 3, 4, 5, 6, 7) 
df_data_rows <- raw_excel_data[-1, selected_raw_data_indices]

clean_col_names <- c("State", "Conservation Activity", 
                     "2025", "2026", "2027", "Total")
colnames(df_data_rows) <- clean_col_names
df <- df_data_rows

numeric_cols <- c("2025", "2026", "2027", "Total")

df <- df %>%
  mutate(across(all_of(numeric_cols), 
                ~ceiling(suppressWarnings(as.numeric(as.character(.))))
  ))

num_data_rows <- nrow(df)
annual_total_row_df_idx <- num_data_rows - 1 
cumulative_total_row_df_idx <- num_data_rows 

annual_total_label <- raw_excel_data[1 + annual_total_row_df_idx, 
                                     1, drop = TRUE] 
cumulative_total_label <- raw_excel_data[1 + cumulative_total_row_df_idx, 
                                         1, drop = TRUE]

if (!is.na(annual_total_label) && annual_total_label == "Annual Total") {
  df[annual_total_row_df_idx, "Conservation Activity"] <- annual_total_label
  df[annual_total_row_df_idx, "State"] <- ""
}
if (!is.na(cumulative_total_label) && 
    cumulative_total_label == "Cumulative Total") {
  df[cumulative_total_row_df_idx, "Conservation Activity"] <- cumulative_total_label
  df[cumulative_total_row_df_idx, "State"] <- ""
}

# --- 4. Create Flextable ---
ft <- flextable(df)

# --- 5. Apply Styling ---
fp_border_thin <- fp_border(color = "black", width = 1)
fp_border_less_thick <- fp_border(color = "black", width = 1.5) 

light_gray_bg_header <- "#F0F0F0" 
zebra_gray_bg <- "#F5F5F5"     

ft <- theme_zebra(ft, odd_body = "transparent", even_body = zebra_gray_bg)
ft <- bg(ft, j = "State", bg = "transparent", part = "body") 

# A. Header Styling
ft <- bold(ft, part = "header", bold = TRUE)
ft <- align(ft, align = "center", part = "header")
ft <- bg(ft, bg = light_gray_bg_header, part = "header") 

# B. State Column Styling
state_merge_groups <- list(1:10, 11:18, 19:20, 21:22) 
for (rows_to_merge in state_merge_groups) {
  ft <- merge_at(ft, i = rows_to_merge, j = "State", part = "body")
}
if (length(annual_total_row_df_idx) > 0 && 
    length(cumulative_total_row_df_idx) > 0 &&
    cumulative_total_row_df_idx == annual_total_row_df_idx + 1) {
  ft <- merge_at(ft, i = annual_total_row_df_idx:cumulative_total_row_df_idx, 
                 j = "State", part = "body")
}

ft <- bold(ft, j = "State", bold = TRUE, part = "body") 
ft <- align(ft, j = "State", align = "center", part = "body")    
ft <- valign(ft, j = "State", valign = "center", part = "body")   
# Text rotation abandoned

# C. Numeric Column Formatting & Alignment
ft <- colformat_num(ft, j = numeric_cols, big.mark = ",", 
                    digits = 0, na_str = "") 
ft <- align(ft, j = numeric_cols, align = "right", part = "body")
if(length(annual_total_row_df_idx) > 0) {
  ft <- align(ft, i = annual_total_row_df_idx, 
              j = numeric_cols, align = "right", part = "body")
}
if(length(cumulative_total_row_df_idx) > 0) {
  ft <- align(ft, i = cumulative_total_row_df_idx, 
              j = numeric_cols, align = "right", part = "body")
}

# D. "Annual Total" & "Cumulative Total" Row Styling
if(length(annual_total_row_df_idx) > 0) {
  ft <- bold(ft, i = annual_total_row_df_idx, bold = TRUE, part = "body")
  ft <- align(ft, i = annual_total_row_df_idx, 
              j = "Conservation Activity", align = "left", part = "body")
}
if(length(cumulative_total_row_df_idx) > 0) {
  ft <- bold(ft, i = cumulative_total_row_df_idx, bold = TRUE, part = "body")
  ft <- align(ft, i = cumulative_total_row_df_idx, 
              j = "Conservation Activity", align = "left", part = "body")
}

# E. General Borders & Specific Thick Horizontal Lines
ft <- border_remove(ft) # Clear all existing borders first

# Apply general inner borders (thin)
ft <- border_inner_h(ft, border = fp_border_thin, part = "all") 
ft <- border_inner_v(ft, border = fp_border_thin, part = "all") 

# Apply outer border for the whole table (using new "less thick")
ft <- border_outer(ft, border = fp_border_less_thick, part = "all") 

# Re-apply specific less_thick hlines 
# (these will override inner thin ones where they apply)
ft <- hline_bottom(ft, part = "header", border = fp_border_less_thick) 

hline_group_separator_indices <- c(10, 18, 20, 22) 
ft <- hline(ft, i = hline_group_separator_indices, 
            border = fp_border_less_thick, part = "body")

# Ensure bottom border of the entire last data row is less_thick
if (nrow(df) > 0) {
  ft <- hline(ft, i = cumulative_total_row_df_idx, 
              border = fp_border_less_thick, part = "body")
}


# F. Fontsize, Padding, and Line Spacing
ft <- fontsize(ft, size = 9, part = "all")
ft <- padding(ft, padding.top = 2, padding.bottom = 2, 
              padding.left = 4, padding.right = 4, part = "all")
ft <- line_spacing(ft, space = 1.2, part = "all") 

# G. Initial Widths & Autofit
if (ncol(df) == 6) {
  ft <- width(ft, j = "State", width = 0.5) 
  ft <- width(ft, j = "Conservation Activity", width = 3.2)
  ft <- width(ft, j = c("2025", "2026", "2027", "Total"), width = 0.7)
}
#ft <- autofit(ft)

