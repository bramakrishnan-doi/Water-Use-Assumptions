# Function to read Projected State Use excel data and process it
process_excel_data <- function(file_path_arg, sheet_name_arg, data_range_arg) {
  
  excel_data <- tryCatch({
    readxl::read_excel(path = file_path_arg, sheet = sheet_name_arg, range = data_range_arg)
  }, error = function(e) {
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

#########################################################################

########## Function to find rows containing a string ####################

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

################################################################
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
#####################################################

######################################################

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

###########################################################

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

#################################