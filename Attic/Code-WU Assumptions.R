library(tidyverse)
library(readxl)


####### Function to read Proj. State Use excel data ############

process_excel_data <- function(file_path_arg, sheet_name_arg, data_range_arg) {
 
  if (!requireNamespace("readxl", quietly = TRUE)) {
    stop("The 'readxl' package is required. Please install it using install.packages('readxl').")
  }
  
  # --- 1. Read Excel Data ---
  # It's good practice to check if the file exists
  if (!file.exists(file_path_arg)) {
    stop(paste("Error: File not found at", file_path_arg))
  }
  
  # Use tryCatch for robust error handling during file reading
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
  
  # --- 2. Column Manipulations ---
  # Store original number of columns for checks
  original_ncol <- ncol(excel_data)
  # print(paste("Original number of columns read:", original_ncol))
  # print("Original column names:")
  # print(names(excel_data))
  
  # A. Rename column 2 to "value"
  #    This assumes the input range results in at least 2 columns.
  if (original_ncol >= 2) {
    names(excel_data)[2] <- "value"
  } else {
    warning("Not enough columns to rename column 2 to 'value'. Expected at least 2 columns.")
    # Depending on desired behavior, you might stop or return excel_data as is, or NULL.
    # For now, we'll proceed but subsequent operations might fail or be incorrect.
  }
  
  # B. Rename column 4 to "notes"
  #    This assumes the input range results in at least 4 columns.
  if (original_ncol >= 4) {
    names(excel_data)[4] <- "notes"
  } else {
    warning("Not enough columns to rename column 4 to 'notes'. Expected at least 4 columns.")
  }
  
  # C. Remove the original 3rd column
  #    This also assumes the input range results in at least 3 columns.
  if (original_ncol >= 3) {
    excel_data <- excel_data[, -3, drop = FALSE] # drop = FALSE ensures it remains a data.frame/tibble
    # print("Column names after renaming 2 & 4, and deleting original 3rd:")
    # print(names(excel_data))
  } else {
    warning("Not enough columns to remove the 3rd column. Expected at least 3 columns.")
  }
  
  # After deleting a column, the number of columns has changed.
  # The first column is still at index 1.
  
  # D. Rename the first column
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
  # print("Column names after all renaming operations:")
  # print(names(excel_data))
  
  # --- 3. Remove Rows Where All Values are NA ---
  if (nrow(excel_data) > 0) {
    # apply works row-wise (MARGIN = 1)
    # For each row, function(row_values) checks if all elements in that row are NA
    all_na_rows_logical <- apply(excel_data, 1, function(row_values) all(is.na(row_values)))
    excel_data <- excel_data[!all_na_rows_logical, , drop = FALSE]
  }
  
  # --- 4. Add Row Numbers ---
  # Add this step only if there are rows remaining after cleaning
  if (nrow(excel_data) > 0) {
    excel_data$Row_Num <- 1:nrow(excel_data)
  }
  
  # print("Final processed data head:")
  # print(head(excel_data))
  return(excel_data)
}




################################################################







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

file_path <- "C:/Temp/Projected State Use -APR25.xlsx"
sheet_name1 <- as.character(year(Sys.Date()))
sheet_name2 <- as.character(year(Sys.Date())+1)
sheet_name3 <- as.character(year(Sys.Date())+2)
data_range1 <- "A1:D58"
data_range2 <- "F1:I58"
data_range3 <- "K1:N58"

#Current Year
Most1 <- process_excel_data(file_path, sheet_name1, data_range1)
Min1 <- process_excel_data(file_path, sheet_name1, data_range2)
Max1 <- process_excel_data(file_path, sheet_name1, data_range3)

#First Out year
Most2 <- process_excel_data(file_path, sheet_name2, data_range1)
Min2 <- process_excel_data(file_path, sheet_name2, data_range2)
Max2 <- process_excel_data(file_path, sheet_name2, data_range3)

#Second Out Year
Most3 <- process_excel_data(file_path, sheet_name3, data_range1)
Min3 <- process_excel_data(file_path, sheet_name3, data_range2)
Max3 <- process_excel_data(file_path, sheet_name3, data_range3)


find_rows_containing_string(Most3, "DCP Contribution")

find_rows_containing_string(Most1, "PSCP")

find_rows_containing_string(excel_data, "Mexico use")$notes
find_rows_containing_string(excel_data, "Tributary")


## Get SCT data ###################################
file_path2 <- "C:/Temp/APR25-Most.xlsx"
sheet_name4 <- "Sheet1"
data_range4 <- "A1:E49"


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

######################################################

apportionment <- sct_data %>% 
  filter(
    str_detect(Item,
               regex(
                 "California_Apportionment|.AzTotalAnnual|.NvTotalAnnual"
                 )
               )
    ) 
  
round(sum(apportionment[1,2])/1000000,3)

sprintf("%.3f", round(as.numeric(find_rows_containing_string(Most1, "Mexico Use")$value)/1000000,3))


round(as.numeric(find_rows_containing_string(sct_data,"MWDDiversionAnnualFC")[2])/1000,0)


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

cons_CA <- remove_all_na_rows(cons_CA)

cons_NonCAWCD <- tryCatch({
  readxl::read_excel(path = file_path, sheet = sheet_cons, range = rangeNonCAWCD)
}, error = function(e) {
  stop(paste("Error reading Excel file:", e$message))
  return(NULL) # Should not reach here due to stop()
})

cons_NonCAWCD <- remove_all_na_rows(cons_NonCAWCD)

cons_CAWCD <- tryCatch({
  readxl::read_excel(path = file_path, sheet = sheet_cons, range = rangeCAWCD)
}, error = function(e) {
  stop(paste("Error reading Excel file:", e$message))
  return(NULL) # Should not reach here due to stop()
})

cons_CAWCD <- cons_CAWCD %>% 
  remove_all_na_rows() %>% 
  slice_head(n = -1) %>% # Removes the last row
  summarise(across(where(is.numeric), \(x) sum(x, na.rm = TRUE))) # Sums only numeric columns


cvwd <- sum(find_rows_containing_string(cons_CA, "CVWD")[column_number])

tot_CA_cons <- round(sum(cons_CA[column_number])/1000,1)


text_to_exclude = "CVWD"
cons_CA_filt <- cons_CA %>%
  filter(is.na(Contractor) | !str_detect(Contractor, fixed(text_to_exclude)))

cons_CA_filt %>% 
  filter(.[[column_number]] > 0) %>% 
  select(1,all_of(column_number))


