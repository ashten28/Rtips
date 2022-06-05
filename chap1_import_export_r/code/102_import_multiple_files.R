#***********************************
# Reading and appending multiple workbooks
# with identical data columns
# to one data frame
#***********************************

# Loading libraries
library(readxl)

# List all workbooks in folder
wb_list <- list.files(path = "chap1_import_export_r/data")

# Pick only workbooks to read - only files that starts with 001b
wb_list_fix <- wb_list[grepl("^102", wb_list)]

# Create function to read one workbook
func_read_workbook <- function(x) {
  
  # Set file path for the workbook with multiple sheet
  wb_path <- paste0("chap1_import_export_r/data/", x)
  
  # Get first sheet
  wb_sheets_first <- excel_sheets(wb_path)[1]
  
  # Read workbook
  df <- read_excel(path = wb_path, sheet = wb_sheets_first)
  
  # Create LOB column using sheet name
  df$LOB <- wb_sheets_first
  
  # Return final data frame
  return(df)
}

# Read all workbooks into a list of data frames by iterating the function
wb_data_ls <- 
  lapply(
    X = wb_list_fix,
    FUN = func_read_workbook
  )

# Append all data frames into one data frame
wb_data_df <- do.call(rbind, wb_data_ls) 

# Write output as csv
write.csv(
  x = wb_data_df, 
  file = "chap1_import_export_r/output/102_output_data.csv",
  row.names = FALSE
  )


