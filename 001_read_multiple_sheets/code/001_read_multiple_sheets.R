#***********************************
# Reading and appending multiple sheets
# with identical data columns
# to one data frame
#***********************************

# Loading libraries
library(readxl)

# Set file path for the workbook with multiple sheet
wb_path <- "001_read_multiple_sheets/data/example_data_001.xlsx"

# Get all the sheets in that workbook into a vector
wb_sheets_all <- excel_sheets(wb_path)

# Get only sheets to read - remove sheets 1 i.e Notes
wb_sheets_fix <- wb_sheets_all[2:6]

# Read all sheets from workbook into a list of data frames
wb_data_ls <- 
  lapply(
    X = wb_sheets_fix,
    FUN = function(x) {
      df <- read_excel(path = wb_path, sheet = x)
      df$LOB <- x
      return(df)
      }
  )

# Append all data frames into one data frame
wb_data_df <- do.call(rbind, wb_data_ls) 

# Write output as csv
write.csv(
  x = wb_data_df, 
  file = "001_read_multiple_sheets/output/001_output_data.csv",
  row.names = FALSE
  )
