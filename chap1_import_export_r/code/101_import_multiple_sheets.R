#***********************************
# Reading and appending multiple sheets
# with identical data columns
# to one data frame
#***********************************

# Loading libraries
library(readxl)
library(xlsx)

# Set file path for the workbook with multiple sheet
wb_path <- "chap1_import_export_r/data/101_example_data.xlsx"

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

# Example 1: Write output as csv
write.csv(
  x = wb_data_df, 
  file = "chap1_import_export_r/output/101a_output_data.csv",
  row.names = FALSE
  )

# Example 2: Write output as excel individually
write.xlsx(x = as.data.frame(wb_data_df[wb_data_df$LOB == "Cargo", ]), file = "chap1_import_export_r/output/101b_output_data.xlsx", sheetName = "Cargo", row.names = FALSE, append = T)
write.xlsx(x = as.data.frame(wb_data_df[wb_data_df$LOB == "Fire", ]) , file = "chap1_import_export_r/output/101b_output_data.xlsx", sheetName = "Fire" , row.names = FALSE, append = T)

# Example 3: Write output as excel using lapply while setting path variable
# Set output path
out_path <- "chap1_import_export_r/output"

# Loop data and write one df per sheet
lapply(
    X = wb_sheets_fix,
    FUN = function(x){
      write.xlsx(x = as.data.frame(wb_data_df[wb_data_df$LOB == x, ]), file = paste0(out_path, "/101c_output_data_lapply.xlsx"), sheetName = x, row.names = FALSE, append = T)
    }
  )