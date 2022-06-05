 #***********************************
# Reading and appending multiple sheets
# with identical data columns
# to one data frame
#***********************************

# Loading libraries
library(readxl)
library(dplyr)
library(tidyr)
library(lubridate)
library(xlsx)

# Set year-end for development calculation
yr_end <- 2021

# Set file path
wb_path <- "chap2_nonlife_reserving_r/data/201_example_data.xlsx"

# Read data
paid_data_raw <- read_excel(path = wb_path, sheet = "Paid")

# Filter for attrional loss only
paid_data_fil <- 
  paid_data_raw %>% 
  mutate(
    class_subclass = paste0(CLASS, "_", SUBCLASS),
    # class_subclass = paste0(CLASS),
    acc_yr = year(LOSSDATE),
    # dev_yr = year(PAYMENTDATE) - year(LOSSDATE)
    dev_yr = floor(as.numeric(PAYMENTDATE - LOSSDATE)/365)
  )
  
# Summarize paid by class, acc, dev 
paid_data_sum <- 
  paid_data_fil %>% 
  group_by(class_subclass, acc_yr, dev_yr) %>% 
  summarise(inc_paid = sum(GROSSPAID, na.rm = TRUE)) %>%
  ungroup()

min_acc_yr <- min(paid_data_sum$acc_yr)
max_acc_yr <- yr_end
max_dev_yr <- max_acc_yr-min_acc_yr

# Complete acc_yr and dev_yr for missing years
paid_data_Com <-  
  paid_data_sum %>% 
  group_by(class_subclass) %>% 
  complete(acc_yr = min_acc_yr:max_acc_yr, dev_yr = 0:max_dev_yr, fill = list(inc_paid = 0)) %>% 
  filter(dev_yr <= max_acc_yr - acc_yr) %>% 
  ungroup()

# Calculate cumulative paid
paid_data_cum <-  
  paid_data_Com %>% 
  arrange(class_subclass, acc_yr, dev_yr) %>% 
  group_by(class_subclass, acc_yr) %>% 
  mutate(cum_paid = cumsum(inc_paid)) %>% 
  ungroup()

# Calculate loss development factor
paid_data_ldf <-  
  paid_data_cum %>% 
  arrange(class_subclass, acc_yr, dev_yr) %>% 
  group_by(class_subclass, acc_yr) %>% 
  mutate(ldf = lead(cum_paid, default = last(cum_paid))/cum_paid) %>% 
  mutate(cdf_tmp = ldf*lead(ldf, default = 1)) %>% 
  mutate(cdf = cdf_tmp*lead(cdf_tmp, default = 1)) %>% 
  ungroup()

# Make ldf triangle
paid_data_tri_ldf <- 
  paid_data_ldf %>% 
  pivot_wider(
    id_cols = c(class_subclass, acc_yr),
    names_from = dev_yr,
    values_from = ldf
  )

# Make cdf triangle
paid_data_tri_cdf <- 
  paid_data_ldf %>% 
  pivot_wider(
    id_cols = c(class_subclass, acc_yr),
    names_from = dev_yr,
    values_from = cdf
  )

# Export ldf triangle
write.xlsx(
  x = paid_data_tri_ldf, 
  file = "chap2_nonlife_reserving_r/output/202_output_data.xlsx",
  sheetName = "motor_pd_ldf",
  # row.names = FALSE,
  showNA = FALSE
  )

# Export cdf triangle
write.xlsx(
  x = paid_data_tri_cdf, 
  file = "chap2_nonlife_reserving_r/output/202_output_data.xlsx",
  sheetName = "motor_pd_cdf",
  # row.names = FALSE,
  showNA = FALSE,
  append = TRUE
)

# Export long triangle
write.xlsx(
  x = paid_data_tri_cdf, 
  file = "chap2_nonlife_reserving_r/output/202_output_data.xlsx",
  sheetName = "motor_pd_long",
  # row.names = FALSE,
  showNA = FALSE,
  append = TRUE
)
