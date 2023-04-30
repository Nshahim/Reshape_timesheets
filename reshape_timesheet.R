# Note: check march timesheet
library(tidyverse)
library(readxl)
`%notin%` <- Negate(`%in%`)

# Join Date with each hour
join_date_hour <- function(timesheet_data, date_rows){
  for(col_i in 2:(length(timesheet_data))){
    # Loop through each column of daily hours
    column <- timesheet_data[,col_i]
    
    curr_date <- NA
    for(row_i in 1:nrow(column)){
      value <- column[[row_i,1]]
      curr_date <- ifelse(row_i %in% date_rows, value, curr_date)
      
      if(!is.na(curr_date) & !is.na(value)){
        timesheet_data[row_i,col_i] <- paste0(value, "-", curr_date)
        # Replace any extra - in holidays
        if(grepl("-", value)){
          timesheet_data[row_i,col_i] <- str_replace(value, " - |-", "_")
        }
      }
    }
  }
  return(timesheet_data)
}

reshape_timesheet <- function(timesheet, file_name){
  ## Remove extra rows
  null_rows <- which(apply(timesheet,1,function(x)all(is.na(x))))
  timesheet <- timesheet[-null_rows,]
  
  ## Time sheet Details
  emp_name <- timesheet[[1,2]] # manual value
  timesheet_date <- as.Date(as.numeric(timesheet[[2,2]]), origin = "1899-12-30") %>% format(format="%h-%Y") # manual value
  th_row <- which(timesheet[[1]] %in% "Total Hours")
  total_hours <- timesheet[[th_row+1,1]]
  total_days <- timesheet[[th_row+1,2]]
  
  ## Rename columns
  cols <- c("Project_name","Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
  ts_sub <- timesheet[-c(1:2,th_row,th_row+1),-length(timesheet)]
  names(ts_sub) <- cols
  
  # Join Date with each hour
  weekday_rows <- which(ts_sub[[2]] %in% "Sunday")
  date_rows <- weekday_rows-1
  ts_sub <- join_date_hour(timesheet_data=ts_sub, date_rows)
  # Remove weekdays
  ts_sub <- ts_sub[-c(weekday_rows, date_rows),]
  
  # Pivot data
  ts_sub_reshaped <- ts_sub %>% 
    pivot_longer(-Project_name, names_to = "Weekdays", values_to = "Hours") %>% 
    separate_wider_delim(Hours, names=c("Hours", "Date"), delim="-", too_few="align_start") %>% 
    mutate(Day_Type = case_when(
      grepl("sick|holiday", Hours, ignore.case = T) ~ Hours,
      TRUE ~ "Regular"
    ), Hours = case_when(
      grepl("sick|holiday", Hours, ignore.case = T) ~ "0",
      TRUE ~ Hours
    ), .before=Hours) %>% 
    mutate(Employee_name = emp_name, Timesheet_date=timesheet_date, .before = Project_name) %>% 
    mutate(Date=as.Date(as.numeric(Date), origin = "1899-12-30"), Hours = as.numeric(Hours))
  
  # Extra Timesheet Daily Totals
  timesheet_totals <- ts_sub_reshaped %>% 
    filter(Project_name == "Total" & Hours!=0) %>% 
    select(Date, timesheet_hours=Hours)
  # Filter unnecessary rows
  ts_sub_reshaped <- ts_sub_reshaped %>% 
    filter(!grepl("Project Name|Total|Add if needed", Project_name) & Hours != 0) %>% 
    arrange(Project_name)
  
  ## Data Quality Checks
  Hours_match <- ts_sub_reshaped %>% 
    group_by(Date) %>% summarize(calculated_hours=sum(Hours)) %>% ungroup() %>% 
    left_join(timesheet_totals, "Date") %>% mutate(is_equal= calculated_hours==timesheet_hours) 
  
  # Total calculate hours and days equals to Total in timesheet
  if(!all(Hours_match$is_equal)){
    message("Daily timesheet total does not match the calculated daily total in: ", file_name)
    Hours_match[which(!Hours_match$is_equal),]
  }
  if(sum(ts_sub_reshaped$Hours) != total_hours | sum(ts_sub_reshaped$Hours)/8 != total_days){
    message("Total Hours/Days calculated does not match in: ", file_name)
  }
  
  # Return reshaped data
  return(ts_sub_reshaped)
}

timesheet <- read_excel("input/2022 Timesheets/Copy of Timesheet of March 2022(66)_approved.xlsx")
reshape_timesheet(timesheet=timesheet, file_name="df")

# Timesheets ---------------------------------------------------------------------------------------
path <- "input/2022 Timesheets/"
files <- list.files(path)
# files <- files[3]

merged_timesheets <- data.frame()
for(file in files){
  timesheet <- read_excel(paste0(path,file))
  
  ts_sub_reshaped <- reshape_timesheet(timesheet=timesheet, file_name=file)
  # Merge
  merged_timesheets <- rbind(merged_timesheets, ts_sub_reshaped)
}

merged_timesheets <- merged_timesheets %>% 
  mutate(Project_name = case_when(
    Project_name %in% c("ATR Education", "ATR-Education", "ATR-Eduation", "ATR-Eduaction", "ATR-Eduacation") ~ "ATR-Education",
    Project_name == "UNICEF P" ~ "UNICEF TV - 252",
    Project_name == "GovIEA" ~ "Govt Indicators",
    TRUE ~ Project_name
  ), Employee_name = case_when(
    Employee_name == "Deeba" ~ "Deeba Ezedyar",
    TRUE ~ Employee_name
  )) %>% 
  arrange(Date)

# tests
merged_timesheets %>% 
  filter(weekdays(Date) != Weekdays)

# Report
report <- merged_timesheets %>% 
  group_by(Employee_name, Timesheet_date, Project_name) %>% 
  summarize(Total_hours = sum(Hours),
            Total_days = total_hours/8) %>% 
  ungroup() %>% 
  arrange(Timesheet_date) %>% 
  mutate(Daily_rate = 145,
         Hourly_rate = Daily_rate/8,
         Salary_USD = Total_hours*Hourly_rate,
         Timesheet_date = format(Timesheet_date, "%h-%Y")) 

merged_timesheets <- merged_timesheets %>% 
  filter(Hours != 0) %>% 
  mutate(Timesheet_date = format(Timesheet_date, "%h-%Y")) 

export_list <- list(
  Merged_timesheet=merged_timesheets,
  Report=report
)

writexl::write_xlsx(export_list, paste0("output/Merged_timesheets_sample_", lubridate::today(),".xlsx"))
writexl::write_xlsx(report, "test.xlsx")
