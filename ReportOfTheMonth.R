#ReportOfTheMonth is an R script I use to parse Nielsen Data to deliver a rundown of programming that aired on the station.
#The report measures # of Telecasts, Viewers, Rating, Month to Month Viewership Change/Rating Change/% Change, Year over Year Viewer Change/Rating Change/% Change,
  #Daypart Rank in Viewership, Daypart Rank in Month over Month Change, Daypart Rank in Year over Year Change, the # Station and Program in the Timeslot, and
  #how far away the station was from being #1 if they were not #1.
#This R script took an an input file, "THEGRID", which is a cleaned, transformed and structured dataset containing all the metrics needed for the final output.
#The Final Output is an Excel Workbook that utilized VBA macros for further visualization enhancements.

library(readxl)
library(openxlsx)
library(readr)
library(dplyr)
library(tidyr)
library(lubridate)

grid_frame <- read_excel("THEGRID.xlsx")

a <- 18
g <- nrow(geog_frame)
g <- 18

wb <- createWorkbook()

for (a in a:g) {
  
  act_geog <- as.character(geog_frame[a, 1])
  
  b <- 1
  h <- length(demo_vector)
  #h <- 1
  
  for (b in b:h) {

    act_demo <- demo_vector[b]
    
    grid_filter_frame <- grid_frame %>% 
      filter(GEOGRAPHY == act_geog, DEMO == act_demo, DAYPART > 0) %>%
      arrange(START_TIME) %>%
      arrange(DAYPART)

    grid_sc_frame_morn <- grid_filter_frame %>%
      filter(SCRIPPS == 1, DAYPART <= 7)
    
    grid_sc_frame_m <- grid_filter_frame %>%
      filter(SCRIPPS == 1, DAYPART == 8, DAY_OF_WEEK == "M......")
    
    grid_sc_frame_t <- grid_filter_frame %>%
      filter(SCRIPPS == 1, DAYPART == 8, DAY_OF_WEEK == ".Tu.....")
    
    grid_sc_frame_w <- grid_filter_frame %>%
      filter(SCRIPPS == 1, DAYPART == 8, DAY_OF_WEEK == "..W....")
    
    grid_sc_frame_r <- grid_filter_frame %>%
      filter(SCRIPPS == 1, DAYPART == 8, DAY_OF_WEEK == "...Th...")
    
    grid_sc_frame_f <- grid_filter_frame %>%
      filter(SCRIPPS == 1, DAYPART == 8, DAY_OF_WEEK == "....F..")
    
    grid_sc_frame_u <- grid_filter_frame %>%
      filter(SCRIPPS == 1, DAYPART == 8, DAY_OF_WEEK == "......Su")
    
    grid_sc_frame_night1 <- grid_filter_frame %>%
      filter(SCRIPPS == 1, DAYPART == 9)
    
    grid_sc_frame_night2 <- grid_filter_frame %>%
      filter(SCRIPPS == 1, DAYPART > 9, START_TIME_HH == "23" | START_TIME_HH == "22")
    
    grid_sc_frame_night3 <- grid_filter_frame %>%
      filter(SCRIPPS == 1, DAYPART > 9, START_TIME_HH  == "00" | START_TIME_HH == "01" | START_TIME_HH == "02" | START_TIME_HH == "03")
    
    grid_sc_frame_all <- rbind(grid_sc_frame_morn, grid_sc_frame_m, grid_sc_frame_t, grid_sc_frame_w, grid_sc_frame_r, grid_sc_frame_f, grid_sc_frame_u, 
                               grid_sc_frame_night1, grid_sc_frame_night2, grid_sc_frame_night3)
    
    grid_sc_frame_all <- grid_sc_frame_all %>%
      select(DAY_OF_WEEK, DAYPART, START_TIME_HH, START_TIME_HFHR, START_TIME, PROGRAM_NAME, TELECAST, IMP_R1, RTG_R1, R1R2_IMP, R1R2_RTG, R1R2_PCT, R1R3_IMP, R1R3_RTG, R1R3_PCT, DRANK_VIEWERSHIP, DRANK_IMPCHG_MOM, DRANK_IMPCHG_YOY)
    
    c <- 1
    i <- nrow(grid_sc_frame_all)
    p <- ncol(grid_sc_frame_all)
    
    for (c in c:i) {
      
      act_hr = as.character(grid_sc_frame_all$START_TIME_HH[c])
      act_dp = as.character(grid_sc_frame_all$DAYPART[c])
      act_dp_num = as.numeric(act_dp)
      
      if (act_dp_num < 8) {
      
        finder_frame <- grid_filter_frame %>%
          filter(START_TIME_HH == act_hr) %>%
          arrange(desc(IMP_R1))
      }
      
      else if (act_dp_num > 8) {
        
        act_hh = as.character(grid_sc_frame_all$START_TIME_HFHR[c])
        finder_frame <- grid_filter_frame %>%
          filter(START_TIME_HH == act_hr, START_TIME_HFHR == act_hh) %>%
          arrange(desc(IMP_R1))
        
      }
      
      else {
        
        act_dow = as.character(grid_sc_frame_all$DAY_OF_WEEK[c])
        
        finder_frame <- grid_filter_frame %>%
          filter(START_TIME_HH == act_hr, DAY_OF_WEEK == act_dow) %>%
          arrange(desc(IMP_R1))
        
      }
      
      k <- p + 1
      o <- p + 2
      r <- p + 3
      s <- p + 4
      
      grid_sc_frame_all[c, k] <- as.character(finder_frame$STATION[1])
      grid_sc_frame_all[c, o] <- as.character(finder_frame$PROGRAM_NAME[1])
      grid_sc_frame_all[c, r] <- as.numeric(finder_frame$IMP_R1[1])
      
    }
    
    colnames(grid_sc_frame_all)[k] <- c("Station_1")
    colnames(grid_sc_frame_all)[o] <- c("Program_1")
    colnames(grid_sc_frame_all)[r] <- c("Viewers_1")
    
    grid_sc_frame_all <- grid_sc_frame_all %>%
      mutate(Difference = IMP_R1 - Viewers_1) %>%
      select(DAY_OF_WEEK, START_TIME, PROGRAM_NAME, TELECAST, IMP_R1, RTG_R1, R1R2_IMP, R1R2_RTG, 
             R1R2_PCT, R1R3_IMP, R1R3_RTG, R1R3_PCT, DRANK_VIEWERSHIP, DRANK_IMPCHG_MOM, DRANK_IMPCHG_YOY, Station_1,
             Program_1, Viewers_1, Difference)
    
    
    sheet_name = paste(substr(act_geog, 1, 12), act_demo, sep = "-")
    
    addWorksheet(wb, sheet_name)
    writeData(wb, sheet = sheet_name, x = grid_sc_frame_all)
    
  }
}

wb_name <- paste(act_geog, "ROTM", sep = " ")
wb_name <- paste(wb_name, ".xlsx", sep = "")
saveWorkbook(wb, wb_name, overwrite = TRUE)
