#This file would take raw data dumps from Nielsen's API and create a weekly "Sports Scorecard" for all Scripps geographies to measure sports programming performance.
#The final output would be an R Markdown HTML file that would be uploaded to a website for viewing by internal executives.
#This code goes through the entire process of taking raw data and cleaning, structuring, analyzing and presenting all in one script.
#This code referred to local databases containing join tables of local teams and past season schedules to ensure apples to apples trending figures.

library(readxl)
library(openxlsx)
library(readr)
library(dplyr)
library(tidyr)
library(lubridate)
library(stringr)
library(ggalt)
library(ggplot2)
library(scales)
library(formattable)
library(reactable)
library(htmltools)
library(rmarkdown)
library(knitr)

do_it_sports <- function(xxxx) {
  
  begin_import_nltv(1)
  overnight_cycle_nltv(1)
  
}

begin_import_nltv <- function(xxxx) {
  
  sports_database <<- read_csv('Sports Database.csv')
  demo_run_vector <<- c("Persons 25-54")
  
  a <- 1
  g <- nrow(sports_database)
  
  for (a in a:g) {
    
    b <- as.character(sports_database[a, 4])
    if (b == "HH") { sports_database[a, 4] <<- "Households" }
    if (b == "A25_54") { sports_database[a, 4] <<- "Persons 25-54" }
    if (b == "A18_49") { sports_database[a, 4] <<- "Persons 18-49" }
    if (b == "M25_54") { sports_database[a, 4] <<- "Males 25-54" }
    if (b == "W25_54") { sports_database[a, 4] <<- "Females 25-54" }
    
  }
  
  sports_database <<- sports_database %>%
    filter(STATION != 'KDVR 31.1')
  sports_database$EPISODE <<- sub("RDSKN", "WA FB", sports_database$EPISODE)
  
  sports_run <<- read_excel('Sports Run.xlsx')
  
  local_database <<- read_excel('Local Sports.xlsx')
  episode_split(1)
  local_database_create(1)
  non_database_avgs(1)
  local_database_avgs(1)
  overall_widget_ct <<- 0
  a_ovn_count <<- 0
  
}

overnight_cycle_nltv <- function(xxxx) {
  
  sports_run_do <<- which(sports_run$Run > 0)
  sports_run_do <<- sports_run[sports_run_do, ]
  run_group <<- sports_run_do %>%
    group_by(GroupName) %>%
    summarize(ct = n())
  run_group <<- as.character(run_group[1, 1])
  
  a <- 1
  g <- nrow(sports_run_do)
  #g <- 1
  a_file_check <<- ""
  
  for (a in a:g) {
    
    key_prog <<- as.character(sports_run_do[a, 2])
    nltv_file <<- as.character(sports_run_do[a, 6])
    full_name <<- as.character(sports_run_do[a, 4])
    lastyear_name <<- as.character(sports_run_do[a, 5])
    lastyear_date <<- sports_run_do$LastYear_Date[a]
    lastyear_compmonth <<- as.character(sports_run_do[a, 8])
    lastyear_compyear <<- as.numeric(sports_run_do[a, 9])
    active_date <<- as.character(sports_run_do[a, 11])
    sub_title <<- paste(full_name, "vs.", lastyear_fullname, sep = " ")
    print(paste("Running:", key_prog, sep = " "))
    if (a_file_check != nltv_file) {
      print(paste("NLTV Load:", nltv_file, sep = " "))
      nltv_load(1)
      a_file_check <<- nltv_file
    }
    print(paste("A FILE CHECK = ", a_file_check, "A FILE = ", nltv_file, sep = " "))
    print(paste("Total Rank Calculation:", key_prog, sep = " "))
    total_rank_nltv_calcu(1)
    print(paste("Daypart Rank Calculation:", key_prog, sep = " "))
    daypart_rank_nltv_calcu(1)
    print(paste("Hour Rank Calculation:", key_prog, sep = " "))
    hour_rank_nltv_calcu(1)
    print(paste("Local Database Trending, Flagging and Pulls:", key_prog, sep = " "))
    
    local_flag(1)
    local_database_pull_nltv(1)
    trend_combine(1)
    
    #print(paste("Arianna and Database Combine:", key_prog, sep = " "))
    #a_and_l_combine_nltv(1)
    print(paste("HTML Widget Generate:", key_prog, sep = " "))
    b <- 1
    h <- length(demo_run_vector)
    
    for (b in b:h) {
      
      curr_demo <<- demo_run_vector[b]
      sports_scorecard(curr_demo)
      tables_generate(1)
      
    }
    
    if (a == 1) { 
      nltv_master_list <<- a_ovn_mutate1[, c(1, 4, 5, 10, 13, 17, 19, 21)]
      total_hour_rank_frame <<- master_hour_rank_frame
    }
    else { 
      nltv_master_list <<- rbind(nltv_master_list, a_ovn_mutate1[, c(1, 4, 5, 10, 13, 17, 19, 21)])
      total_hour_rank_frame <<- rbind(total_hour_rank_frame, master_hour_rank_frame)
    }
  }
  print("Final Markdown Execute")
  markdown_execute(1)
  print("Complete")
}

#Create Database of Local Airings Only

episode_split <- function(xxxx) {
  
  split_mx <<- as.matrix(sports_database$EPISODE)
  
  mx_1 <<- split_mx
  mx_2 <<- split_mx
  
  a <- 1
  g <- nrow(split_mx)
  
  for (a in a:g) {
    
    act_string <<- split_mx[a, 1]
    act_split <<- strsplit(act_string, split = "&")
    away_team <<- act_split[[1]][1]
    home_team <<- act_split[[1]][2]
    
    mx_1[a, 1] <<- away_team
    mx_2[a, 1] <<- home_team
    
  }
  sports_database[, 11] <<- mx_1
  sports_database[, 12] <<- mx_2
  colnames(sports_database)[8:9] <<- c("RATING", "IMPRESSIONS")
  colnames(sports_database)[11:12] <<- c("AWAY", "HOME")
}

local_database_create <<- function(xxxx) {
  
  sports_run_tmp <<- sports_run %>%
    filter(Run > 0)
  
  ct <- 0
  ct2 <- 0
  a <- 1
  g <- nrow(sports_run_tmp)
  
  for (a in a:g) {
    
    act_telecast <<- as.character(sports_run_tmp[a, 2])
    
    database_small <<- sports_database %>%
      filter(PROGRAM_NAME == act_telecast)
    
    if (nrow(database_small) > 0) {
      
      b <- 1
      h <- nrow(local_database)
        
      for (b in b:h) {
        
        act_geog <<- as.character(local_database[b, 1])
        act_team <<- as.character(local_database[b, 2])
        
        away_half <<- which(database_small$AWAY == act_team & database_small$GEOGRAPHY == act_geog)
        if (length(away_half) > 0) {
          ct <- ct + 1
          away_frame <<- database_small[away_half, ]
          away_frame[, 13] <<- act_team
          if (ct == 1) { local_frame <<- away_frame }
          else { local_frame <<- rbind(local_frame, away_frame) }
        }
        
        home_half <<- which(database_small$HOME == act_team & database_small$GEOGRAPHY == act_geog)
        if (length(home_half) > 0) {
          ct <- ct + 1
          home_frame <<- database_small[home_half, ]
          home_frame[, 13] <<- act_team
          if (ct == 1) { local_frame <<- home_frame }
          else { local_frame <<- rbind(local_frame, home_frame) }
        }
        
        non_half1 <<- which(database_small$AWAY != act_team & database_small$HOME != act_team)
        if (length(non_half1) > 0) {
          non_half11 <<- database_small[non_half1, ]
          non_half2 <<- which(non_half11$GEOGRAPHY == act_geog)
          if (length(non_half2) > 0) { 
            non_half22 <<- non_half11[non_half2, ]
            non_frame <<- non_half22
            non_frame[, 13] <<- act_team
            ct2 <- ct2 + 1
            if (ct2 == 1) { 
              nonlocal_frame <<- non_frame
            }
            else { 
              nonlocal_frame <<- rbind(nonlocal_frame, non_frame) 
            }
          }
        }
      }
    }
  }
  colnames(local_frame)[13] <<- "TEAM"
  colnames(nonlocal_frame)[13] <<- "TEAM"
}

local_database_avgs <- function(xxxx) {
  
  a <- 1
  g <- nrow(local_database)
  ct <- 0
  
  for (a in a:g) {
   
    act_geog <<- as.character(local_database[a, 1]) 
    act_team <<- as.character(local_database[a, 2]) 
    
    ld_frame <<- local_frame %>%
      filter(GEOGRAPHY == act_geog, TEAM == act_team) %>%
      group_by(GEOGRAPHY, DEMO, TEAM) %>%
      summarize(RATING = mean(RATING), IMPRESSIONS = mean(IMPRESSIONS))
    
    if (nrow(ld_frame) > 0) {
      ct <- ct + 1
      if (ct == 1) { local_team_avg_frame <<- ld_frame }
      else { local_team_avg_frame <<- rbind(local_team_avg_frame, ld_frame) }

    }
  }
}

non_database_avgs <- function(xxxx) {
  
  a <- 1
  g <- nrow(local_database)
  ct <- 0
  
  for (a in a:g) {
    
    act_geog <<- as.character(local_database[a, 1]) 
    act_team <<- as.character(local_database[a, 2]) 
    
    ld_frame <<- nonlocal_frame %>%
      filter(GEOGRAPHY == act_geog, TEAM == act_team) %>%
      group_by(GEOGRAPHY, PROGRAM_NAME, DEMO, VA_NIELSEN_MONTH_NAME, VA_NIELSEN_YEAR_NUMBER) %>%
      summarize(RATING = mean(RATING), IMPRESSIONS = mean(IMPRESSIONS))
    
    if (nrow(ld_frame) > 0) {
      ct <- ct + 1
      if (ct == 1) { local_non_avg_frame <<- ld_frame }
      else { local_non_avg_frame <<- rbind(local_non_avg_frame, ld_frame) }
      
    }
  }
  local_non_avg_frame <<- local_non_avg_frame %>%
    group_by(GEOGRAPHY, PROGRAM_NAME, DEMO, VA_NIELSEN_MONTH_NAME, VA_NIELSEN_YEAR_NUMBER) %>%
    summarize(RATING = mean(RATING), IMPRESSIONS = mean(IMPRESSIONS))
}


#Plot List Export Simply Combines different graph exports into a Single PDF
#Useful for different airings of the same sport(I.E. NFL Sundays)

plot_list_export <- function(xxxx) {
  
  pdf(pdf_output)
  for (i in 1:length(plot_list)) {
    print(plot_list[[i]])
  }
  dev.off()
  
}

nltv_load <- function(xxxx) {
  
  a_file <<- paste("Total Day Ranker Template", nltv_file, sep = " ")
  a_file <<- paste(a_file, ".xlsx", sep = "")
  a_ovn <<- read_excel(a_file, sheet = "Live+Same Day, Custom Range 1, ")
  
  a_ovn_len <- nrow(a_ovn)
  a_ovn_len_short <- a_ovn_len - 8
  
  a_ovn <<- a_ovn[11:a_ovn_len_short, ]
  a_ovn <<- a_ovn[, c(1:4, 6, 10:13)]
  
  #Change 0 Ratings and Imps to 0.01 and 10
  
  a_ovn_rtg_mx <<- as.matrix(a_ovn[, 6])
  a_ovn_rtg_mx <<- matrix(sapply(a_ovn_rtg_mx, as.numeric))
  a <- 1
  g <- nrow(a_ovn_rtg_mx)
  for (a in a:g) {
    
    b <- as.numeric(a_ovn_rtg_mx[a, 1])
    if (b <= 0) { 
      a_ovn_rtg_mx[a, 1] <<- 0.01
    }
  }
  
  a_ovn_imp_mx <<- as.matrix(a_ovn[, 7])
  a_ovn_imp_mx <<- matrix(sapply(a_ovn_imp_mx, as.numeric))
  a <- 1
  g <- nrow(a_ovn_imp_mx)
  for (a in a:g) {
    
    b <- as.numeric(a_ovn_imp_mx[a, 1])
    if (b <= 0) { 
      a_ovn_imp_mx[a, 1] <<- 10
    }
  }
  
  #Loop to separate Call from Station
  
  a_ovn_mx <<- as.matrix(a_ovn[, 2])
  a_ovn_new_mx <<- a_ovn_mx
  
  a <- 1
  g <- nrow(a_ovn_mx)
  
  for (a in a:g) {
    
    b <- as.character(a_ovn_mx[a, 1])
    c <- strsplit(b, split = " ")[[1]]
    f <- c[[1]]
    h <- strsplit(f, split = "-")[[1]]
    k <- h[[1]]
    a_ovn_new_mx[a, 1] <<- k
    
  }
  
  #Loop to Rename Demos
  
  a_ovn_demo_mx <<- as.matrix(a_ovn[ ,5])
  a <- 1
  g <- nrow(a_ovn_demo_mx)
  
  for (a in a:g) {
    
    b <- as.character(a_ovn_demo_mx[a, 1])
    if (b == "HH") { k <- "Households" }
    if (b == "P25-54") { k <- "Persons 25-54" }
    if (b == "P18-49") { k <- "Persons 18-49" }
    if (b == "M25-54") { k <- "Males 25-54" }
    if (b == "F25-54") { k <- "Females 25-54" }
    
    a_ovn_demo_mx[a, 1] <<- k
    
  }
  
  #Loop to Create Timeslots
  
  a_ovn_stime_mx <<- as.matrix(a_ovn[, 8])
  a_ovn_stime_mx <<- matrix(sapply(a_ovn_stime_mx, as.numeric))
  a_ovn_stime_mx <<- a_ovn_stime_mx + .000001
  a_ovn_stime_hour <<- a_ovn_stime_mx * 24
  a_ovn_stime_hour_int <<- as.integer(a_ovn_stime_hour)
  a_ovn_stime_min <<- (a_ovn_stime_hour - a_ovn_stime_hour_int) * 60
  a_ovn_stime_mx_final <<- cbind(a_ovn_stime_hour_int, as.integer(a_ovn_stime_min))
  
  a_ovn_etime_mx <<- as.matrix(a_ovn[, 9])
  a_ovn_etime_mx <<- matrix(sapply(a_ovn_etime_mx, as.numeric))
  a_ovn_etime_mx <<- a_ovn_etime_mx + .000001
  a_ovn_etime_hour <<- a_ovn_etime_mx * 24
  a_ovn_etime_hour_int <<- as.integer(a_ovn_etime_hour)
  a_ovn_etime_min <<- (a_ovn_etime_hour - a_ovn_etime_hour_int) * 60
  a_ovn_etime_mx_final <<- cbind(a_ovn_etime_hour_int, as.integer(a_ovn_etime_min))
  
  #Create Final Dataframe
  
  a_ovn_almost <<- data.frame(a_ovn[, 1], a_ovn_new_mx, a_ovn[, c(2:4)], a_ovn_demo_mx, a_ovn_stime_mx_final, a_ovn_etime_mx_final, a_ovn_rtg_mx, a_ovn_imp_mx)
  
  colnames(a_ovn_almost) <<- c("Geography", "Call", "Station", "Program", "Episode", "Demo", "n_s_hr", "n_s_mm", "n_e_hr", "n_e_mm", "r_rtg", "r_imp")
  a_ovn_new_filt <<- a_ovn_almost %>%
    filter(Call != "WWSB")
  
  #-----------JOIN WITH ARIANNA WING SCRIPT-----------------
  
  
  #Timezone List
  
  tz_list <<- a_ovn_new_filt %>% group_by(Geography) %>% summarize(ct = n())
  a <- 1
  g <- nrow(tz_list)
  
  for (a in a:g) {
    
    act_geog <- tz_list[a, 1]
    if (act_geog == "Baltimore") { tz_list[a, 2] <<- 1 }
    if (act_geog == "Cincinnati") { tz_list[a, 2] <<- 1 }
    if (act_geog == "Cleveland-Akron (Canton)") { tz_list[a, 2] <<- 1 }
    if (act_geog == "Denver") { tz_list[a, 2] <<- 2 }
    if (act_geog == "Detroit") { tz_list[a, 2] <<- 1 }
    if (act_geog == "Indianapolis") { tz_list[a, 2] <<- 1 }
    if (act_geog == "Kansas City") { tz_list[a, 2] <<- 2 }
    if (act_geog == "Las Vegas") { tz_list[a, 2] <<- 1 }
    if (act_geog == "Miami-Ft. Lauderdale") { tz_list[a, 2] <<- 1 }
    if (act_geog == "Milwaukee") { tz_list[a, 2] <<- 2 }
    if (act_geog == "Nashville") { tz_list[a, 2] <<- 2 }
    if (act_geog == "New York") { tz_list[a, 2] <<- 1 }
    if (act_geog == "Norfolk-Portsmth-Newpt Nws") { tz_list[a, 2] <<- 1 }
    if (act_geog == "Phoenix (Prescott)") { tz_list[a, 2] <<- 2 }
    if (act_geog == "Salt Lake City") { tz_list[a, 2] <<- 2 }
    if (act_geog == "San Diego") { tz_list[a, 2] <<- 1 }
    if (act_geog == "Tampa-St. Pete (Sarasota)") { tz_list[a, 2] <<- 1 }
    if (act_geog == "West Palm Beach-Ft. Pierce") { tz_list[a, 2] <<- 1 }
    
  }
  
  #Add Column for Daypart, Scripps ID
  
  a_ovn_new_filt[, 13:14] <<- 0
  
  #Loop to Determine Daypart and Scripps Flag
  
  num_matrix <<- as.matrix(a_ovn_new_filt[, 7:14])
  char_matrix <<- as.matrix(a_ovn_new_filt[, 1:6])
  
  a <- 1
  g <- nrow(char_matrix)
  
  for (a in a:g) {
    
    act_geog <- as.character(char_matrix[a, 1])
    act_net <- as.character(char_matrix[a, 2])
    act_tz <- which(tz_list$Geography == act_geog)
    act_tz <- as.numeric(tz_list[act_tz, 2])
    s_flag <- 0
    
    if (act_net == "WMAR") { s_flag <- 1 }
    if (act_net == "WCPO") { s_flag <- 1 }
    if (act_net == "WEWS") { s_flag <- 1 }
    if (act_net == "KMGH") { s_flag <- 1 }
    if (act_net == "WXYZ") { s_flag <- 1 }
    if (act_net == "WRTV") { s_flag <- 1 }
    if (act_net == "KSHB") { s_flag <- 1 }
    if (act_net == "KTNV") { s_flag <- 1 }
    if (act_net == "WSFL") { s_flag <- 1 }
    if (act_net == "WTMJ") { s_flag <- 1 }
    if (act_net == "WTVF") { s_flag <- 1 }
    if (act_net == "WPIX") { s_flag <- 1 }
    if (act_net == "WTKR") { s_flag <- 1 }
    if (act_net == "KNXV") { s_flag <- 1 }
    if (act_net == "KSTU") { s_flag <- 1 }
    if (act_net == "KGTV") { s_flag <- 1 }
    if (act_net == "WFTS") { s_flag <- 1 }
    if (act_net == "WPTV") { s_flag <- 1 }
    
    s_hour <- as.numeric(num_matrix[a, 1])
    s_min <- as.numeric(num_matrix[a, 2])
    e_hour <- as.numeric(num_matrix[a, 3])
    e_min <- as.numeric(num_matrix[a, 4])
    
    s_min_pct <- s_min / 60
    e_min_pct <- e_min / 60
    
    s_comb <- s_hour + s_min_pct
    e_comb <- e_hour + e_min_pct
    
    slot_avg <- mean(c(s_comb, e_comb))
    
    if (s_hour > e_hour) { dp_use <- 10 }
    else {
      
      if (slot_avg >= 4 & slot_avg < 7) { dp_use <- 1 }
      if (slot_avg >= 7 & slot_avg < 9) { dp_use <- 2 }
      if (slot_avg >= 9 & slot_avg < 11) { dp_use <- 3 }
      if (slot_avg >= 11 & slot_avg < 13) { dp_use <- 4 }
      if (slot_avg >= 13 & slot_avg < 16) { dp_use <- 5 }
      
      if (act_tz == 1) {
        
        if (slot_avg >= 16 & slot_avg < 19) { dp_use <- 6 }
        if (slot_avg >= 19 & slot_avg < 20) { dp_use <- 7 }
        if (slot_avg >= 20 & slot_avg < 23) { dp_use <- 8 }
        if (slot_avg >= 23 & slot_avg < 23.55) { dp_use <- 9 }
        if (slot_avg >= 23.55 | slot_avg < 4) { dp_use <- 10 }
        
      }
      
      else {
        
        if (slot_avg >= 16 & slot_avg < 18) { dp_use <- 6 }
        if (slot_avg >= 18 & slot_avg < 19) { dp_use <- 7 }
        if (slot_avg >= 19 & slot_avg < 22) { dp_use <- 8 }
        if (slot_avg >= 22 & slot_avg < 22.55) { dp_use <- 9 }
        if (slot_avg > 22.55 | slot_avg < 4) { dp_use <- 10 }
        
      }
    }
    
    num_matrix[a, 7] <<- dp_use
    num_matrix[a, 8] <<- s_flag
  }
  
  #Create Hour Matrix
  
  hour_matrix <<- data.frame(char_matrix[, c(1:6)], num_matrix[, 1:2])
  hour_matrix[, 9:32] <<- 0
  colnames(hour_matrix) <<- c("Geography", "Call", "Station", "Program", "Episode", "Demo", "SHour", "SMin", "H0", "H1", "H2", "H3", "H4", "H5", "H6", "H7",
                              "H8", "H9", "H10", "H11", "H12", "H13", "H14", "H15", "H16", "H17", "H18", "H19", "H20", "H21", "H22", "H23")
  
  a <- 1
  g <- nrow(hour_matrix)
  
  for (a in a:g) {
    
    s_hour <- as.numeric(num_matrix[a, 1])
    e_end <- as.numeric(num_matrix[a, 4])
    e_hour <- as.numeric(num_matrix[a, 3])
    if (e_end == 0) { e_hour = e_hour - 1 }
    if (e_hour < s_hour) { e_hour <- 23 }
    
    for (s_hour in s_hour:e_hour) {
      
      hr_slot <<- s_hour + 9
      hour_matrix[a, hr_slot] <<- 1
      
    }
  }
  
  
  #Recombine and add 3 columns for Rank
  
  a_ovn_recomb <<- data.frame(char_matrix)
  a_ovn_recomb <<- data.frame(lapply(a_ovn_recomb, as.character), stringsAsFactors=FALSE)
  a_ovn_recomb <<- data.frame(a_ovn_recomb, num_matrix)
  a_ovn_recomb[, 15:17] <<- 0
  geog_vector <<- tz_list[, 1]
  geog_vector <<- data.frame(lapply(geog_vector, as.character), stringsAsFactors=FALSE)
  geog_vector <<- c(geog_vector[, 1])
  
  
  colnames(a_ovn_recomb)[7:17] <<- c("SHour", "SMin", "EHour", "EMin", "Rtg", "Imp", "Daypart", "Scripps", "TotRank", "DPRank", "HRRank")
  
}


#Total Rank Calculation

total_rank_nltv_calcu <- function(xxxx) {
  
  demo_vector <- c("Households", "Persons 18-49", "Persons 25-54", "Females 25-54", "Males 25-54")
  dp_vector <- c(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
  
  a <- 1
  g <- length(geog_vector)
  ct <- 0
  
  for (a in a:g) {
    
    act_geog <- geog_vector[a]
    
    b <- 1
    h <- length(demo_vector)
    
    for (b in b:h) {
      
      act_demo <- demo_vector[b]
      
      tr_frame <- a_ovn_recomb %>%
        filter(Geography == act_geog, Demo == act_demo) %>%
        arrange(desc(Imp))
      
      if (nrow(tr_frame) > 0) { 
        
        ct <- ct + 1
        tr_frame_length <- nrow(tr_frame)
        tr_rank_vector <- c(1:tr_frame_length)
        
        tr_frame[, 15] <- tr_rank_vector
        
        if (ct == 1) { tr_frame_with_dp_rank <<- tr_frame }
        else { tr_frame_with_dp_rank <<- rbind(tr_frame_with_dp_rank, tr_frame) }
        
      }
    }
  }
  
  a_ovn_recomb <<- tr_frame_with_dp_rank
}

#Daypart Rank Calculation

daypart_rank_nltv_calcu <- function(xxxx) { 
  demo_vector <- c("Households", "Persons 18-49", "Persons 25-54", "Females 25-54", "Males 25-54")
  dp_vector <- c(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
  
  a <- 1
  g <- length(geog_vector)
  ct <- 0
  
  for (a in a:g) {
    
    act_geog <- geog_vector[a]
    
    b <- 1
    h <- length(demo_vector)
    
    for (b in b:h) {
      
      act_demo <- demo_vector[b]
      
      c <- 1
      i <- length(dp_vector)
      
      for (c in c:i) {
        
        act_daypart <- dp_vector[c]
        
        tr_frame <- a_ovn_recomb %>%
          filter(Geography == act_geog, Demo == act_demo, Daypart == act_daypart) %>%
          arrange(desc(Imp))
        
        if (nrow(tr_frame) > 0) { 
          
          ct <- ct + 1
          tr_frame_length <- nrow(tr_frame)
          tr_rank_vector <- c(1:tr_frame_length)
          
          tr_frame[, 16] <- tr_rank_vector
          
          if (ct == 1) { tr_frame_with_dp_rank <<- tr_frame }
          else { tr_frame_with_dp_rank <<- rbind(tr_frame_with_dp_rank, tr_frame) }
          
        }
      }
    }
  }
  
  a_ovn_recomb <<- tr_frame_with_dp_rank
}


#Hour Rank Calculation

hour_rank_nltv_calcu <- function(xxxx) { 
  a_ovn_recomb_short <<- a_ovn_recomb[, c(1:8, 12)]
  
  hr_vector <- c(9:32)
  a <- 1
  g <- length(hr_vector)
  ct <- 0
  
  for (a in a:g) {
    
    act_hour <<- hr_vector[a]
    
    b <- 1
    h <- length(geog_vector)
    
    for (b in b:h) {
      
      act_geog <<- geog_vector[b]
      
      c <- 1
      i <- length(demo_vector)
      
      for (c in c:i) {
        
        act_demo <- demo_vector[c]
        hour_subset <<- which(hour_matrix[, 1] == act_geog & hour_matrix[, act_hour] == 1 & hour_matrix[, 6] == act_demo)
        hour_frame <<- hour_matrix[hour_subset, 1:8]
        
        if (nrow(hour_frame) > 0) {
          
          ct <- ct + 1
          
          hour_frame_join <<- hour_frame %>%
            left_join(a_ovn_recomb_short, by = c("Geography", "Call", "Station", "Program", "Episode", "Demo", "SHour", "SMin")) %>%
            arrange(desc(Imp))
          
          actual_hour <- act_hour - 9
          recomb_frame <- a_ovn_recomb %>%
            filter(Geography == act_geog, Demo == act_demo, SHour == actual_hour)
          
          if (ct == 1) {
            ultra_hour_rank_frame_tmp <<- hour_frame_join %>%
              left_join(a_ovn_recomb, by = c("Geography", "Call", "Station", "Program", "Episode", "Demo", "SHour", "SMin", "Imp"))
            ultra_hour_rank_frame_tmp[, 18] <<- actual_hour
            ultra_hour_rank_frame <<- ultra_hour_rank_frame_tmp
          }
          
          else {
            ultra_hour_rank_frame_tmp <<- hour_frame_join %>%
              left_join(a_ovn_recomb, by = c("Geography", "Call", "Station", "Program", "Episode", "Demo", "SHour", "SMin", "Imp"))
            ultra_hour_rank_frame_tmp[, 18] <<- actual_hour
            ultra_hour_rank_frame <<- rbind(ultra_hour_rank_frame, ultra_hour_rank_frame_tmp)
          }
          
          d <- 1
          j <- nrow(recomb_frame)
          
          if (j > 0) {
            for (d in d:j) {
              act_prog <<- as.character(recomb_frame[d, 4])
              act_prog_shour <<- as.numeric(recomb_frame[d, 7])
              act_prog_smin <<- as.numeric(recomb_frame[d, 8])
              act_prog_call <<- as.character(recomb_frame[d, 2])
              hour_frame_join_line <<- which(hour_frame_join$Program == act_prog & hour_frame_join$SHour == act_prog_shour & hour_frame_join$SMin == act_prog_smin
                                            & hour_frame_join$Call == act_prog_call)
              #print(paste(act_prog, act_prog_shour, act_prog_smin, act_prog_call, act_geog, act_hour))
              recomb_frame[d, 17] <- hour_frame_join_line
              
            }
            
            if (ct == 1) { 
              tr_frame_with_hr_rank <<- recomb_frame 
            }
            else { 
              tr_frame_with_hr_rank <<- rbind(tr_frame_with_hr_rank, recomb_frame) 
            }
          }
        }
      }
    }
  }
  
  tr_frame_with_hr_rank$Geography <- toupper(tr_frame_with_hr_rank$Geography)
  a_ovn_final <<- tr_frame_with_hr_rank

}

local_flag <- function(xxxx) {
  
  a_ovn_final[, 18:21] <<- 0 
  episode_mx <<- a_ovn_final[, 5]
  a <- 1
  g <- length(episode_mx)
  
  for (a in a:g) {
    s <- as.character(episode_mx[a])
    amp_check <- grepl("&", s)
    if (amp_check == TRUE) {
      split_tmp <- strsplit(s, split = "&")
      away_tmp <- split_tmp[[1]][1]
      home_tmp <- split_tmp[[1]][2]
      a_ovn_final[a, 18] <<- away_tmp
      a_ovn_final[a, 19] <<- home_tmp
      geog_tmp <- a_ovn_final[a, 1]
      geog_grid <- local_database %>% 
        filter(Geography == geog_tmp)

      if (nrow(geog_grid) > 0) {
        b <- 1
        h <- nrow(geog_grid)
        
        for (b in b:h) {
          check_team <- as.character(geog_grid[b, 2])
          if (check_team == away_tmp) { 
            a_ovn_final[a, 20] <<- away_tmp 
            a_ovn_final[a, 21] <<- 1
          }
          if (check_team == home_tmp) {
            a_ovn_final[a, 20] <<- home_tmp 
            a_ovn_final[a, 21] <<- 1
          }
        }
      }
    }
  }
  colnames(a_ovn_final)[18:21] <<- c("AWAY", "HOME", "TEAM", "LOCAL")
}

local_database_pull_nltv <- function(xxxx) {
  
  local_flag_frame <<- a_ovn_final %>%
    filter(LOCAL > 0)
  
  local_ly_frame <<- local_flag_frame %>%
    left_join(local_team_avg_frame, by = c("Geography" = "GEOGRAPHY", "Demo" = "DEMO", "TEAM"))
  
  non_ly_filter <<- which(local_non_avg_frame$PROGRAM_NAME == lastyear_name & 
                          local_non_avg_frame$VA_NIELSEN_MONTH_NAME == lastyear_compmonth &
                          local_non_avg_frame$VA_NIELSEN_YEAR_NUMBER == lastyear_compyear)
  
  non_ly_frame <<- local_non_avg_frame[non_ly_filter, ]

}

#Formats the previous airing database into a nice table to be joined into the Arianna Table

trend_combine <- function(xxxx) {
  
  a_ovn_final_short <<- a_ovn_final %>%
    filter(Program == key_prog)
  
  a <- 1
  g <- nrow(a_ovn_final_short)
  
  for (a in a:g) {
    geog_tmp <- as.character(a_ovn_final_short[a, 1])
    demo_tmp <- as.character(a_ovn_final_short[a, 6])
    local_tmp <- as.character(a_ovn_final_short[a, 20])
    flag_tmp <- as.character(a_ovn_final_short[a, 21])
    
    
    if (flag_tmp > 0) {
      
      line_tmp <- which(local_team_avg_frame$GEOGRAPHY == geog_tmp &
                        local_team_avg_frame$DEMO == demo_tmp &
                        local_team_avg_frame$TEAM == local_tmp)
      
      if (length(line_tmp) > 0) {
        
        tmp_rtg <- local_team_avg_frame$RATING[line_tmp]
        tmp_imp <- local_team_avg_frame$IMPRESSIONS[line_tmp]
        tmp_desc <- paste(lastyear_compyear, local_tmp, sep = " ")
         
      }
      else {
        
        line_tmp <- which(local_non_avg_frame$GEOGRAPHY == geog_tmp &
                          local_non_avg_frame$DEMO == demo_tmp &
                          local_non_avg_frame$PROGRAM_NAME == key_prog &
                          local_non_avg_frame$VA_NIELSEN_MONTH_NAME == lastyear_compmonth &
                          local_non_avg_frame$VA_NIELSEN_YEAR_NUMBER == lastyear_compyear)
        
        tmp_rtg <- local_non_avg_frame$RATING[line_tmp]
        tmp_imp <- local_non_avg_frame$IMPRESSIONS[line_tmp]
        tmp_desc <- paste(lastyear_compmonth, lastyear_compyear, key_prog, sep = " ")
        
      }
    }
    
    else {
       
      line_tmp <- which(local_non_avg_frame$GEOGRAPHY == geog_tmp &
                        local_non_avg_frame$DEMO == demo_tmp &
                        local_non_avg_frame$PROGRAM_NAME == key_prog &
                        local_non_avg_frame$VA_NIELSEN_MONTH_NAME == lastyear_compmonth &
                        local_non_avg_frame$VA_NIELSEN_YEAR_NUMBER == lastyear_compyear)
      
      tmp_rtg <- local_non_avg_frame$RATING[line_tmp]
      tmp_imp <- local_non_avg_frame$IMPRESSIONS[line_tmp]
      tmp_desc <- paste(lastyear_compmonth, lastyear_compyear, key_prog, sep = " ")
      
    }
    
    #print(paste(geog_tmp, demo_tmp, local_tmp, flag_tmp, key_prog, sep = " "))
    if (length(line_tmp) == 0) { 
      tmp_desc = NA
      tmp_rtg = NA
      tmp_imp = NA
    }
    a_ovn_final_short[a, 22] <<- tmp_desc
    a_ovn_final_short[a, 23] <<- tmp_rtg
    a_ovn_final_short[a, 24] <<- tmp_imp
  }
  colnames(a_ovn_final_short)[22:24] <<- c("COMP_DESC", "LYRTG", "LYIMP")
  
  a_ovn_mutate1 <<- a_ovn_final_short %>%
    mutate(RtgChg = Rtg - LYRTG) %>%
    mutate(ImpChg = Imp - LYIMP) %>%
    mutate(RtgPctChg = (Rtg - LYRTG) / LYRTG) %>%
    mutate(ImpPctChg = (Imp - LYIMP) / LYIMP)
  
  a <- 1
  g <- nrow(a_ovn_mutate1)
  
  ultra_hour_rank_frame$Geography <<- toupper(ultra_hour_rank_frame$Geography)
  
  for (a in a:g) {
    
    act_geog <- as.character(a_ovn_mutate1[a, 1])
    act_demo <- as.character(a_ovn_mutate1[a, 6])
    act_shour <- as.numeric(a_ovn_mutate1[a, 7])
    act_dp <- as.numeric(a_ovn_mutate1[a, 13])
    
    tmp_hour <- ultra_hour_rank_frame %>%
      filter(Geography == act_geog, Demo == act_demo, V18 == act_shour)
    
    tmp_dp <- a_ovn_final %>%
      filter(Geography == act_geog, Demo == act_demo, Daypart == act_dp)
    
    if (a == 1) { 
      master_hour_rank_frame <<- tmp_hour
      master_dp_rank_frame <<- tmp_dp
    }
    
    else { 
      master_hour_rank_frame <<- rbind(master_hour_rank_frame, tmp_hour)
      master_dp_rank_frame <<- rbind(master_dp_rank_frame, tmp_dp)
    }
  }
  if (a_ovn_count == 0) { a_ovn_bind <<- a_ovn_mutate1 }
  else { a_ovn_bind <<- rbind(a_ovn_bind, a_ovn_mutate1) }
  a_ovn_count <<- a_ovn_count + 1
}

#Joins Arianna and Database tables together

a_and_l_combine_nltv <- function(xxxx) {
  
  a_ovn_combine <<- a_ovn_final_short %>%
    left_join(l_ovn_trim2, by = c("Geography", "Demo"))
  
  a_ovn_combine2 <<- a_ovn_combine[, c(-17, -18, -19)]
  
  a_ovn_mutate1 <<- a_ovn_combine2 %>%
    mutate(RtgChg = Rtg.x - Rtg.y) %>%
    mutate(ImpChg = Imp.x - Imp.y) %>%
    mutate(RtgPctChg = (Rtg.x - Rtg.y) / Rtg.y) %>%
    mutate(ImpPctChg = (Imp.x - Imp.y) / Imp.y)
  
  colnames(a_ovn_mutate1)[10:11] <<- c("Rtg", "Imp")
  colnames(a_ovn_mutate1)[17:18] <<- c("PrevRtg", "PrevImp")
  
  a <- 1
  g <- nrow(a_ovn_mutate1)
  
  for (a in a:g) {
    
    act_geog <- as.character(a_ovn_mutate1[a, 1])
    act_demo <- as.character(a_ovn_mutate1[a, 5])
    act_shour <- as.numeric(a_ovn_mutate1[a, 6])
    act_dp <- as.numeric(a_ovn_mutate1[a, 12])
    
    tmp_hour <- ultra_hour_rank_frame %>%
      filter(Geography == act_geog, Demo == act_demo, V17 == act_shour)
    
    tmp_dp <- a_ovn_final %>%
      filter(Geography == act_geog, Demo == act_demo, Daypart == act_dp)
    
    if (a == 1) { 
      master_hour_rank_frame <<- tmp_hour
      master_dp_rank_frame <<- tmp_dp
    }
    
    else { 
      master_hour_rank_frame <<- rbind(master_hour_rank_frame, tmp_hour)
      master_dp_rank_frame <<- rbind(master_dp_rank_frame, tmp_dp)
    }
  }
  
  a_ovn_database <<- a_ovn_mutate1 %>% 
    group_by(Geography, Demo) %>%
    summarize(ct = n())
  

  
}


sports_scorecard <- function(demo_feed) {
  
  sports_sc <<- a_ovn_mutate1 %>%
    filter(Program == key_prog, Demo == demo_feed) %>%
    select(Geography, Program, Episode, SHour, SMin, Rtg, Scripps, LOCAL, COMP_DESC, 
           LYRTG, LYIMP, RtgPctChg) %>%
    arrange(desc(RtgPctChg))
  
  meaner_table2 <<- which(is.na(sports_sc$LYRTG) == FALSE)
  meaner_table <<- sports_sc[meaner_table2, ]
  
  avg_pctchg <<- (avg_thisyear - avg_lastyear) / avg_lastyear
  avg_pctchg <<- scales::percent(avg_pctchg, accuracy=0.1)
  
  avg_thisyear <<- mean(meaner_table$Rtg)
  avg_lastyear <<- mean(meaner_table$LYRTG)
  
  avg_thisyear <<- as.numeric(format(round(avg_thisyear, digits = 1), nsmall = 2))
  avg_lastyear <<- as.numeric(format(round(avg_lastyear, digits = 1), nsmall = 2))
  
  
  a <- 1
  g <- nrow(sports_sc)
  
  for (a in a:g) {
    
    hhh <- as.numeric(sports_sc[a, 4])
    mmm <- as.numeric(sports_sc[a, 5])
    fff <- time_convert(hhh, mmm)
    sports_sc[a, 13] <- fff
    
  }
  
  if (demo_feed == "Persons 25-54") { demo_bar <<- "A2554 Rating" }
  if (demo_feed == "Men 25-54") { demo_bar <<- "M2554 Rating" }
  if (demo_feed == "Households") { demo_bar <<- "Household Rating" }
  if (demo_feed == "Women 25-54") { demo_bar <<- "W2554 Rating" }
  if (demo_feed == "Persons 18-49") { demo_bar <<- "A1849 Rating" }
  
  colnames(sports_sc)[13] <- "StartTime"
  sports_sc$RtgPctChg <<- scales::percent(sports_sc$RtgPctChg, accuracy=0.1)
  sports_sc$Rtg <<- format(round(sports_sc$Rtg, digits = 1), nsmall = 1)
  sports_sc$LYRTG <<- format(round(sports_sc$LYRTG, digits = 1), nsmall = 1)
  
  scripps_sc <<- sports_sc %>%
    filter(Scripps == 1) %>%
    select(Geography, StartTime, Episode, Rtg, LYRTG, RtgPctChg, COMP_DESC)
  
  if (nrow(scripps_sc) > 0) { 
    scripps_sc$RtgPctChg <<- scales::percent(scripps_sc$RtgPctChg, accuracy=0.1)
    scripps_sc$Rtg <<- format(round(scripps_sc$Rtg, digits = 1), nsmall = 1)
    scripps_sc$LYRTG <<- format(round(scripps_sc$LYRTG, digits = 1), nsmall = 1)
    scripps_sc[, 8] <<- scripps_sc$RtgPctChg
    scripps_sc[, 9]  <<- "Scripps"
    colnames(scripps_sc)[8:9] <<- c("z", "Class")
    scripps_sc <<- scripps_sc %>% select(Class, Geography, StartTime, Episode, Rtg, LYRTG, z, RtgPctChg, COMP_DESC)
    colnames(scripps_sc) <<- c("Class", "Geography", "Start Time", "Matchup", demo_bar, "Comp Rating", " ", "% Change", "Comparison")
  } 
  
  local_sc <<- sports_sc %>%
    filter(Scripps == 0 & LOCAL == 1) %>%
    select(Geography, StartTime, Episode, Rtg, LYRTG, RtgPctChg, COMP_DESC)
  
  if (nrow(local_sc) > 0) {
    local_sc$RtgPctChg <<- scales::percent(local_sc$RtgPctChg, accuracy=0.1)
    local_sc$Rtg <<- format(round(local_sc$Rtg, digits = 1), nsmall = 1)
    local_sc$LYRTG <<- format(round(local_sc$LYRTG, digits = 1), nsmall = 1)
    local_sc[, 8] <<- local_sc$RtgPctChg
    local_sc[, 9]  <<- "Local"
    colnames(local_sc)[8:9] <<- c("z", "Class")
    local_sc <<- local_sc %>% select(Class, Geography, StartTime, Episode, Rtg, LYRTG, z, RtgPctChg, COMP_DESC)
    colnames(local_sc) <<- c("Class", "Geography", "Start Time", "Matchup", demo_bar, "Comp Rating", " ", "% Change", "Comparison")
  }
  
  other_sc <<- sports_sc %>%
    filter(Scripps == 0 & LOCAL == 0) %>%
    select(Geography, StartTime, Episode, Rtg, LYRTG, RtgPctChg, COMP_DESC)
  
  if (nrow(other_sc) > 0) { 
    other_sc$RtgPctChg <<- scales::percent(other_sc$RtgPctChg, accuracy=0.1)
    other_sc$Rtg <<- format(round(other_sc$Rtg, digits = 1), nsmall = 1)
    other_sc$LYRTG <<- format(round(other_sc$LYRTG, digits = 1), nsmall = 1)
    other_sc[, 8] <<- other_sc$RtgPctChg
    other_sc[, 9] <<- "Other"
    colnames(other_sc)[8:9] <<- c("z", "Class")
    other_sc <<- other_sc %>% select(Class, Geography, StartTime, Episode, Rtg, LYRTG, z, RtgPctChg, COMP_DESC)
    colnames(other_sc) <<- c("Class", "Geography", "Start Time", "Matchup", demo_bar, "Comp Rating", " ", "% Change", "Comparison")
  }
  
  master_table <<- rbind(scripps_sc, local_sc, other_sc)
  avg_row <<- nrow(master_table) + 1
  master_table[avg_row, 1] <<- "Average"
  master_table[avg_row, 2] <<- key_prog
  master_table[avg_row, 3] <<- " "
  master_table[avg_row, 4] <<- " "
  master_table[avg_row, 5] <<- avg_thisyear
  master_table[avg_row, 6] <<- avg_lastyear
  master_table[avg_row, 7:8] <<- avg_pctchg
  master_table[avg_row, 9] <<- " "
  
  master_table_form <<- formattable(
    
    master_table,
                                    
    align = c("l", "l", rep("c", NCOL(master_table) - 2)),
    list(
    `Class` = formatter("span", style = x ~ style(
                        color = ifelse(x == "Scripps", "darkblue", ifelse(x == "Local", "sienna", 
                                ifelse(x == "Other", "grey", "black"))),
                        font.weight=ifelse(x == "Scripps" | x == "Average", "bold", "italic"))),
    
    `Geography` = formatter("span", style = ~ style(color = ifelse(Class == "Scripps", "darkblue", 
                            ifelse(Class == "Local", "sienna", ifelse(Class == "Other", "grey", "black"))),                 
                            font.weight = ifelse(Class == "Scripps" | Class == "Average", "bold", "italic"))),
                                      
    `Start Time` = formatter("span", style = ~ style(color = ifelse(Class == "Scripps", "darkblue", 
                            ifelse(Class == "Local", "sienna", ifelse(Class == "Other", "grey", "black"))),                 
                            font.weight = ifelse(Class == "Scripps" | Class == "Average", "bold", "italic"))),
    
                                      
    `Matchup` = formatter("span", style = ~ style(color = ifelse(Class == "Scripps", "darkblue", 
                             ifelse(Class == "Local", "sienna", ifelse(Class == "Other", "grey", "black"))),                 
                             font.weight = ifelse(Class == "Scripps" | Class == "Average", "bold", "italic"))),
    
     ` ` = formatter("span", style = x ~ style(color = ifelse(x > 0, "green", "red")),
                     x ~ icontext(ifelse(x > 0, "arrow-up", "arrow-down"), ifelse(x > 0, "Up", "Down"))),
                                      
    `% Change` = formatter("span", style = ~ style(color = ifelse(Class == "Scripps", "darkblue", 
                             ifelse(Class == "Local", "sienna", ifelse(Class == "Other", "grey", "black"))),                 
                             font.weight = ifelse(Class == "Scripps" | Class == "Average", "bold", "italic"))),
                   
    
    `Comparison` = formatter("span", style = ~ style(color = ifelse(Class == "Scripps", "darkblue", 
                             ifelse(Class == "Local", "sienna", ifelse(Class == "Other", "grey", "black"))),                 
                             font.weight = ifelse(Class == "Scripps" | Class == "Average", "bold", "italic"))),
    
    
     area(col = c(demo_bar)) ~ color_tile("#fafeff", "#008fb3"),
     area(col = c(`Comp Rating`)) ~ color_tile("#fcfafa", "#615d5d")
     
    ))
}

tables_generate <- function(xxxx) {
  
  #Convert Frame to HTML
  ggg <<- as.htmlwidget(master_table_form, width = "80%")
  
  #Add HTML Tags to Widgets
  
  if (overall_widget_ct == 0) { 
    bbb <<- tags$div(tags$h1(class = "text-center", run_group),
                     tags$h1(class = "text-center", "."),
                     tags$h2(class = "text-center", key_prog),
                     tags$h3(class = "text-center" , demo_bar),
                     tags$div(align = "center", class = "table", ggg))
    browsable_list <<- list(bbb)
    overall_widget_ct <<- length(browsable_list)
  } 
  else { 
    bbb <<- tags$div(tags$h1(class = "text-center", "."),
                     tags$h2(class = "text-center", key_prog),
                     tags$h3(class = "text-center" , demo_bar),
                     tags$div(align = "center", class = "table", ggg))
    overall_widget_ct <<- overall_widget_ct + 1
    browsable_list[[overall_widget_ct]] <<- bbb
  }
}

markdown_execute <- function(xxxx) {
  
  #Put it All Together
  lester <<- browsable(tagList(browsable_list))
  output_name <<- paste(run_group, ".html", sep = "")
  save.image(file='NLTVWorkspace.RData')
  render("NLTV Sports Markdown.Rmd", "all", output_file = output_name)
  
}

local_story_extract <- function(xxxx) {
  
  scripps_frame <<- a_ovn_bind %>%
    filter(Scripps == 1)
  
  geog_vectors <<- scripps_frame %>%
    group_by(Geography) %>%
    summarize(ct = n())
  
  program_vectors <<- scripps_frame %>%
    group_by(Program) %>%
    summarize(ct = n())
  
  geog_vectors <<- as.matrix(geog_vectors[, 1])
  program_vectors <<- as.matrix(program_vectors[, 1])
  demo_vector <<- c("Households", "Persons 18-49", "Persons 25-54", "Females 25-54", "Males 25-54")
  
  a <- 1
  g <- length(geog_vectors)
  
  for (a in a:g) {
     
    act_geogbind <<- as.character(geog_vectors[a])
    
    good_stories_ct <<- 0
    
    b <- 1
    h <- length(program_vectors)
    
    for (b in b:h) {
      
      act_prog <<- program_vectors[b]
      
      print(paste("Story Generator: ", act_geogbind, act_prog, sep = " "))
      
      geog_group <<- scripps_frame %>%
        filter(Geography == act_geogbind, Program == act_prog)
      
      if (nrow(geog_group) > 0) { 
        db_group <<- geog_group
        
        act_epis <<- as.character(geog_group[1, 5])
        act_comp <<- as.character(geog_group[1, 22])
        act_dp <<- as.character(geog_group[1, 13])
        
        lollipop_overview(1)
        lollipop_pos_plot(1)
        
        c <- 1
        i <- length(demo_vector)
        
        for (c in c:i) {
         
          d_u <- demo_vector[c]
          
          rank_check <<- a_ovn_final %>%
            filter(Demo == d_u, Geography == act_geogbind, Program == act_prog) 
          
          t_rank_check <<- rank_check$TotRank
          h_rank_check <<- rank_check$HRRank
          hr_use <<- rank_check$SHour
          sc_check <<- rank_check$Scripps
   
          if (sc_check > 0) { 
            if (t_rank_check == 1) { 
              single_trank_plot(d_u) 
            }
            if (t_rank_check > 1 & h_rank_check == 1) { 
              single_hrank_plot(d_u, hr_use) 
            }
          }
        }
      }
    }
    
    if (good_stories_ct > 0) { 
      
      print(paste(good_stories_ct, act_geogbind, sep = " "))
      pdf_output <<- paste(act_geogbind, "Sports Success Stories", active_date, sep = " - ")
      pdf_output <<- paste(pdf_output, ".pdf", sep = "")
      
      pdf(pdf_output)
      for (i in 1:length(good_stories_list)) {
        print(good_stories_list[[i]])
      }
      dev.off()
        
    }
  }
}

single_trank_plot <- function(demo_use) {
  
  s_group <<- a_ovn_final %>%
    filter(Demo == demo_use, Geography == act_geogbind) %>%
    arrange(desc(Imp))
  
  no1_prog <<- s_group$Program[1]
  
  if (no1_prog == act_prog) {
    
    chart_title <- paste(act_prog, "#1 in", demo_use, "among all Broadcasts", sep = " ")
    chart_subtitle <- paste("Top 10 Total Day", demo_use, "Programs - ", active_date, sep = " ")
    
    u <- nrow(s_group)
    if (u > 10) { u <- 10 }
    s_group2 <<- s_group[1:u, ]
    ggplot_max <- max(s_group2$Imp)
    
    s_group2 <<- s_group2[order(s_group2$Imp), ]
    s_group2$Program <<- factor(s_group2$Program, levels=as.character(s_group2$Program))
    
    s_trank_plot <<- ggplot(s_group2, aes(x=Program, y=Imp, label=Imp, fill=Scripps)) +
      geom_bar(stat='identity') +
      geom_text(color="black", size=3.5, hjust=1.3, aes(label = scales::comma(Imp))) +
      geom_text(color="darkblue", aes(y=-100, label=TotRank), hjust=2, vjust=-2, size=2) +
      scale_fill_gradient(
        low = "gray",
        high = "deepskyblue"
      ) +
      ylim(-100, (ggplot_max * 1.01)) +
      theme(legend.position = "none") +
      ylab("Viewers") +
      xlab(" ") +
      labs(title=chart_title,
           subtitle=chart_subtitle) + 
      coord_flip()
    
    good_stories_ct <<- good_stories_ct + 1
    if (good_stories_ct == 1) { good_stories_list <<- list(s_trank_plot) }
    else { good_stories_list[[good_stories_ct]] <<- s_trank_plot }
  }
}

single_hrank_plot <- function(demo_use, hour_check) {
  
  s_group <<- total_hour_rank_frame %>%
    filter(Demo == demo_use, Geography == act_geogbind, V18 == hour_check) %>%
    group_by(Geography, Call, Station, Program, Episode, Demo, SHour,
             Daypart, Scripps, HRRank) %>%
    summarize(ct = n(), Rtg = mean(Rtg), Imp = mean(Imp)) %>%
    arrange(desc(Imp))
  
  no1_prog <<- s_group$Program[1]
  
  if (no1_prog == act_prog) {
    
    if (hour_check > 12) { ts_use <- paste(hour_check - 12, "pm", sep = "") }
    if (hour_check == 12) { ts_use <- paste(12, "pm", sep = "") }
    if (hour_check == 0) { ts_use <- paste(12, "am", sep = "") }
    if (hour_check < 12 & hour_check > 0) { ts_use <- paste(hour_check, "am", sep = "") }
    
    chart_title <- paste(act_prog, "#1 in", demo_use, "during Timeslot", sep = " ")
    chart_subtitle <- paste("Top", ts_use, "Programs -", demo_use, "-", active_date, sep = " ")
    
    u <- nrow(s_group)
    if (u > 10) { u <- 10 }
    s_group2 <<- s_group[1:u, ]
    ggplot_max <- max(s_group2$Imp)
    hr_slot <- c(u:1)
    
    s_group2 <<- s_group2[order(s_group2$Imp), ]
    s_group2$Program <<- factor(s_group2$Program, levels=as.character(s_group2$Program))
    s_group2$HRRank <<- hr_slot
    
    s_hrank_plot <<- ggplot(s_group2, aes(x=Program, y=Imp, label=Imp, fill=Scripps)) +
      geom_bar(stat='identity') +
      geom_text(color="black", size=3.5, hjust=1.1, aes(label = scales::comma(Imp))) +
      geom_text(color="darkblue", aes(y=-100, label=HRRank), hjust=2, vjust=-2, size=2) +
      scale_fill_gradient(
        low = "gray",
        high = "deepskyblue"
      ) +
      ylim(-100, (ggplot_max * 1.01)) +
      theme(legend.position = "none") +
      ylab("Viewers") +
      xlab(" ") +
      labs(title=chart_title,
           subtitle=chart_subtitle) + 
      coord_flip()
    
    good_stories_ct <<- good_stories_ct + 1
    if (good_stories_ct == 1) { good_stories_list <<- list(s_hrank_plot) }
    else { good_stories_list[[good_stories_ct]] <<- s_hrank_plot }
  }
}

lollipop_pos_plot <- function(xxxx) {
  
  chart_title <<- paste(active_date, act_prog, "Grows in Key Demos", sep = " ")
  chart_subtitle <<- paste("Viewership Increase Compared to", act_comp, "Broadcasts", sep = " ")
  
  theme_set(theme_bw())
  
  db_group1 <<- db_group[order(db_group$ImpPctChg), ]
  db_group1$Demo <<- factor(db_group1$Demo, levels=as.character(db_group1$Demo))
  db_group1 <<- db_group1 %>%
    filter(ImpPctChg > .074)
  
  if (nrow(db_group1) > 0) {
   
    ggplot_max <- max(db_group1$ImpPctChg)
    
    pos_lollipop_plot <<- ggplot(db_group1, aes(x=Demo, y=ImpPctChg, label=ImpPctChg)) + 
      geom_bar(stat='identity', fill="green")  +
      geom_segment(aes(y = 0, 
                       x = Demo, 
                       yend = ImpPctChg, 
                       xend = Demo),
                       color="darkgreen") + 
      geom_text(color="darkgreen", size=3.5, hjust=-0.5, aes(label = scales::percent(ImpPctChg, accuracy = 0.1))) +
      geom_text(aes(y=-0.15, label=scales::comma(Imp)), 
                color="darkgreen", size=4, hjust=0, vjust = -1) +
      geom_text(aes(y=-0.15, label=scales::comma(LYIMP)), 
                color="dimgray", size=4, hjust=0, vjust = 1) +
      labs(title=chart_title,
           subtitle=chart_subtitle) + 
      theme(axis.title.x = element_blank(), axis.title.y = element_blank()) +
      ylim(-0.2, (ggplot_max * 1.4)) +
      coord_flip()
    
    good_stories_ct <<- good_stories_ct + 1
    
    if (good_stories_ct == 1) { good_stories_list <<- list(pos_lollipop_plot) }
    else { good_stories_list[[good_stories_ct]] <<- pos_lollipop_plot }
 
  } 
}

lollipop_overview <- function(xxxx) {
  
  chart_title <<- paste(act_prog, active_date, act_epis, sep = " - ")
  chart_subtitle <<- paste("Viewers By Demo Compared to", act_comp, "Broadcasts", sep = " ")
  
  theme_set(theme_bw())
  ggplot_max <- max(geog_group$ImpPctChg)
  if (ggplot_max < .1) { ggplot_max <- .1 }

  ggplot_min <- min(geog_group$ImpPctChg)
  if (abs(ggplot_min) < ggplot_max) { ggplot_min <- ggplot_max * -1 }
  
  geog_group <<- geog_group %>%
    mutate(Type = ifelse(ImpPctChg > 0, "Increase", "Decrease"))
  
  geog_group1 <- geog_group[order(geog_group$ImpPctChg), ]
  geog_group1$Demo <- factor(geog_group1$Demo, levels = geog_group1$Demo)
  
  overview_lollipop_plot <<- ggplot(geog_group1, aes(x=Demo, y=ImpPctChg, label=ImpPctChg)) + 
    geom_bar(stat='identity', aes(fill=Type))  +
    scale_fill_manual(values = c("Decrease"="tomato", "Increase"="green")) + 
    geom_text(color="black", size=3.5, hjust=-0.25, aes(label = scales::percent(ImpPctChg, accuracy = 0.1))) +
    geom_text(aes(y=(ggplot_min * 1.35), label=scales::comma(Imp)), 
              color="darkblue", size=4, hjust=0, vjust = 0.5) +
    labs(title=chart_title,
         subtitle=chart_subtitle) + 
    theme(axis.title.x = element_blank(), axis.title.y = element_blank()) +
    ylim((ggplot_min * 1.4) , (ggplot_max * 1.2)) +
    coord_flip()
  
  geog_group1 <- geog_group[order(desc(geog_group$TotRank)), ]
  geog_group1$Demo <- factor(geog_group1$Demo, levels = geog_group1$Demo)
  
  totrank_ct <- a_ovn_final %>%
    filter(Geography == act_geogbind, Demo == "Households")
  totrank_ct <- nrow(totrank_ct)
  chart_title <- paste(act_prog, active_date, act_epis, sep = " - ")
  chart_subtitle <- paste("Viewership/Market Rank Among", totrank_ct, "Aired Programs For the Day", sep = " ")
  
  overview_trank_plot <<- ggplot(geog_group1, aes(x=Demo, y=TotRank, label=TotRank, col=Demo)) + 
    geom_point(size=8) +   
    geom_text(color="white", size=4) +
    geom_segment(aes(x=Demo, 
                     xend=Demo, 
                     y=(0 - (totrank_ct * .15)), 
                     yend=totrank_ct), 
                 linetype="dashed", 
                 size=0.1) +
    geom_text(aes(y=(0 - (totrank_ct * .15)), label=scales::comma(Imp)), 
              color="darkblue", size=4, hjust=0, vjust = 0.5) +
    labs(title=chart_title, 
         subtitle=chart_subtitle, 
         caption="source: Nielsen NLTV") +  
    xlab(" ") +
    ylab("Total Day Market Rank") +
    theme(legend.position = "none") +
    coord_flip()
  
  
  
  geog_group1 <- geog_group[order(desc(geog_group$DPRank)), ]
  geog_group1$Demo <- factor(geog_group1$Demo, levels = geog_group1$Demo)
  
  drank_ct <- a_ovn_final %>%
    filter(Geography == act_geogbind, Demo == "Households", Daypart == act_dp)
  drank_ct <- nrow(drank_ct)
  chart_title <- paste(act_prog, active_date, act_epis, sep = " - ")
  chart_subtitle <- paste("Viewership/Daypart Rank Among", drank_ct, "Aired Programs in the Daypart", sep = " ")
  
  overview_drank_plot <<- ggplot(geog_group1, aes(x=Demo, y=DPRank, label=DPRank, col=Demo)) + 
    geom_point(size=8) + 
    geom_text(color="white", size=4) +
    geom_segment(aes(x=Demo, 
                     xend=Demo, 
                     y=(0 - (drank_ct * .15)), 
                     yend=drank_ct), 
                 linetype="dashed", 
                 size=0.1) +
    geom_text(aes(y=(0 - (drank_ct * .15)), label=scales::comma(Imp)), 
              color="darkblue", size=4, hjust=0, vjust = 0.5) +
    labs(title=chart_title, 
         subtitle=chart_subtitle, 
         caption="source: Nielsen NLTV") +  
    xlab(" ") +
    ylab("Daypart Rank") +
    theme(legend.position = "none") +
    coord_flip()
  
  
  geog_group1 <<- geog_group[order(desc(geog_group$HRRank)), ]
  geog_group1$Demo <<- factor(geog_group1$Demo, levels = geog_group1$Demo)
  hr_group_use <<- geog_group1$SHour[1]
  
  hrank_ct <<- total_hour_rank_frame %>%
    filter(Geography == act_geogbind, Demo == "Households", V18 == hr_group_use)
  hrank_ct <<- nrow(hrank_ct)
  chart_title <- paste(act_prog, active_date, act_epis, sep = " - ")
  chart_subtitle <- paste("Viewership/Timeslot Rank Among", hrank_ct, "Aired Programs in the Timeslot", sep = " ")
  
  overview_hrank_plot <<- ggplot(geog_group1, aes(x=Demo, y=HRRank, label=HRRank, col=Demo)) + 
    geom_point(size=8) + 
    geom_text(color="white", size=4) +
    geom_segment(aes(x=Demo, 
                     xend=Demo, 
                     y=(0 - (hrank_ct * .15)), 
                     yend=hrank_ct), 
                 linetype="dashed", 
                 size=0.1) +
    geom_text(aes(y=(0 - (hrank_ct * .15)), label=scales::comma(Imp)), 
              color="darkblue", size=4, hjust=0, vjust = 0.5) +
    labs(title=chart_title, 
         subtitle=chart_subtitle, 
         caption="source: Nielsen NLTV") +  
    xlab(" ") +
    ylab("Timeslot Rank") +
    theme(legend.position = "none") +
    coord_flip() 
  
  pdf_output <<- paste(act_geogbind, act_prog, active_date, "Overview", sep = " ")
  pdf_output <<- paste(pdf_output, ".pdf", sep = "")
  plot_list <<- list(overview_lollipop_plot, overview_trank_plot, overview_hrank_plot)
  
  pdf(pdf_output)
  for (i in 1:length(plot_list)) {
    print(plot_list[[i]])
  }
  dev.off()
  
}

time_convert <- function(h, m) {

  if (h == 0) { 
    hh <- 12 
    pp <- "AM"
  }
  if (h > 0 & h < 12) { 
    hh <- h 
    pp <- "AM"
  }
  if (h == 12) {
    hh <- 12
    pp <- "PM"
  }
  if (h > 12) { 
    hh <- h - 12
    pp <- "PM"
  }
  hh <- as.character(hh)
  
  if (nchar(m) == 1) { mm <- as.character(paste(m, 0, sep = "")) }
  else { mm <- as.character(m) }
  
  f <- paste(hh, mm, sep = ":")
  f <- paste(f, pp, sep = " ")
  
  return(f)
  
}
