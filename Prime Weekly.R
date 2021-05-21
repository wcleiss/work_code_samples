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

#Import Snowflake Export File and Parse to Remove Uneeded Rows

begin_prime_sc <- function(xxxx) {

  l_file <<- "Prime Snowflake.csv"
  a_file <<- "Actual Prime Scorecard.xlsx"
  d_file <<- "Prime Run.xlsx"
  network_vector <<- c("ABC", "CBS", "FOX", "NBC")
  
  l_file <<- read_csv(l_file)
  a_file <<- read_excel(a_file, sheet = "Live+Same Day, Average")
  d_file <<- read_excel(d_file)
  
  a_ovn <<- a_file
  l_ovn <<- l_file
  d_ovn <<- d_file
  
  demo_run_vector <<- c("Persons 25-54")
  overall_widget_ct <<- 0
}

prime_sc_cycle <- function(xxxx) { 
  
}

prime_sc_afile_clean <- function(xxxx) {
  a_ovn_len <- nrow(a_ovn)
  a_ovn_len_short <- a_ovn_len - 6
  
  a_ovn <<- a_ovn[12:a_ovn_len_short, 1:11]
  a_ovn <<- a_ovn[, c(1:4, 7:11)]
  
  #Change 0 Ratings and Imps to 0.01 and 10
  
  a_ovn_rtg_mx <<- as.matrix(a_ovn[, 5])
  a_ovn_rtg_mx <<- matrix(sapply(a_ovn_rtg_mx, as.numeric))
  a <- 1
  g <- nrow(a_ovn_rtg_mx)
  for (a in a:g) {
    
    b <- as.numeric(a_ovn_rtg_mx[a, 1])
    if (b <= 0) { 
      a_ovn_rtg_mx[a, 1] <<- 0.01
    }
  }
  a_ovn_imp_mx <<- as.matrix(a_ovn[, 6])
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
  
  a_ovn_demo_mx <<- as.matrix(a_ovn[ ,4])
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
  
  a_ovn_stime_mx <<- as.matrix(a_ovn[, 7])
  a_ovn_stime_mx <<- matrix(sapply(a_ovn_stime_mx, as.numeric))
  a_ovn_stime_mx <<- a_ovn_stime_mx + .000001
  a_ovn_stime_hour <<- a_ovn_stime_mx * 24
  a_ovn_stime_hour_int <<- as.integer(a_ovn_stime_hour)
  a_ovn_stime_min <<- (a_ovn_stime_hour - a_ovn_stime_hour_int) * 60
  a_ovn_stime_mx_final <<- cbind(a_ovn_stime_hour_int, as.integer(a_ovn_stime_min))
  
  a_ovn_etime_mx <<- as.matrix(a_ovn[, 8])
  a_ovn_etime_mx <<- matrix(sapply(a_ovn_etime_mx, as.numeric))
  a_ovn_etime_mx <<- a_ovn_etime_mx + .000001
  a_ovn_etime_hour <<- a_ovn_etime_mx * 24
  a_ovn_etime_hour_int <<- as.integer(a_ovn_etime_hour)
  a_ovn_etime_min <<- (a_ovn_etime_hour - a_ovn_etime_hour_int) * 60
  a_ovn_etime_mx_final <<- cbind(a_ovn_etime_hour_int, as.integer(a_ovn_etime_min))

  #capitalize geographies
  
  a_ovn_geo_mx <<- as.matrix(a_ovn[, 1])
  a_ovn_geo_mx <<- toupper(a_ovn_geo_mx)
  
  #Create Final Dataframe
  
  a_ovn_almost <<- data.frame(a_ovn_geo_mx, a_ovn_new_mx, a_ovn[, c(2:3)], a_ovn_demo_mx, a_ovn_stime_mx_final, a_ovn_etime_mx_final, a_ovn_rtg_mx, a_ovn_imp_mx, a_ovn[, 9])
  
  colnames(a_ovn_almost) <<- c("Geography", "Call", "Station", "Program", "Demo", "n_s_hr", "n_s_mm", "n_e_hr", "n_e_mm", "r_rtg", "r_imp", "dow")
  a_ovn_new_filt <<- a_ovn_almost %>%
    filter(Call != "WWSB")
}


#----------LFILE MANIPULATION-----------

prime_sc_lfile_clean <- function(xxxx) {
  #Change 0 Ratings and Imps to 0.01 and 10
  
  a_ovn_rtg_mx <<- as.matrix(l_ovn[, 7])
  a_ovn_rtg_mx <<- matrix(sapply(a_ovn_rtg_mx, as.numeric))
  a <- 1
  g <- nrow(a_ovn_rtg_mx)
  for (a in a:g) {
    
    b <- as.numeric(a_ovn_rtg_mx[a, 1])
    if (b <= 0) { 
      a_ovn_rtg_mx[a, 1] <<- 0.01
    }
  }
  a_ovn_imp_mx <<- as.matrix(l_ovn[, 8])
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
  
  a_ovn_mx <<- as.matrix(l_ovn[, 2])
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
  
  a_ovn_demo_mx <<- as.matrix(l_ovn[ ,4])
  a <- 1
  g <- nrow(a_ovn_demo_mx)
  
  for (a in a:g) {
    
    b <- as.character(a_ovn_demo_mx[a, 1])
    if (b == "HH") { k <- "Households" }
    if (b == "A25_54") { k <- "Persons 25-54" }
    if (b == "A18_49") { k <- "Persons 18-49" }
    if (b == "M25_54") { k <- "Males 25-54" }
    if (b == "W25_54") { k <- "Females 25-54" }
    
    a_ovn_demo_mx[a, 1] <<- k
    
  }
  
  #Create Final Data Frame
  
  l_ovn_almost <<- data.frame(l_ovn[, 1], a_ovn_new_mx, l_ovn[, 3], a_ovn_demo_mx, l_ovn[, 5:6], a_ovn_rtg_mx, a_ovn_imp_mx)
  l_ovn_new_filt <<- l_ovn_almost
  colnames(l_ovn_new_filt) <<- c("Geography", "Call", "Program", "Demo", "Month", "Year", "Rating", "Impressions")
}

#------------------TRIM PRIME MONTHLY DATABSE------------

prime_sc_dfile_clean <- function(xxxx) {
  a <- 1
  d_ovn_short_mx <<- d_ovn[, 6:8]
  g <- nrow(d_ovn_short_mx)
  
  for (a in a:g) {
    
    prog <- as.character(d_ovn_short_mx[a, 1])
    cmonth <- as.character(d_ovn_short_mx[a, 2])
    cyear <- as.numeric(d_ovn_short_mx[a, 3])
    
    l_ovn_subset <<- which(l_ovn_new_filt$Program == prog & l_ovn_new_filt$Month == cmonth & l_ovn_new_filt$Year == cyear)
    
    
    if (a == 1) { l_ovn_stack <<- l_ovn_new_filt[l_ovn_subset, ] }
    else { l_ovn_stack <<- rbind(l_ovn_stack, l_ovn_new_filt[l_ovn_subset, ]) }
    
  }
}

#------------------MERGE WITH DATABSE--------------------

prime_sc_merge <- function(xxxx) {
  
  #Create 2020 Dataframe
  
  a_prog_mx <<- as.matrix(a_ovn_new_filt[, 4])
  
  a <- 1
  g <- nrow(d_ovn)
  ct <- 0
  
  for (a in a:g) {
    
    prog <- as.character(d_ovn[a, 5])
    if (is.na(prog) == FALSE) {
      
      a_match <- which(a_prog_mx[, 1] == prog)
      if (length(a_match) > 0) {
          
        ct <- ct + 1
        
        match_subset <<- a_ovn_new_filt[a_match, ]
        if (ct == 1) { a_match_frame <<- match_subset }
        else { a_match_frame <<- rbind(a_match_frame, match_subset) }
          
      }
    }
  }
  
  d_ovn_short <<- d_ovn[, 5:6]
  
  a_ovn_with_d <<- a_match_frame %>%
    left_join(d_ovn_short, by = c("Program" = "2020_PROG"))
  
  a_ovn_merge <<- a_ovn_with_d %>%
    left_join(l_ovn_stack, by = c("Geography", "2019_PROG" = "Program", "Demo"))
  
  a_ovn_merge <<- a_ovn_merge[, c(-13, -14)]
  colnames(a_ovn_merge) <<- c("Geography", "Call", "Station", "Program", "Demo", "SHour", "SMin", "EHour", "EMin", "Rtg", "Imp", "DOW", "CMonth", "CYear", "PRtg", "PImp")
  
  #Timezone List
  
  tz_list <<- a_ovn_new_filt %>% group_by(Geography) %>% summarize(ct = n())
  tz_list <<- tz_list[, 1]
  tz_list <<- data.frame(lapply(tz_list, as.character), stringsAsFactors=FALSE)
  
  a <- 1
  g <- nrow(tz_list)
  
  for (a in a:g) {
    
    act_geog <- tz_list[a, 1]
    if (act_geog == "BALTIMORE") { tz_list[a, 2] <<- 1 }
    if (act_geog == "CINCINNATI") { tz_list[a, 2] <<- 1 }
    if (act_geog == "CLEVELAND-AKRON (CANTON)") { tz_list[a, 2] <<- 1 }
    if (act_geog == "DENVER") { tz_list[a, 2] <<- 2 }
    if (act_geog == "DETROIT") { tz_list[a, 2] <<- 1 }
    if (act_geog == "INDIANAPOLIS") { tz_list[a, 2] <<- 1 }
    if (act_geog == "KANSAS CITY") { tz_list[a, 2] <<- 2 }
    if (act_geog == "LAS VEGAS") { tz_list[a, 2] <<- 1 }
    if (act_geog == "MIAMI-FT. LAUDERDALE") { tz_list[a, 2] <<- 1 }
    if (act_geog == "MILWAUKEE") { tz_list[a, 2] <<- 2 }
    if (act_geog == "NASHVILLE") { tz_list[a, 2] <<- 2 }
    if (act_geog == "NEW YORK") { tz_list[a, 2] <<- 1 }
    if (act_geog == "NORFOLK-PORTSMTH-NEWPT NWS") { tz_list[a, 2] <<- 1 }
    if (act_geog == "PHOENIX (PRESCOTT)") { tz_list[a, 2] <<- 2 }
    if (act_geog == "SALT LAKE CITY") { tz_list[a, 2] <<- 2 }
    if (act_geog == "SAN DIEGO") { tz_list[a, 2] <<- 1 }
    if (act_geog == "TAMPA-ST. PETE (SARASOTA)") { tz_list[a, 2] <<- 1 }
    if (act_geog == "WEST PALM BEACH-FT. PIERCE") { tz_list[a, 2] <<- 1 }
    
  }



  #Move Day of Week Column
  
  a_ovn_new_filt_new <<- data.frame(a_ovn_new_filt[, 1:5], a_ovn_new_filt[, 12], a_ovn_new_filt[, 6:11])
  a_ovn_new_filt_new[, 13] <<- 0
  colnames(a_ovn_new_filt_new) <<- c("Geography", "Call", "Station", "Program", "Demo", "DOW", "SHour", "SMin", "EHour", "EMin", "Rtg", "Imp", "Scripps")
  
  #Loop to Determine Scripps Flag
  
  num_matrix <<- as.matrix(a_ovn_new_filt_new[, 7:13])
  char_matrix <<- as.matrix(a_ovn_new_filt_new[, 1:6])
  
  
  a <- 1
  g <- nrow(char_matrix)
  
  for (a in a:g) {
    
    act_geog <- as.character(char_matrix[a, 1])
    act_net <- as.character(char_matrix[a, 2])
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
    
    num_matrix[a, 7] <<- s_flag
  }
  
  #Create Day of Week and Hour Matrix
  
  hour_matrix <<- data.frame(char_matrix[, 1:6], num_matrix[, 1:2])
  hour_matrix[, 9:33] <<- 0
  colnames(hour_matrix) <<- c("Geography", "Call", "Station", "Program", "Demo", "DOW", "SHour", "SMin", "MON", "TUE", "WED", "THU", "FRI", "SUN",
                             "MON1", "MON2", "MON3", "TUE1", "TUE2", "TUE3", "WED1", "WED2", "WED3", "THU1", "THU2", "THU3", "FRI1", "FRI2", "FRI3",
                             "SUN1", "SUN2", "SUN3", "SUN4")
  
  a <- 1
  g <- nrow(hour_matrix)
  
  for (a in a:g) {
    
    act_geog <- as.character(hour_matrix[a, 1])
    act_tz <- which(tz_list$Geography == act_geog)
    act_tz <- tz_list[act_tz, 2]
    s_hour <- as.numeric(num_matrix[a, 1])
    e_end <- as.numeric(num_matrix[a, 4])
    e_hour <- as.numeric(num_matrix[a, 3])
    
    if (e_end == 0) { e_hour = e_hour - 1 }
    if (e_hour < s_hour) { e_hour <- 23 }
    
    if (act_tz == 2) { 
      s_hour <- s_hour + 1 
      e_hour <- e_hour + 1
    }
    
    dow <- hour_matrix[a, 6]
    
    if (dow == "M......") { 
      hour_matrix[a, 9] <<- 1
      e <- 14
      ct <- 0
      for (s_hour in s_hour:e_hour) {
        if (s_hour == 20) { hour_matrix[a, e + 1] <<- 1 }
        if (s_hour == 21) { hour_matrix[a, e + 2] <<- 1 }
        if (s_hour == 22) { hour_matrix[a, e + 3] <<- 1 }
      }
    }
    
    if (dow == ".Tu.....") { 
      hour_matrix[a, 10] <<- 1 
      
      e <- 17
      ct <- 0
      for (s_hour in s_hour:e_hour) {
        if (s_hour == 20) { hour_matrix[a, e + 1] <<- 1 }
        if (s_hour == 21) { hour_matrix[a, e + 2] <<- 1 }
        if (s_hour == 22) { hour_matrix[a, e + 3] <<- 1 }
      }
    }
    
    if (dow == "..W....") { 
      hour_matrix[a, 11] <<- 1 
      
      e <- 20
      ct <- 0
      for (s_hour in s_hour:e_hour) {
        if (s_hour == 20) { hour_matrix[a, e + 1] <<- 1 }
        if (s_hour == 21) { hour_matrix[a, e + 2] <<- 1 }
        if (s_hour == 22) { hour_matrix[a, e + 3] <<- 1 }
      }
    }
    if (dow == "...Th...") { 
      hour_matrix[a, 12] <<- 1 
      
      e <- 23
      ct <- 0
      for (s_hour in s_hour:e_hour) {
        if (s_hour == 20) { hour_matrix[a, e + 1] <<- 1 }
        if (s_hour == 21) { hour_matrix[a, e + 2] <<- 1 }
        if (s_hour == 22) { hour_matrix[a, e + 3] <<- 1 }
      }
    }
    if (dow == "....F..") { 
      hour_matrix[a, 13] <<- 1
      
      e <- 26
      ct <- 0
      for (s_hour in s_hour:e_hour) {
        if (s_hour == 20) { hour_matrix[a, e + 1] <<- 1 }
        if (s_hour == 21) { hour_matrix[a, e + 2] <<- 1 }
        if (s_hour == 22) { hour_matrix[a, e + 3] <<- 1 }
      }
    }
    if (dow == "......Su") { 
      hour_matrix[a, 14] <<- 1 
      
      e <- 29
      ct <- 0
      for (s_hour in s_hour:e_hour) {
        if (s_hour == 19) { hour_matrix[a, e + 1] <<- 1 }
        if (s_hour == 20) { hour_matrix[a, e + 2] <<- 1 }
        if (s_hour == 21) { hour_matrix[a, e + 3] <<- 1 }
        if (s_hour == 22) { hour_matrix[a, e + 4] <<- 1 }
      }
    }
  }
  
  #Recombine Frames
  
  a_ovn_recomb <<- data.frame(char_matrix)
  a_ovn_recomb <<- data.frame(lapply(a_ovn_recomb, as.character), stringsAsFactors=FALSE)
  a_ovn_recomb <<- data.frame(a_ovn_recomb, num_matrix)
  a_ovn_recomb[, 14:16] <<- 0
  colnames(a_ovn_recomb)[14:16] <<- c("PrimeRank", "DOWRank", "TSRank")
}

#-----------------------RANKINGS--------------------------
#Ranks

demo_vector <- c("Households", "Persons 18-49", "Persons 25-54", "Females 25-54", "Males 25-54")
geog_vector <- a_ovn_recomb %>%
  group_by(Geography) %>%
  summarize(ct = n())
geog_vector <- data.frame(lapply(geog_vector, as.character), stringsAsFactors=FALSE)
geog_vector <- c(geog_vector[, 1])
dow_vector <- c("M......", ".Tu.....", "..W....", "...Th...", "....F..", "......Su")
timeslot_vector <- c("MON1", "MON2", "MON3", "TUE1", "TUE2", "TUE3", "WED1", "WED2", "WED3",
                     "THU1", "THU2", "THU3", "FRI1", "FRI2", "FRI3", "SUN1", "SUN2", "SUN3", "SUN4")

#Total Prime Rank Calculation

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
      
      tr_frame[, 14] <- tr_rank_vector
      
      if (ct == 1) { tr_frame_with_dp_rank <- tr_frame }
      else { tr_frame_with_dp_rank <- rbind(tr_frame_with_dp_rank, tr_frame) }
      
    }
  }
}

a_ovn_recomb <- tr_frame_with_dp_rank

#Day of Week Rank Calculation

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
    i <- length(dow_vector)
    
    for (c in c:i) {
      
      act_daypart <- dow_vector[c]
      
      tr_frame <- a_ovn_recomb %>%
        filter(Geography == act_geog, Demo == act_demo, DOW == act_daypart) %>%
        arrange(desc(Imp))
      
      if (nrow(tr_frame) > 0) { 
        
        ct <- ct + 1
        tr_frame_length <- nrow(tr_frame)
        tr_rank_vector <- c(1:tr_frame_length)
        
        tr_frame[, 15] <- tr_rank_vector
        
        if (ct == 1) { tr_frame_with_dp_rank <- tr_frame }
        else { tr_frame_with_dp_rank <- rbind(tr_frame_with_dp_rank, tr_frame) }
        
      }
    }
  }
}

a_ovn_recomb <- tr_frame_with_dp_rank

#Timeslot Rank Calculation

a_ovn_recomb_short <- a_ovn_recomb[, c(1:8, 12)]

hr_vector <- c(15:33)
a <- 1
g <- length(hr_vector)
ct <- 0

for (a in a:g) {
  
  act_hour <- hr_vector[a]
  
  b <- 1
  h <- length(geog_vector)
  
  for (b in b:h) {
    
    act_geog <- geog_vector[b]
    act_tz <- which(tz_list$Geography == act_geog)
    act_tz <- tz_list[act_tz, 2]
      
    if (act_hour == 15) { 
      s_hour_recomb <- 20 
      s_dow_recomb <- dow_vector[1]
    }
    if (act_hour == 16) { 
      s_hour_recomb <- 21 
      s_dow_recomb <- dow_vector[1]
    }
    if (act_hour == 17) { 
      s_hour_recomb <- 22 
      s_dow_recomb <- dow_vector[1]
    }
    if (act_hour == 18) { 
      s_hour_recomb <- 20 
      s_dow_recomb <- dow_vector[2]
    }
    if (act_hour == 19) { 
      s_hour_recomb <- 21 
      s_dow_recomb <- dow_vector[2]
    }
    if (act_hour == 20) { 
      s_hour_recomb <- 22 
      s_dow_recomb <- dow_vector[2]
    }
    if (act_hour == 21) { 
      s_hour_recomb <- 20 
      s_dow_recomb <- dow_vector[3]
    }
    if (act_hour == 22) { 
      s_hour_recomb <- 21 
      s_dow_recomb <- dow_vector[3]
    }
    if (act_hour == 23) { 
      s_hour_recomb <- 22 
      s_dow_recomb <- dow_vector[3]
    }
    if (act_hour == 24) { 
      s_hour_recomb <- 20 
      s_dow_recomb <- dow_vector[4]
    }
    if (act_hour == 25) { 
      s_hour_recomb <- 21 
      s_dow_recomb <- dow_vector[4]
    }
    if (act_hour == 26) { 
      s_hour_recomb <- 22 
      s_dow_recomb <- dow_vector[4]
    }
    if (act_hour == 27) { 
      s_hour_recomb <- 20 
      s_dow_recomb <- dow_vector[5]
    }
    if (act_hour == 28) { 
      s_hour_recomb <- 21 
      s_dow_recomb <- dow_vector[5]
    }
    if (act_hour == 29) { 
      s_hour_recomb <- 22 
      s_dow_recomb <- dow_vector[5]
    }
    if (act_hour == 30) { 
      s_hour_recomb <- 19 
      s_dow_recomb <- dow_vector[6]
    }
    if (act_hour == 31) { 
      s_hour_recomb <- 20 
      s_dow_recomb <- dow_vector[6]
    }
    if (act_hour == 32) { 
      s_hour_recomb <- 21 
      s_dow_recomb <- dow_vector[6]
    }
    if (act_hour == 33) { 
      s_hour_recomb <- 22 
      s_dow_recomb <- dow_vector[6]
    }
    
    if (act_tz == 2) { s_hour_recomb <- s_hour_recomb - 1 }
    
    c <- 1
    i <- length(demo_vector)
    
    for (c in c:i) {
      
      act_demo <- demo_vector[c]
      hour_subset <- which(hour_matrix[, 1] == act_geog & hour_matrix[, act_hour] == 1 & hour_matrix[, 5] == act_demo)
      hour_frame <- hour_matrix[hour_subset, 1:8]
      
      if (nrow(hour_frame) > 0) {
        
        ct <- ct + 1
        
        hour_frame_join <- hour_frame %>%
          left_join(a_ovn_recomb_short, by = c("Geography", "Call", "Station", "Program", "Demo", "DOW", "SHour", "SMin")) %>%
          arrange(desc(Imp))
        
        recomb_frame <- a_ovn_recomb %>%
          filter(Geography == act_geog, Demo == act_demo, SHour == s_hour_recomb, DOW == s_dow_recomb)
        
        if (ct == 1) {
          ultra_hour_rank_frame_tmp <- hour_frame_join %>%
            left_join(a_ovn_recomb, by = c("Geography", "Call", "Station", "Program", "Demo", "DOW", "SHour", "SMin", "Imp"))
          ultra_hour_rank_frame_tmp[, 17] <- s_hour_recomb
          ultra_hour_rank_frame_tmp[, 18] <- s_dow_recomb
          ultra_hour_rank_frame <- ultra_hour_rank_frame_tmp
        }
        
        else {
          ultra_hour_rank_frame_tmp <- hour_frame_join %>%
            left_join(a_ovn_recomb, by = c("Geography", "Call", "Station", "Program", "Demo", "DOW", "SHour", "SMin", "Imp"))
          ultra_hour_rank_frame_tmp[, 17] <- s_hour_recomb
          ultra_hour_rank_frame_tmp[, 18] <- s_dow_recomb
          ultra_hour_rank_frame <- rbind(ultra_hour_rank_frame, ultra_hour_rank_frame_tmp)
        }
        
        d <- 1
        j <- nrow(recomb_frame)
        
        if (j > 0) {
          for (d in d:j) {
            act_prog <- as.character(recomb_frame[d, 4])
            act_prog_shour <- as.numeric(recomb_frame[d, 7])
            act_prog_smin <- as.numeric(recomb_frame[d, 8])
            act_prog_call <- as.character(recomb_frame[d, 2])
            act_prog_dow <- as.character(recomb_frame[d, 6])
            hour_frame_join_line <- which(hour_frame_join$Program == act_prog & hour_frame_join$SHour == act_prog_shour & hour_frame_join$SMin == act_prog_smin
                                          & hour_frame_join$Call == act_prog_call & hour_frame_join$DOW == act_prog_dow)
            recomb_frame[d, 16] <- hour_frame_join_line
            
          }
          
          if (ct == 1) { 
            tr_frame_with_hr_rank <- recomb_frame 
          }
          else { 
            tr_frame_with_hr_rank <- rbind(tr_frame_with_hr_rank, recomb_frame) 
          }
        }
      }
    }
  }
}

a_ovn_final <- tr_frame_with_hr_rank

#Merge with NLTV Previous Airing Compare Numbers

l_ovn_merge <- a_ovn_merge %>%
  left_join(a_ovn_final, by = c("Geography", "Call", "Station", "Program", "Demo", "DOW", "SHour", "SMin", "EHour", "EMin", "Rtg", "Imp"))

l_ovn_mutate <-  l_ovn_merge %>%
  mutate(RtgChg = Rtg - PRtg) %>%
  mutate(ImpChg = Imp - PImp) %>%
  mutate(RtgPctChg = (Rtg - PRtg) / PRtg) %>%
  mutate(ImpPctChg = (Imp - PImp) / PImp)

#Create Master Day/Timeslot Rank Frme

a <- 1
g <- nrow(l_ovn_mutate)

for (a in a:g) {
  
  act_geog <- as.character(l_ovn_mutate[a, 1])
  act_demo <- as.character(l_ovn_mutate[a, 5])
  act_shour <- as.numeric(l_ovn_mutate[a, 6])
  act_dp <- as.character(l_ovn_mutate[a, 12])
  
  tmp_hour <- ultra_hour_rank_frame %>%
    filter(Geography == act_geog, Demo == act_demo, V17 == act_shour, V18 == act_dp)
  
  tmp_dp <- a_ovn_final %>%
    filter(Geography == act_geog, Demo == act_demo, DOW == act_dp)
  
  if (a == 1) { 
    master_hour_rank_frame <- tmp_hour
    master_dp_rank_frame <- tmp_dp
  }
  
  else { 
    master_hour_rank_frame <- rbind(master_hour_rank_frame, tmp_hour)
    master_dp_rank_frame <- rbind(master_dp_rank_frame, tmp_dp)
  }
}

prime_scorecard <- function(demo_feed, network_use) {
  
  dow_vector_sc <- c( "......Su", "M......", ".Tu.....", "..W....", "...Th...", "....F..")
  
  l_ovn_mutate_join <<- l_ovn_mutate %>%
    left_join(d_file, by = c("Program" = "2020_PROG"))
  
  l_ovn_mutate_join <<- l_ovn_mutate_join[, c(1:25, 29)]
  colnames(l_ovn_mutate_join)[12] <<- "DOW"
  
  prime_sc <<- l_ovn_mutate_join %>%
    filter(Network == network_use, Demo == demo_feed)
  
  
  a <- 1
  g <- length(dow_vector_sc)
  ct <- 0
  
  for (a in a:g) {
    
    act_dow <- dow_vector_sc[a]
    
    prime_sc_cut <<- prime_sc %>%
      filter(DOW == act_dow)
    
    if (nrow(prime_sc_cut) > 0) {
      
      ct <- ct + 1
      prime_sc_group <<- prime_sc_cut %>%
        group_by(Program, DOW, CMonth, CYear, Network, `2019_PROG`) %>%
        summarize(RtgAvg = mean(Rtg), LYRtgAvg = mean(PRtg)) %>%
        mutate(PctChg = ((RtgAvg - LYRtgAvg) / LYRtgAvg))
      
      pscg_row <- nrow(prime_sc_group) + 1
      
      prime_sc_group[pscg_row, 1] <<- " "
      
      if (ct == 1) { prime_sc_stack <<- prime_sc_group }
      else { prime_sc_stack <<- rbind(prime_sc_stack, prime_sc_group )}
      
    }
  }
  
  meaner_table2 <<- which(is.na(prime_sc_stack$LYRtgAvg) == FALSE)
  meaner_table <<- prime_sc_stack[meaner_table2, ]
  
  avg_thisyear <<- mean(meaner_table$RtgAvg)
  avg_lastyear <<- mean(meaner_table$LYRtgAvg)
  
  avg_pctchg <<- (avg_thisyear - avg_lastyear) / avg_lastyear
  avg_pctchg <<- scales::percent(avg_pctchg, accuracy=0.1)
  
  avg_thisyear <<- as.numeric(format(round(avg_thisyear, digits = 2), nsmall = 2))
  avg_lastyear <<- as.numeric(format(round(avg_lastyear, digits = 2), nsmall = 2))
  
  prime_sc_stack$PctChg <<- as.numeric(prime_sc_stack$PctChg)
  prime_sc_stack$PctChg <<- scales::percent(prime_sc_stack$PctChg, accuracy=0.1)
  prime_sc_stack$RtgAvg <<- format(round(prime_sc_stack$RtgAvg, digits = 1), nsmall = 1)
  prime_sc_stack$LYRtgAvg <<- format(round(prime_sc_stack$LYRtgAvg, digits = 1), nsmall = 1)
  
  prime_sc_stack <<- prime_sc_stack %>%
    select(Network, DOW, Program, RtgAvg, LYRtgAvg, PctChg, CMonth, CYear, `2019_PROG`)
  
  prime_sc_use <<- prime_sc_stack[, c(2:9)]
  psu_row <<- nrow(prime_sc_use) - 1
  prime_sc_use <<- prime_sc_use[1:psu_row, ]
  psu_row <<- psu_row + 1
  
  prime_sc_use[psu_row, ] <<- " "
  prime_sc_use[psu_row, 2] <<- "Average"
  prime_sc_use[psu_row, 3] <<- avg_thisyear
  prime_sc_use[psu_row, 4] <<- avg_lastyear
  prime_sc_use[psu_row, 5] <<- avg_pctchg
  
  if (demo_feed == "Persons 25-54") { demo_bar <<- "A2554 Rating" }
  if (demo_feed == "Men 25-54") { demo_bar <<- "M2554 Rating" }
  if (demo_feed == "Households") { demo_bar <<- "Household Rating" }
  if (demo_feed == "Women 25-54") { demo_bar <<- "W2554 Rating" }
  if (demo_feed == "Persons 18-49") { demo_bar <<- "A1849 Rating" }
  
  prime_sc_use[, 9] <<- prime_sc_use$PctChg
  colnames(prime_sc_use)[9] <<- "z"
  
  prime_sc_use <<- cbind(prime_sc_use[, 1:4], prime_sc_use[, 9], prime_sc_use[, 5:8])
  
  blank_which <<- which(prime_sc_use$Program == " ")
  prime_sc_use[blank_which, ] <<- " "
  
  colnames(prime_sc_use) <<- c("Day", "Program", demo_bar, "Comp Rating", " ", "% Change", "Comp Month", "Comp Year", "Comp Program")
  
  a <- 1
  g <- nrow(prime_sc_use)
  
  for (a in a:g) {
    
    b <- prime_sc_use$Day[a]
    if (b == "......Su") { prime_sc_use$Day[a] <- "SUN" }
    if (b == "M......") { prime_sc_use$Day[a] <- "MON" }
    if (b == ".Tu.....") { prime_sc_use$Day[a] <- "TUE" }
    if (b == "..W....") { prime_sc_use$Day[a] <- "WED" }
    if (b == "...Th...") { prime_sc_use$Day[a] <- "THU" }
    if (b == "....F..") { prime_sc_use$Day[a] <- "FRI" }
    if (b == ".....Sa.") { prime_sc_use$Day[a] <- "SAT" }
    
  }
  
  prime_sc_form <<- formattable(
    
    prime_sc_use,
    
    align = c("c", "l", rep("c", NCOL(prime_sc_use) - 2)),
    list(
      
      `Day` = formatter("span", style = ~ style(color = ifelse(Program == "Average", "black", "darkblue"),                 
       font.weight = ifelse(Program == "Average", "bold", "italic"))),
      
      `Program` = formatter("span", style = ~ style(color = ifelse(Program == "Average", "black", "gray"),                 
              font.weight = "bold")),
      
      ` ` = formatter("span", style = x ~ style(color = ifelse(x == " ", "white", ifelse(x > 0, "green", "red"))),
                      x ~ icontext(ifelse(x > 0, "arrow-up", "arrow-down"), ifelse(x == " ", " ", ifelse(x > 0, "Up", "Down")))),
      
      `% Change` = formatter("span", style = ~ style(color = ifelse(Program == "Average", "black", "sienna"),                 
                             font.weight = ifelse(Program == "Average", "bold", "italic"))),
      
      area(col = c(demo_bar)) ~ color_tile("#fafeff", "#008fb3"),
      area(col = c(`Comp Rating`)) ~ color_tile("#fcfafa", "#615d5d")
      
    ))
  
}

prime_scorecard("Persons 25-54", "ABC")

#Export to Excel

wb <- createWorkbook()
addWorksheet(wb, "KeyExport")
writeData(wb, sheet = "KeyExport", x = l_ovn_mutate)
addWorksheet(wb, "DPRank")
writeData(wb, sheet = "DPRank", x = master_dp_rank_frame)
addWorksheet(wb, "HRRank")
writeData(wb, sheet = "HRRank", x = master_hour_rank_frame)
addWorksheet(wb, "TotalExport")
writeData(wb, sheet = "TotalExport", x = a_ovn_final)
saveWorkbook(wb, "PrimeExport.xlsx", overwrite = TRUE)