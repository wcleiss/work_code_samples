library(readxl)
library(openxlsx)
library(readr)
library(dplyr)
library(tidyr)
library(lubridate)
library(stringr)
library(scales)
library(ggalt)
library(ggplot2)


#Import Arianna Export File and Parse to Remove Uneeded Rows

a_file <- "Total Day Ranker Template Oct 2.xlsx"
l_file <- "NBA Finals 2.xlsx"
a_prog_name <- "NBA FINAL-10/2"
l_prog_name <- "NBA FINAL-10/2"
full_name <- "2020 NBA Finals Game 2"
past_name <- "2019 NBA Finals Game 2"

a_ovn <- read_excel(a_file, sheet = "Live+Same Day, Custom Range 1, ")
l_ovn <- read_excel(l_file, sheet = "Live+Same Day, Average")

a_ovn_len <- nrow(a_ovn)
a_ovn_len_short <- a_ovn_len - 6

a_ovn <- a_ovn[11:a_ovn_len_short, ]
a_ovn <- a_ovn[, c(1:3, 5, 8:11)]

#Change 0 Ratings and Imps to 0.01 and 10

a_ovn_rtg_mx <- as.matrix(a_ovn[, 5])
a_ovn_rtg_mx <- matrix(sapply(a_ovn_rtg_mx, as.numeric))
a <- 1
g <- nrow(a_ovn_rtg_mx)
for (a in a:g) {
  
  b <- as.numeric(a_ovn_rtg_mx[a, 1])
  if (b <= 0) { 
    a_ovn_rtg_mx[a, 1] <- 0.01
  }
}
a_ovn_imp_mx <- as.matrix(a_ovn[, 6])
a_ovn_imp_mx <- matrix(sapply(a_ovn_imp_mx, as.numeric))
a <- 1
g <- nrow(a_ovn_imp_mx)
for (a in a:g) {
  
  b <- as.numeric(a_ovn_imp_mx[a, 1])
  if (b <= 0) { 
    a_ovn_imp_mx[a, 1] <- 10
  }
}



#Loop to separate Call from Station

a_ovn_mx <- as.matrix(a_ovn[, 2])
a_ovn_new_mx <- a_ovn_mx

a <- 1
g <- nrow(a_ovn_mx)

for (a in a:g) {
 
  b <- as.character(a_ovn_mx[a, 1])
  c <- strsplit(b, split = " ")[[1]]
  f <- c[[1]]
  h <- strsplit(f, split = "-")[[1]]
  k <- h[[1]]
  a_ovn_new_mx[a, 1] <- k
   
}

#Loop to Rename Demos

a_ovn_demo_mx <- as.matrix(a_ovn[ ,4])
a <- 1
g <- nrow(a_ovn_demo_mx)

for (a in a:g) {
  
  b <- as.character(a_ovn_demo_mx[a, 1])
  if (b == "HH") { k <- "Households" }
  if (b == "P25-54") { k <- "Persons 25-54" }
  if (b == "P18-49") { k <- "Persons 18-49" }
  if (b == "M25-54") { k <- "Males 25-54" }
  if (b == "F25-54") { k <- "Females 25-54" }
  
  a_ovn_demo_mx[a, 1] <- k
  
}

#Loop to Create Timeslots

a_ovn_stime_mx <- as.matrix(a_ovn[, 7])
a_ovn_stime_mx <- matrix(sapply(a_ovn_stime_mx, as.numeric))
a_ovn_stime_mx <- a_ovn_stime_mx + .000001
a_ovn_stime_hour <- a_ovn_stime_mx * 24
a_ovn_stime_hour_int <- as.integer(a_ovn_stime_hour)
a_ovn_stime_min <- (a_ovn_stime_hour - a_ovn_stime_hour_int) * 60
a_ovn_stime_mx_final <- cbind(a_ovn_stime_hour_int, as.integer(a_ovn_stime_min))

a_ovn_etime_mx <- as.matrix(a_ovn[, 8])
a_ovn_etime_mx <- matrix(sapply(a_ovn_etime_mx, as.numeric))
a_ovn_etime_mx <- a_ovn_etime_mx + .000001
a_ovn_etime_hour <- a_ovn_etime_mx * 24
a_ovn_etime_hour_int <- as.integer(a_ovn_etime_hour)
a_ovn_etime_min <- (a_ovn_etime_hour - a_ovn_etime_hour_int) * 60
a_ovn_etime_mx_final <- cbind(a_ovn_etime_hour_int, as.integer(a_ovn_etime_min))

#Create Final Dataframe

a_ovn_almost <- data.frame(a_ovn[, 1], a_ovn_new_mx, a_ovn[, c(2:3)], a_ovn_demo_mx, a_ovn_stime_mx_final, a_ovn_etime_mx_final, a_ovn_rtg_mx, a_ovn_imp_mx)

colnames(a_ovn_almost) <- c("Geography", "Call", "Station", "Program", "Demo", "n_s_hr", "n_s_mm", "n_e_hr", "n_e_mm", "r_rtg", "r_imp")
a_ovn_new_filt <- a_ovn_almost %>%
  filter(Call != "WWSB")



#-----------JOIN WITH ARIANNA WING SCRIPT-----------------


#Timezone List

tz_list <- a_ovn_new_filt %>% group_by(Geography) %>% summarize(ct = n())
a <- 1
g <- nrow(tz_list)

for (a in a:g) {
  
  act_geog <- tz_list[a, 1]
  if (act_geog == "Baltimore") { tz_list[a, 2] <- 1 }
  if (act_geog == "Cincinnati") { tz_list[a, 2] <- 1 }
  if (act_geog == "Cleveland-Akron (Canton)") { tz_list[a, 2] <- 1 }
  if (act_geog == "Denver") { tz_list[a, 2] <- 2 }
  if (act_geog == "Detroit") { tz_list[a, 2] <- 1 }
  if (act_geog == "Indianapolis") { tz_list[a, 2] <- 1 }
  if (act_geog == "Kansas City") { tz_list[a, 2] <- 2 }
  if (act_geog == "Las Vegas") { tz_list[a, 2] <- 1 }
  if (act_geog == "Miami-Ft. Lauderdale") { tz_list[a, 2] <- 1 }
  if (act_geog == "Milwaukee") { tz_list[a, 2] <- 2 }
  if (act_geog == "Nashville") { tz_list[a, 2] <- 2 }
  if (act_geog == "New York") { tz_list[a, 2] <- 1 }
  if (act_geog == "Norfolk-Portsmth-Newpt Nws") { tz_list[a, 2] <- 1 }
  if (act_geog == "Phoenix (Prescott)") { tz_list[a, 2] <- 2 }
  if (act_geog == "Salt Lake City") { tz_list[a, 2] <- 2 }
  if (act_geog == "San Diego") { tz_list[a, 2] <- 1 }
  if (act_geog == "Tampa-St. Pete (Sarasota)") { tz_list[a, 2] <- 1 }
  if (act_geog == "West Palm Beach-Ft. Pierce") { tz_list[a, 2] <- 1 }
  
}

#Add Column for Daypart, Scripps ID

a_ovn_new_filt[, 12:13] <- 0

#Loop to Determine Daypart and Scripps Flag

num_matrix <- as.matrix(a_ovn_new_filt[, 6:13])
char_matrix <- as.matrix(a_ovn_new_filt[, 1:5])


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
  
  num_matrix[a, 7] <- dp_use
  num_matrix[a, 8] <- s_flag
}

#Create Hour Matrix

hour_matrix <- data.frame(char_matrix[, 1:5], num_matrix[, 1:2])
hour_matrix[, 8:31] <- 0
colnames(hour_matrix) <- c("Geography", "Call", "Station", "Program", "Demo", "SHour", "SMin", "H0", "H1", "H2", "H3", "H4", "H5", "H6", "H7",
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
    
    hr_slot <- s_hour + 8
    hour_matrix[a, hr_slot] <- 1
    
  }
}

#Recombine and add 3 columns for Rank

a_ovn_recomb <- data.frame(char_matrix)
a_ovn_recomb <- data.frame(lapply(a_ovn_recomb, as.character), stringsAsFactors=FALSE)
a_ovn_recomb <- data.frame(a_ovn_recomb, num_matrix)
a_ovn_recomb[, 14:16] <- 0
geog_vector <- tz_list[, 1]
geog_vector <- data.frame(lapply(geog_vector, as.character), stringsAsFactors=FALSE)
geog_vector <- c(geog_vector[, 1])


colnames(a_ovn_recomb)[6:16] <- c("SHour", "SMin", "EHour", "EMin", "Rtg", "Imp", "Daypart", "Scripps", "TotRank", "DPRank", "HRRank")

#Total Rank Calculation

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
      
      tr_frame[, 14] <- tr_rank_vector
      
      if (ct == 1) { tr_frame_with_dp_rank <- tr_frame }
      else { tr_frame_with_dp_rank <- rbind(tr_frame_with_dp_rank, tr_frame) }
      
    }
  }
}

a_ovn_recomb <- tr_frame_with_dp_rank

#Daypart Rank Calculation

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
        
        tr_frame[, 15] <- tr_rank_vector
        
        if (ct == 1) { tr_frame_with_dp_rank <- tr_frame }
        else { tr_frame_with_dp_rank <- rbind(tr_frame_with_dp_rank, tr_frame) }
        
      }
    }
  }
}

a_ovn_recomb <- tr_frame_with_dp_rank

#Hour Rank Calculation

a_ovn_recomb_short <- a_ovn_recomb[, c(1:7, 11)]

hr_vector <- c(8:31)
a <- 1
g <- length(hr_vector)
ct <- 0

for (a in a:g) {
  
  act_hour <- hr_vector[a]
  
  b <- 1
  h <- length(geog_vector)
  
  for (b in b:h) {
    
    act_geog <- geog_vector[b]
    
    c <- 1
    i <- length(demo_vector)
    
    for (c in c:i) {
      
      act_demo <- demo_vector[c]
      hour_subset <- which(hour_matrix[, 1] == act_geog & hour_matrix[, act_hour] == 1 & hour_matrix[, 5] == act_demo)
      hour_frame <- hour_matrix[hour_subset, 1:7]
      
      if (nrow(hour_frame) > 0) {
        
        ct <- ct + 1
        
        hour_frame_join <- hour_frame %>%
          left_join(a_ovn_recomb_short, by = c("Geography", "Call", "Station", "Program", "Demo", "SHour", "SMin")) %>%
          arrange(desc(Imp))
        
        actual_hour <- act_hour - 8
        
        recomb_frame <- a_ovn_recomb %>%
          filter(Geography == act_geog, Demo == act_demo, SHour == actual_hour)
        
        if (ct == 1) {
          ultra_hour_rank_frame_tmp <- hour_frame_join %>%
            left_join(a_ovn_recomb, by = c("Geography", "Call", "Station", "Program", "Demo", "SHour", "SMin", "Imp"))
          ultra_hour_rank_frame_tmp[, 17] <- actual_hour
          ultra_hour_rank_frame <- ultra_hour_rank_frame_tmp
        }
        
        else {
          ultra_hour_rank_frame_tmp <- hour_frame_join %>%
            left_join(a_ovn_recomb, by = c("Geography", "Call", "Station", "Program", "Demo", "SHour", "SMin", "Imp"))
          ultra_hour_rank_frame_tmp[, 17] <- actual_hour
          ultra_hour_rank_frame <- rbind(ultra_hour_rank_frame, ultra_hour_rank_frame_tmp)
        }
        
        d <- 1
        j <- nrow(recomb_frame)
        
        if (j > 0) {
          for (d in d:j) {
            act_prog <- as.character(recomb_frame[d, 4])
            act_prog_shour <- as.numeric(recomb_frame[d, 6])
            act_prog_smin <- as.numeric(recomb_frame[d, 7])
            act_prog_call <- as.character(recomb_frame[d, 2])
            hour_frame_join_line <- which(hour_frame_join$Program == act_prog & hour_frame_join$SHour == act_prog_shour & hour_frame_join$SMin == act_prog_smin
                                          & hour_frame_join$Call == act_prog_call)
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


l_ovn_len <- nrow(l_ovn)
l_ovn_len2 <- l_ovn_len - 6
l_ovn_trim <- l_ovn[12:l_ovn_len2, ]
l_ovn_trim2 <- l_ovn_trim[, c(-6, -9)]
l_ovn_trim_num <- data.frame(lapply(l_ovn_trim2[, 6:7], as.numeric), stringsAsFactors=FALSE)
l_ovn_trim2 <- cbind(l_ovn_trim2[, 1:5], l_ovn_trim_num)
colnames(l_ovn_trim2) <- c("Geography", "Call", "Program", "Daypart", "Demo", "Rtg", "Imp")

a <- 1
g <- nrow(l_ovn_trim2)

for (a in a:g) {
  
  b <- as.character(l_ovn_trim2[a, 5])
  if (b == "HH") { l_ovn_trim2[a, 5] <- "Households" }
  if (b == "P25-54") { l_ovn_trim2[a, 5] <- "Persons 25-54" }
  if (b == "P18-49") { l_ovn_trim2[a, 5] <- "Persons 18-49" }
  if (b == "M25-54") { l_ovn_trim2[a, 5] <- "Males 25-54" }
  if (b == "F25-54") { l_ovn_trim2[a, 5] <- "Females 25-54" }
  
}

a_ovn_final_short <- a_ovn_final %>%
  filter(Program == a_prog_name)

a_ovn_combine <- a_ovn_final_short %>%
  left_join(l_ovn_trim2, by = c("Geography", "Demo"))

a_ovn_combine2 <- a_ovn_combine[, c(-17, -18, -19)]

a_ovn_mutate1 <- a_ovn_combine2 %>%
  mutate(RtgChg = Rtg.x - Rtg.y) %>%
  mutate(ImpChg = Imp.x - Imp.y) %>%
  mutate(RtgPctChg = (Rtg.x - Rtg.y) / Rtg.y) %>%
  mutate(ImpPctChg = (Imp.x - Imp.y) / Imp.y)

colnames(a_ovn_mutate1)[10:11] <- c("Rtg", "Imp")
colnames(a_ovn_mutate1)[17:18] <- c("PrevRtg", "PrevImp")

#Create Master Hour Rank Frake

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
    master_hour_rank_frame <- tmp_hour
    master_dp_rank_frame <- tmp_dp
  }
  
  else { 
    master_hour_rank_frame <- rbind(master_hour_rank_frame, tmp_hour)
    master_dp_rank_frame <- rbind(master_dp_rank_frame, tmp_dp)
  }
}

a_ovn_database <- a_ovn_mutate1 %>% 
  group_by(Geography, Demo) %>%
  summarize(ct = n())


#Export to Excel

wb <- createWorkbook()
addWorksheet(wb, "KeyExport")
writeData(wb, sheet = "KeyExport", x = a_ovn_mutate1)
addWorksheet(wb, "DPRank")
writeData(wb, sheet = "DPRank", x = master_dp_rank_frame)
addWorksheet(wb, "HRRank")
writeData(wb, sheet = "HRRank", x = master_hour_rank_frame)
addWorksheet(wb, "TotalExport")
writeData(wb, sheet = "TotalExport", x = a_ovn_final)
addWorksheet(wb, "Database")
writeData(wb, sheet = "Database", x = a_ovn_database)
saveWorkbook(wb, "AriannaExport.xlsx", overwrite = TRUE)


#Market Specific Slides

#Rating Overview

act_geo = "Las Vegas"

geo_spec_frame <- a_ovn_mutate1 %>%
  filter(Geography == act_geo)
geo_spec_frame <- geo_spec_frame[order(desc(geo_spec_frame$Rtg)), ]
geo_spec_frame$Demo <- factor(geo_spec_frame$Demo, levels = geo_spec_frame$Demo) 

overview_plot <- ggplot(geo_spec_frame, aes(x = Demo, y = Rtg, fill = Demo)) +
                geom_bar(stat="identity", width=.5) +
                geom_text(color = "black", size=4, aes(label = round(Rtg,1)), vjust=1)

#Pct Change Dumbell Plot

geo_spec_frame <- geo_spec_frame[order(geo_spec_frame$Rtg), ]
geo_spec_frame$Demo <- factor(geo_spec_frame$Demo, levels=as.character(geo_spec_frame$Demo))

pct_chg_plot <- ggplot(geo_spec_frame, aes(x=PrevRtg, xend=Rtg, y=Demo, group=Demo, label)) + 
  geom_dumbbell(color="#a3c4dc", 
                size=0.75, 
                colour_x = "gray",
                size_x = 3.5,
                colour_xend="dark blue",
                size_xend = 3.5) + 
  scale_x_continuous() + 
  labs(x=NULL, 
       y=NULL, 
       title="Rating Change", 
       subtitle=paste(full_name, "vs.", past_name, sep = " "), 
       caption="Source: https://github.com/hrbrmstr/ggalt") +
  geom_text(color="black", size=3, hjust=1.5,
            aes(x=PrevRtg, label=format(round(PrevRtg,1), nsmall=1)))+
  geom_text(aes(x=((Rtg+PrevRtg)/2), label=percent(RtgPctChg, accuracy = 0.1)), 
            color="darkred", size=4, hjust=0, vjust = -1)+
  geom_text(aes(x=Rtg, label=format(round(Rtg,1), nsmall=1)), 
            color="darkblue", size=4, hjust=-0.2)

#Pct Change Dumbell Positive O nly


}


#Total Market Rank Dot Plot

totrank_ct <- a_ovn_final %>%
  filter(Geography == act_geo, Demo == "Households")
totrank_ct <- nrow(totrank_ct)
chart_title <- paste(full_name, "Total Day Rank", sep = " ")
sub_chart_title <- paste("Market Rank Among", totrank_ct, "Aired Programs For the Day", sep = " ")

geo_spec_frame <- geo_spec_frame[order(desc(geo_spec_frame$TotRank)), ]
geo_spec_frame$Demo <- factor(geo_spec_frame$Demo, levels = geo_spec_frame$Demo) 

totrank_plot <- ggplot(geo_spec_frame, aes(x=Demo, y=TotRank, label=TotRank, col=Demo)) + 
  geom_point(size=8) +   # Draw points +
  geom_text(color="white", size=4) +
  geom_segment(aes(x=Demo, 
                   xend=Demo, 
                   y=1, 
                   yend=totrank_ct), 
               linetype="dashed", 
               size=0.1) +   # Draw dashed lines
  labs(title=chart_title, 
       subtitle=sub_chart_title, 
       caption="source: Nielsen NLTV") +  
  coord_flip()

#Daypart Market Rank Dot Plot

act_dp <- geo_spec_frame$Daypart[1]
totrank_dp <- master_dp_rank_frame %>%
  filter(Geography == act_geo, Demo == "Households", Daypart == act_dp)
totrank_dp <- nrow(totrank_dp)
chart_title <- paste(full_name, "Daypart Rank", sep = " ")
sub_chart_title <- paste("Market Rank Among", totrank_dp, "Aired Programs in the Daypart", sep = " ")

geo_spec_frame <- geo_spec_frame[order(desc(geo_spec_frame$DPRank)), ]
geo_spec_frame$Demo <- factor(geo_spec_frame$Demo, levels = geo_spec_frame$Demo) 

dprank_plot <- ggplot(geo_spec_frame, aes(x=Demo, y=DPRank, label=DPRank, col=Demo)) + 
  geom_point(size=8) +   # Draw points +
  geom_text(color="white", size=4) +
  geom_segment(aes(x=Demo, 
                   xend=Demo, 
                   y=1, 
                   yend=totrank_dp), 
               linetype="dashed", 
               size=0.1) +   # Draw dashed lines
  labs(title=chart_title, 
       subtitle=sub_chart_title, 
       caption="source: Nielsen NLTV") +  
  coord_flip()

#Timeslot Rank Dot Plot

totrank_hr <- master_hour_rank_frame %>%
  filter(Geography == act_geo, Demo == "Households")
totrank_hr <- nrow(totrank_hr)
chart_title <- paste(full_name, "Timeslot Rank", sep = " ")
sub_chart_title <- paste("Market Rank Among", totrank_hr, "Aired Programs in the Timeslot", sep = " ")

geo_spec_frame <- geo_spec_frame[order(desc(geo_spec_frame$DPRank)), ]
geo_spec_frame$Demo <- factor(geo_spec_frame$Demo, levels = geo_spec_frame$Demo) 

hrrank_plot <- ggplot(geo_spec_frame, aes(x=Demo, y=HRRank, label=HRRank, col=Demo)) + 
  geom_point(size=8) +   # Draw points +
  geom_text(color="white", size=4) +
  geom_segment(aes(x=Demo, 
                   xend=Demo, 
                   y=1, 
                   yend=totrank_hr), 
               linetype="dashed", 
               size=0.1) +   # Draw dashed lines
  labs(title=chart_title, 
       subtitle=sub_chart_title, 
       caption="source: Nielsen NLTV") +  
  coord_flip()

#Total Day Rank Dot Plot

geo_spec_ranker <- a_ovn_final %>%
  filter(Geography == act_geo, Demo == "Persons 25-54") %>%
  arrange(desc(Rtg))

if (nrow(geo_spec_ranker) > 12) {
  geo_spec_ranker <- geo_spec_ranker[1:12, ]
}

geo_spec_ranker <- geo_spec_ranker[order(geo_spec_ranker$Rtg), ]
geo_spec_ranker$Program <- factor(geo_spec_ranker$Program, levels = geo_spec_ranker$Program) 

totday_dot_plot <- ggplot(geo_spec_ranker, aes(x=Program, y=Rtg, label=format(round(Rtg,1), nsmall=1), col=Call)) + 
  geom_point(size=9) +   # Draw points +
  geom_segment(aes(x=Program, 
                   xend=Program, 
                   y=min(Rtg), 
                   yend=max(Rtg)), 
               linetype="dashed", 
               size=0.2,
               color="black") +   # Draw dashed lines
  labs(title=chart_title, 
       subtitle=sub_chart_title, 
       caption="source: Nielsen NLTV") +  
  geom_text(color="black", size=4) +
  coord_flip()



print(act_geogbind)
chart_title <<- paste(active_date, act_prog, "Grows in Key Demos", sep = " ")
chart_subtitle <<- paste("Viewership Increase Compared to ", act_comp, sep = "")
xxmin <<- min(db_group1$LYIMP) * .98
xxmax <<- max(db_group1$Imp) * 1.02
pct_chg_plot_pos <<- ggplot(db_group1, aes(x=LYIMP, xend=Imp, y=Demo, group=Demo, label)) + 
  geom_segment(aes(x=LYIMP, 
                   xend=Imp, 
                   y=Demo, 
                   yend=Demo), 
               color="lightgreen", size=1.5) +
  
  geom_dumbbell(color="lightgreen", 
                size_x = 4,
                size_xend = 5,
                colour_x = "gray",
                colour_xend="darkgreen") + 
  
  scale_x_continuous() + 
  labs(x=NULL, 
       y=NULL, 
       title=chart_title, 
       subtitle=chart_subtitle, 
       caption="Source: Nielsen NLTV") +
  geom_text(color="black", size=3, hjust=1.5, vjust=2,
            aes(x=LYIMP, label=scales::comma(LYIMP)))+
  geom_text(aes(x=((Imp+LYIMP)/2), label=scales::percent(ImpPctChg, accuracy = 0.1)), 
            color="darkolivegreen", size=4, hjust=0, vjust = -1)+
  geom_text(aes(x=Imp, label=scales::comma(Imp)), 
            color="darkgreen", size=4, hjust=-0.5, vjust=2) +
  xlim(xxmin, xxmax) + 
  xlab("Viewers")