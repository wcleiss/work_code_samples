#The New LP is an R project that I single handedly built to take raw Nielsen data from our Snowflake database and transform it into a cleaned, structured file
  #know as THEGRID. THEGRID is a file that measures many metrics from the Nielsen Data. The R script's main job is to merge the right programs with the proper
  #Time Period Data for trending purposes, and then rank the program's market performance across 20 different ranking metrics. It also calculates dayparts by
  #Time Zone and also whether or not the station is a Scripps Station.

#The Final Output is a 4 tab excel file which contains a massive tab known as THEGRID, which contains all PPM and Set Meter market programming performance.
#The Final Output's other 3 tabs highlight 10 success stories for each geography in the report for News, Prime, and Strip Programming.

library(readxl)
library(openxlsx)
library(readr)
library(dplyr)
library(tidyr)
library(lubridate)

#These Variables are Manually Changed Based on Desired Months and Years to trend against.

mo1 <- "JULY"
mo2 <- "JUNE"
mo3 <- "JULY"

yr1 <- 2020
yr2 <- 2020
yr3 <- 2019

#Program.csv is the raw data export from the Snowflake server containing raw Nielsen data scraped from the Nielsen API.
#The Cleaning Process Begins Below.

program_data <- read_csv("Program.csv", col_names = TRUE)

colnames(program_data)[9] <- "IMPRESSIONS"
colnames(program_data)[10] <- "RATING"
colnames(program_data)[11] <- "SHARE"
colnames(program_data)[19] <- "TELECASTS"

#These frames are databases of the Scripps Markets and their timezones.
geog_frame <- read_csv("GEOFRAME.csv", col_names = TRUE)
etzdp_frame <- read_csv("ETZDP.csv", col_names = TRUE)
ctzdp_frame <- read_csv("CTZDP.csv", col_names = TRUE)
rank_frame <- read_csv("RANKS.csv", col_names = TRUE)

#Assign the proper time zones for local news start and end times.

program_data2 <- program_data %>%
  full_join(geog_frame[,2:3], by = c("STATION"))

program_data_e <- program_data2 %>%
  filter(TIMEZONE == "EST" | TIMEZONE == "PST") %>%
  inner_join(etzdp_frame, by = c("START_TIME_HH", "START_TIME_HFHR"))

program_data_c <- program_data2 %>%
  filter(TIMEZONE == "CST" | TIMEZONE == "MST") %>%
  inner_join(ctzdp_frame, by = c("START_TIME_HH", "START_TIME_HFHR"))

#Determine if programming is Local News, Prime, or Run of Mill Strip programming

program_data3 <- rbind(program_data_e, program_data_c)

news_prime <- which(program_data3$DAYPART == 8 & program_data3$GENRE == "LN - LOCAL NEWS")
program_data3[news_prime, 22] <- 9

program_prime <- program_data3 %>%
  filter(DAYPART == 8 & GENRE != "LN - LOCAL NEWS")

program_runofmill <- program_data3 %>%
  filter(DAYPART != 8 | GENRE == "LN - LOCAL NEWS")

#Determine the Day of Week the program airs, or if it is a Monday-Friday strip program

program_runofmill$DAY_OF_WEEK <- "M-F"

program_weekin <- rbind(program_prime, program_runofmill)
program_weekin <- program_weekin %>%
  filter(VA_NIELSEN_MONTH_NAME == mo1 & VA_NIELSEN_YEAR_NUMBER == yr1) %>%
  select(GEOGRAPHY, PROGRAM_NAME, DEMO, START_TIME_HH, VA_NETWORK_AFFILIATION, DAY_OF_WEEK) %>%
  group_by(GEOGRAPHY, PROGRAM_NAME, DEMO, START_TIME_HH, VA_NETWORK_AFFILIATION, DAY_OF_WEEK) %>%
  summarize(ct = n())

program_weekin <- program_weekin[, c(1:6)]

#Since data is broken out day by day, this process groups each individual airing and calculates average rating, share, and viewers

program_data33 <- program_data3[, -20]

program_data33 <- program_data33 %>%
  group_by(GEOGRAPHY, PROGRAM_NAME, DEMO, START_TIME, START_TIME_HH, START_TIME_QTRH, START_TIME_HFHR,
           GENRE, TIMEZONE, STATION, VA_NIELSEN_YEAR_NUMBER, VA_NIELSEN_MONTH_NAME, VA_NIELSEN_MONTH_NUMBER,
           VA_NETWORK_AFFILIATION, VA_QTR_HR_START_TIME, SCRIPPS, DAYPART) %>%
  summarize(IMPR = sum(IMPRESSIONS), RATI = sum(RATING), SHAR = sum(SHARE), TELECAST = sum(TELECASTS)) %>%
  mutate(IMPRESSIONS = IMPR / TELECAST) %>%
  mutate(RATING = RATI / TELECAST) %>%
  mutate(SHARE = SHAR / TELECAST)

ccc <- c(-18, -19, -20)

program_data33 <- program_data33[, ccc]

#Determine the network affiliaton of the station

program_weekin3 <- program_data33 %>%
  left_join(program_weekin, by = c("GEOGRAPHY", "PROGRAM_NAME", "DEMO", "START_TIME_HH",
                                    "VA_NETWORK_AFFILIATION"))
  
#This process filters out sport events and other airings that only aired once or twice.
program_data3 <- program_weekin3
program_data4 <- program_data3 %>%
  filter(TELECAST > 2 & DAYPART != 8 & GENRE != "SE - SPORTS EVENT")

program_data5 <- program_data3 %>%
  filter(TELECAST > 1 & DAYPART == 8 & GENRE != "SE - SPORTS EVENT")

program_data6 <- program_data3 %>%
  filter(TELECAST > 1 & GENRE == "SE - SPORTS EVENT")

program_data3 <- rbind(program_data4, program_data5, program_data6)

#Split the program data into 3 frames by time period(month). 

mprog_frame_1 <- program_data3 %>%
  filter(VA_NIELSEN_MONTH_NAME == mo1 & VA_NIELSEN_YEAR_NUMBER == yr1)

mprog_frame_2 <- program_data3 %>%
  filter(VA_NIELSEN_MONTH_NAME == mo2 & VA_NIELSEN_YEAR_NUMBER == yr2)

mprog_frame_3 <- program_data3 %>%
  filter(VA_NIELSEN_MONTH_NAME == mo3 & VA_NIELSEN_YEAR_NUMBER == yr3)

prog_keep <- c(1, 2, 3, 4, 19, 20, 21)

mprog_frame_2 <- mprog_frame_2[, prog_keep]
mprog_frame_3 <- mprog_frame_3[, prog_keep]

#Join program's past airings from Time Periods 2 and 3 for trending purposes.

join_frame2 <- mprog_frame_1 %>%
  left_join(mprog_frame_2, 
  by = c("GEOGRAPHY", "PROGRAM_NAME", "DEMO", "START_TIME"),
  suffix = c("_R1", "_R2"))

join_frame3 <- join_frame2 %>%
  left_join(mprog_frame_3, 
  by = c("GEOGRAPHY", "PROGRAM_NAME", "DEMO", "START_TIME"),
  suffix = c("_R1", "_R3"))

tf_col <- ncol(join_frame3)

colnames(join_frame3)[(tf_col - 2)] <- "IMP_R3"
colnames(join_frame3)[(tf_col - 1)] <- "RTG_R3"
colnames(join_frame3)[(tf_col - 0)] <- "SHR_R3"

colnames(join_frame3)[19] <- "IMP_R1"
colnames(join_frame3)[20] <- "RTG_R1"
colnames(join_frame3)[21] <- "SHR_R1"

colnames(join_frame3)[23] <- "IMP_R2"
colnames(join_frame3)[24] <- "RTG_R2"
colnames(join_frame3)[25] <- "SHR_R2"

#Calculate Impression, Rating, and Percentage Change from Time Period 1 to 2 and Time Period 1 to 3.

join_frame4 <- join_frame3 %>% 
  mutate("R1R2_IMP" = IMP_R1 - IMP_R2) %>%
  mutate("R1R2_RTG" = RTG_R1 - RTG_R2) %>%
  mutate("R1R2_PCT" = (IMP_R1 - IMP_R2) / IMP_R2) %>%
  mutate("R1R3_IMP" = IMP_R1 - IMP_R3) %>%
  mutate("R1R3_RTG" = RTG_R1 - RTG_R3) %>%
  mutate("R1R3_PCT" = (IMP_R1 - IMP_R3) / IMP_R3)

join_frame4$R1R3_PCT <- sapply(join_frame4$R1R3_PCT, replace_inf)
join_frame4$R1R2_PCT <- sapply(join_frame4$R1R2_PCT, replace_inf)

replace_inf <- function(x) {
  if (is.nan(x) == TRUE) {
    return(-1)
  }
  else if (is.infinite(x) == TRUE) {
    return(1)
  }
  else if (is.na(x) == TRUE) {
    return(NA)
  }
  else {
    return(x)
  }
}

#Add 20 blank columns for rank columns
#20 different ranking metrics rank by genre, daypart, timeslot, day split by viewership rank, month over month change rank % and raw, year over year change rank % and raw

x <- ncol(join_frame4) + 1
y <- ncol(join_frame4) + 20

join_frame4[,x:y] <- NA

bb <- colnames(rank_frame)
colnames(join_frame4)[x:y] <- bb

r_ct <- nrow(join_frame4)
r_vector <- c(1:r_ct)
r_frame <- as.data.frame(r_vector)

join_frame4 <- as.data.frame(join_frame4)
join_frame4 <- as.data.frame(cbind(r_frame, join_frame4))

#---------------
#VECTOR CREATION
#----------------

#create vectors to iterate through in the looping process.

geo_vector <- join_frame4 %>%
  dplyr::group_by(GEOGRAPHY) %>%
  dplyr::summarize(ct = n())

geo_vector <- geo_vector[,1]

daypart_vector <- c(1:10)
genre_vector <- c("LN - LOCAL NEWS", "SE - SPORTS EVENT")
demo_vector <- c("HH", "A25_54", "A18_49", "M25_54", "W25_54", "A18_34")
hour_vector <- c("00", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", 
                 "14", "15", "16", "17", "18", "19", "20", "21", "22", "23")
dow_vector <- c("M......", ".Tu.....", "..W....", "...Th...", "....F..", "......Su")

#-----------------
#DAYPART RANKS
#-----------------

test_frame <- join_frame4
tmp_frame1 <- test_frame[,1]
tmp_frame2 <- test_frame[,36:55]
tmp_frame3 <- as.matrix(cbind(tmp_frame1, tmp_frame2))
tmp_frame4 <- test_frame[,2:35]

col_vector <- c(7:11)
sort_vector <- c("IMP_R1", "R1R2_IMP", "R1R3_IMP", "R1R2_PCT", "R1R3_PCT")
sortcol_vector <- sapply(sort_vector, colname_no)

colname_no <- function(x) {
  
  return(which(colnames(join_frame4) == x))
  
}

x <- 1
y <- nrow(geo_vector)

for (x in 1:y) {

  act_geo <- as.character(geo_vector[x,1])
  
  test_frame2 <- test_frame %>%
    filter(GEOGRAPHY == act_geo)
  
  b <- 1
  h <- length(daypart_vector)
  
  for (b in 1:h) {
  
    act_dp <- as.numeric(daypart_vector[b])
    
    test_frame4 <- test_frame2 %>%
      filter(DAYPART == act_dp)
    
    c <- 1
    j <- length(demo_vector)
    
    for (c in 1:j) {
      
      act_demo <- as.character(demo_vector[c])
      test_frame5 <- test_frame4 %>%
        filter(DEMO == act_demo)
      
      d <- 1
      k <- length(sort_vector)
      
      for (d in 1:k) {
  
        act_sort <- sort_vector[d]
        cn <- as.numeric(col_vector[d])
        srt_no <- as.numeric(sortcol_vector[d])

        test_frame3 <- test_frame5 %>%
          select(r_vector, !!act_sort)
        
        test_frame3 <- test_frame3[order(test_frame3[,2], decreasing = TRUE),]
        
        test_frame3 <- test_frame3 %>%
          select(r_vector)
        
        a <- 1
        g <- nrow(test_frame3)
        
        for (a in 1:g) {
          rn <- test_frame3[a, 1]
          xx <- test_frame[rn, srt_no]
          if (is.na(xx) == FALSE) {
            tmp_frame3[rn, cn] <- a
            
          }
        }
      }
    }
  }
}

tmp_drame1 <- tmp_frame3[,1]
tmp_drame2 <- tmp_frame3[,2:21]
test_frame <- as.data.frame(cbind(tmp_drame1, tmp_frame4, tmp_drame2))
colnames(test_frame)[1] <- "r_vector"
  


#----------------------
#HOUR RANKS - NON PRIME
#----------------------

tmp_frame1 <- test_frame[,1]
tmp_frame2 <- test_frame[,36:55]
tmp_frame3 <- as.matrix(cbind(tmp_frame1, tmp_frame2))
tmp_frame4 <- test_frame[,2:35]
  
col_vector <- c(2:6)
sort_vector <- c("IMP_R1", "R1R2_IMP", "R1R3_IMP", "R1R2_PCT", "R1R3_PCT")
sortcol_vector <- sapply(sort_vector, colname_no)

x <- 1
y <- nrow(geo_vector)

for (x in 1:y) {
  
  act_geo <- as.character(geo_vector[x,1])
  
  test_frame2 <- test_frame %>%
    filter(GEOGRAPHY == act_geo & DAYPART != 8)
  
  b <- 1
  h <- length(hour_vector)
  
  for (b in 1:h) {
    
    act_hour <- hour_vector[b]
    
    test_frame4 <- test_frame2 %>%
      filter(START_TIME_HH == act_hour)
    
    c <- 1
    j <- length(demo_vector)
    
    for (c in 1:j) {
      
      act_demo <- as.character(demo_vector[c])
      test_frame5 <- test_frame4 %>%
        filter(DEMO == act_demo)
      
      d <- 1
      k <- length(sort_vector)
      
      for (d in 1:k) {
        
        act_sort <- sort_vector[d]
        cn <- as.numeric(col_vector[d])
        srt_no <- as.numeric(sortcol_vector[d])
        
        test_frame3 <- test_frame5 %>%
          select(r_vector, !!act_sort)
        
        test_frame3 <- test_frame3[order(test_frame3[,2], decreasing = TRUE),]
        
        test_frame3 <- test_frame3 %>%
          select(r_vector)
        
        a <- 1
        g <- nrow(test_frame3)
        
        if (g > 0) { 
        for (a in 1:g) {
          rn <- test_frame3[a, 1]
          xx <- test_frame[rn, srt_no]
          if (is.na(xx) == FALSE) {
            tmp_frame3[rn, cn] <- a
            
          }
        }
        }
      }
    }
  }
}

tmp_drame1 <- tmp_frame3[,1]
tmp_drame2 <- tmp_frame3[,2:21]
test_frame <- as.data.frame(cbind(tmp_drame1, tmp_frame4, tmp_drame2))
colnames(test_frame)[1] <- "r_vector"


#----------------------
#HOUR RANKS - PRIME
#----------------------

tmp_frame1 <- test_frame[,1]
tmp_frame2 <- test_frame[,36:55]
tmp_frame3 <- as.matrix(cbind(tmp_frame1, tmp_frame2))
tmp_frame4 <- test_frame[,2:35]

col_vector <- c(2:6)
sort_vector <- c("IMP_R1", "R1R2_IMP", "R1R3_IMP", "R1R2_PCT", "R1R3_PCT")
sortcol_vector <- sapply(sort_vector, colname_no)

x <- 1
y <- nrow(geo_vector)

for (x in 1:y) {
  
  act_geo <- as.character(geo_vector[x,1])
  print(act_geo)
  
  test_frame2 <- test_frame %>%
    filter(GEOGRAPHY == act_geo & DAYPART == 8)
  
  b <- 17
  h <- length(hour_vector)
  
  for (b in b:h) {
    
    act_hour <- hour_vector[b]
    
    test_frame4 <- test_frame2 %>%
      filter(START_TIME_HH == act_hour)
    
    c <- 1
    j <- length(demo_vector)
    
    for (c in 1:j) {
      
      act_demo <- as.character(demo_vector[c])
      test_frame5 <- test_frame4 %>%
        filter(DEMO == act_demo)
      
      ee <- 1
      ll <- length(dow_vector)
      
      for (ee in 1:ll) {
        
        act_dow <- dow_vector[ee]
        test_frame6 <- test_frame5 %>%
          filter(DAY_OF_WEEK == act_dow)
      
        d <- 1
        k <- length(sort_vector)
        
        for (d in 1:k) {
        
          act_sort <- sort_vector[d]
          cn <- as.numeric(col_vector[d])
          srt_no <- as.numeric(sortcol_vector[d])
        
          test_frame3 <- test_frame6 %>%
            select(r_vector, !!act_sort)
        
          test_frame3 <- test_frame3[order(test_frame3[,2], decreasing = TRUE),]
        
          test_frame3 <- test_frame3 %>%
            select(r_vector)
        
          a <- 1
          g <- nrow(test_frame3)
        
          if (g > 0) { 
            for (a in 1:g) {
              rn <- test_frame3[a, 1]
              xx <- test_frame[rn, srt_no]
              if (is.na(xx) == FALSE) {
                tmp_frame3[rn, cn] <- a
              
              }
            }
          }
        }
      }
    }
  }
}

tmp_drame1 <- tmp_frame3[,1]
tmp_drame2 <- tmp_frame3[,2:21]
test_frame <- as.data.frame(cbind(tmp_drame1, tmp_frame4, tmp_drame2))
colnames(test_frame)[1] <- "r_vector"

#--------------------
#GENRE RANKS - NEWS
#--------------------

tmp_frame1 <- test_frame[,1]
tmp_frame2 <- test_frame[,36:55]
tmp_frame3 <- as.matrix(cbind(tmp_frame1, tmp_frame2))
tmp_frame4 <- test_frame[,2:35]

col_vector <- c(12:16)
sort_vector <- c("IMP_R1", "R1R2_IMP", "R1R3_IMP", "R1R2_PCT", "R1R3_PCT")
sortcol_vector <- sapply(sort_vector, colname_no)

x <- 1
y <- nrow(geo_vector)

for (x in 1:y) {
  
  act_geo <- as.character(geo_vector[x,1])
  
  test_frame2 <- test_frame %>%
    filter(GEOGRAPHY == act_geo)
  
  b <- 1
  h <- length(genre_vector)
  
  for (b in 1:h) {
    
    act_genre <- genre_vector[b]
    
    test_frame4 <- test_frame2 %>%
      filter(GENRE == act_genre)
    
    c <- 1
    j <- length(demo_vector)
    
    for (c in 1:j) {
      
      act_demo <- as.character(demo_vector[c])
      test_frame5 <- test_frame4 %>%
        filter(DEMO == act_demo)
      
      d <- 1
      k <- length(sort_vector)
      
      for (d in 1:k) {
        
        act_sort <- sort_vector[d]
        cn <- as.numeric(col_vector[d])
        srt_no <- as.numeric(sortcol_vector[d])
        
        test_frame3 <- test_frame5 %>%
          select(r_vector, !!act_sort)
        
        test_frame3 <- test_frame3[order(test_frame3[,2], decreasing = TRUE),]
        
        test_frame3 <- test_frame3 %>%
          select(r_vector)
        
        a <- 1
        g <- nrow(test_frame3)
        
        for (a in 1:g) {
          rn <- test_frame3[a, 1]
          xx <- test_frame[rn, srt_no]
          if (is.na(xx) == FALSE) {
            tmp_frame3[rn, cn] <- a
          }
        }
      }
    }
  }
}

tmp_drame1 <- tmp_frame3[,1]
tmp_drame2 <- tmp_frame3[,2:21]
test_frame <- as.data.frame(cbind(tmp_drame1, tmp_frame4, tmp_drame2))
colnames(test_frame)[1] <- "r_vector"


#------------------------------------------------
#DAY OF WEEK PRIME RANKS (PUT IN GENRE RANK COL)
#------------------------------------------------

tmp_frame1 <- test_frame[,1]
tmp_frame2 <- test_frame[,36:55]
tmp_frame3 <- as.matrix(cbind(tmp_frame1, tmp_frame2))
tmp_frame4 <- test_frame[,2:35]

col_vector <- c(12:16)
sort_vector <- c("IMP_R1", "R1R2_IMP", "R1R3_IMP", "R1R2_PCT", "R1R3_PCT")
sortcol_vector <- sapply(sort_vector, colname_no)

x <- 1
y <- nrow(geo_vector)

for (x in 1:y) {
  
  act_geo <- as.character(geo_vector[x,1])
  
  test_frame2 <- test_frame %>%
    filter(GEOGRAPHY == act_geo, DAYPART == 8)
  
  b <- 1
  h <- length(dow_vector)
  
  for (b in 1:h) {
    
    act_genre <- dow_vector[b]
    
    test_frame4 <- test_frame2 %>%
      filter(DAY_OF_WEEK == act_genre)
    
    c <- 1
    j <- length(demo_vector)
    
    for (c in 1:j) {
      
      act_demo <- as.character(demo_vector[c])
      test_frame5 <- test_frame4 %>%
        filter(DEMO == act_demo)
      
      d <- 1
      k <- length(sort_vector)
      
      for (d in 1:k) {
        
        act_sort <- sort_vector[d]
        cn <- as.numeric(col_vector[d])
        srt_no <- as.numeric(sortcol_vector[d])
        
        test_frame3 <- test_frame5 %>%
          select(r_vector, !!act_sort)
        
        test_frame3 <- test_frame3[order(test_frame3[,2], decreasing = TRUE),]
        
        test_frame3 <- test_frame3 %>%
          select(r_vector)
        
        a <- 1
        g <- nrow(test_frame3)
        
        for (a in 1:g) {
          rn <- test_frame3[a, 1]
          xx <- test_frame[rn, srt_no]
          if (is.na(xx) == FALSE) {
            tmp_frame3[rn, cn] <- a
          }
        }
      }
    }
  }
}

tmp_drame1 <- tmp_frame3[,1]
tmp_drame2 <- tmp_frame3[,2:21]
test_frame <- as.data.frame(cbind(tmp_drame1, tmp_frame4, tmp_drame2))
colnames(test_frame)[1] <- "r_vector"

#-----------------
#MONTH RANKS
#-----------------

tmp_frame1 <- test_frame[,1]
tmp_frame2 <- test_frame[,36:55]
tmp_frame3 <- as.matrix(cbind(tmp_frame1, tmp_frame2))
tmp_frame4 <- test_frame[,2:35]

col_vector <- c(17:21)
sort_vector <- c("IMP_R1", "R1R2_IMP", "R1R3_IMP", "R1R2_PCT", "R1R3_PCT")
sortcol_vector <- sapply(sort_vector, colname_no)

x <- 1
y <- nrow(geo_vector)

for (x in 1:y) {
  
  act_geo <- as.character(geo_vector[x,1])
  
  test_frame2 <- test_frame %>%
    filter(GEOGRAPHY == act_geo)
 
    c <- 1
    j <- length(demo_vector)
    
    for (c in 1:j) {
      
      act_demo <- as.character(demo_vector[c])
      test_frame5 <- test_frame2 %>%
        filter(DEMO == act_demo)
      
      d <- 1
      k <- length(sort_vector)
      
      for (d in 1:k) {
        
        act_sort <- sort_vector[d]
        cn <- as.numeric(col_vector[d])
        srt_no <- as.numeric(sortcol_vector[d])
        
        test_frame3 <- test_frame5 %>%
          select(r_vector, !!act_sort)
        
        test_frame3 <- test_frame3[order(test_frame3[,2], decreasing = TRUE),]
        
        test_frame3 <- test_frame3 %>%
          select(r_vector)
        
        a <- 1
        g <- nrow(test_frame3)
        
        for (a in 1:g) {
          rn <- test_frame3[a, 1]
          xx <- test_frame[rn, srt_no]
          if (is.na(xx) == FALSE) {
            tmp_frame3[rn, cn] <- a
          
        }
      }
    }
  }
}

tmp_drame1 <- tmp_frame3[,1]
tmp_drame2 <- tmp_frame3[,2:21]
test_frame <- as.data.frame(cbind(tmp_drame1, tmp_frame4, tmp_drame2))
colnames(test_frame)[1] <- "r_vector"


#--------------
#NEWS HIERARCHY
#--------------

#This is a hierarchy of success stories determined by local sales manager and marketing director input.

corr_desc <- function(x) {
  
  if (x == 20) { return ("VIEWERSHIP") }
  else if (x == 33) { return ("YOY VIEWER INC") }
  else if (x == 30) { return ("MOM VIEWER INC") }
  else if (x == 35) { return ("YOY % INC") }
  else if (x == 32) { return ("MOM % INC") }
  
}

sort_desc <- function(x) {
  
  return(colnames(htest_matrix)[x])
  
}

master_frame <- matrix(nrow = 1, ncol = 9)

aa <- 1
gg <- nrow(geo_vector)

for (aa in 1:gg) {

  act_geo <- as.character(geo_vector[aa, 1])
  print(act_geo)
  
  bb <- 1
  hh <- length(demo_vector)
  
  for (bb in 1:hh) {
    
    act_demo <- as.character(demo_vector[bb])
  
    htest_frame <- test_frame %>%
      filter(GEOGRAPHY == act_geo, SCRIPPS == 1, GENRE == "LN - LOCAL NEWS", DEMO == act_demo)

    htest_matrix <- as.matrix(htest_frame[, 36:55])
    htest_sort <- c(16, 17, 19, 18, 20, 11, 12, 14, 13, 15, 6, 7, 9, 8, 10, 1, 2, 4, 3, 5)
    htest_corr <- c(20, 30, 32, 33, 35, 20, 30, 32, 33, 35, 20, 30, 32, 33, 35, 20, 30, 32, 33, 35)
    htest_fill <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_type <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_rank <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_value <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_ct <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_hr <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)


    a <- 0
    o <- 0
    r <- 1
    t <- 10
    not_enough <- 1
    
    if (nrow(htest_frame) < 10) {
      print("NEWS LOWER THAN 10:")
      print(act_geo)
      t <- nrow(htest_frame)
      not_enough <- 0
      if (act_geo == "MIAMI-FT. LAUDERDALE") { o <- 1 }
    }
    
    if (o == 0) {
      
      print(a)

      while (a < t) {

        b <- 1
        h <- length(htest_sort)
  
        for (b in 1:h) {
    
          s <- htest_sort[b]
          corr <- htest_corr[b]
          htest_col <- htest_matrix[, s]
    
          c <- 1
          j <- length(htest_col)
    
          for (c in 1:j) {
      
            n <- htest_col[c]
            v <- as.numeric(htest_frame[c, corr])
            hhh <- as.numeric(htest_frame[c, 6])
            if (is.na(n) == TRUE) { n <- "Charles" }
            if (n == r & a < t & c %in% htest_fill == FALSE & v > .07) {
              a <- a + 1
              htest_fill[a] <- c
              htest_type[a] <- s
              htest_rank[a] <- r
              htest_value[a] <- v
              htest_ct[a] <- corr
              htest_hr[a] <- hhh
            }
          }
        }
        r <- r + 1
      }
      
      if (not_enough == 0 & o != 1) {
        
        a <- 10 - (10 - t)
        o <- 0
        r <- 1
        t <- 10
        not_enough <- 1
        
        if (o == 0) {
          
          while (a < t) {
            
            b <- 1
            h <- length(htest_sort)
            
            for (b in 1:h) {
              
              s <- htest_sort[b]
              corr <- htest_corr[b]
              htest_col <- htest_matrix[, s]
              
              c <- 1
              j <- length(htest_col)
              
              for (c in 1:j) {
                
                n <- htest_col[c]
                v <- as.numeric(htest_frame[c, corr])
                hhh <- as.numeric(htest_frame[c, 6])
                if (is.na(n) == TRUE) { n <- "Charles" }
                ne_switch <- which(htest_fill == c)
                ne_col <- htest_fill[ne_switch[1]]
                ne_type <- htest_type[ne_switch[1]]
                if (ne_col == c & ne_type == s) { orson_switch = 0 }
                else { orson_switch = 1 }
                if (n == r & a < t & orson_switch > 0 & v > .07) {
                  a <- a + 1
                  htest_fill[a] <- c
                  htest_type[a] <- s
                  htest_rank[a] <- r
                  htest_value[a] <- v
                  htest_ct[a] <- corr
                  htest_hr[a] <- hhh
                }
              }
            }
            r <- r + 1
          }
        }
      }

      a <- 1
      g <- length(htest_fill)
      htest_df <- matrix(nrow = 10, ncol = 9)

      for (a in 1:g) {
  
        rno <- htest_fill[a]
        corr <- htest_ct[a]
        hty <- htest_type[a]
        hhr <- htest_hr[a]
  
        prg <- htest_frame[rno, 3]
        val <- htest_value[a]
        dmo <- act_demo
        dsc <- corr_desc(corr)
        rnk <- htest_rank[a]
        typ <- "Local Newscasts"
        rkt <- sort_desc(hty)
  
        op <- c(act_geo, prg, dmo, typ, dsc, val, rnk, rkt, hhr)
        htest_df[a,] <- op
        htest_df <- htest_df[order(htest_df[,7]),]
      }
    
      master_frame <- rbind(master_frame, htest_df)
    }
  }
}


#---------------
#PRIME HIERARCHY
#---------------

master_prame <- matrix(nrow = 1, ncol = 10)

aa <- 1
gg <- nrow(geo_vector)

for (aa in 1:gg) {
  
  act_geo <- as.character(geo_vector[aa, 1])
  print(act_geo)
  
  bb <- 1
  hh <- length(demo_vector)
  
  for (bb in 1:hh) {
    
    act_demo <- as.character(demo_vector[bb])
    
    htest_frame <- test_frame %>%
      filter(GEOGRAPHY == act_geo, SCRIPPS == 1, DAYPART == 8, DEMO == act_demo)
    
    htest_matrix <- as.matrix(htest_frame[, 36:55])
    htest_sort <- c(16, 17, 19, 18, 20, 6, 7, 9, 8, 10, 11, 12, 14, 13, 15, 1, 2, 4, 3, 5)
    htest_corr <- c(20, 30, 32, 33, 35, 20, 30, 32, 33, 35, 20, 30, 32, 33, 35, 20, 30, 32, 33, 35)
    htest_fill <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_type <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_rank <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_value <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_ct <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_hr <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_dow <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    
    
    a <- 0
    o <- 0
    r <- 1
    t <- 10
    
    if (nrow(htest_frame) < 10) { 
      a <- 100 
      o <- 1
    }
    
    if (o == 0) {
      
      while (a < t) {
        
        b <- 1
        h <- length(htest_sort)
        
        for (b in 1:h) {
          
          s <- htest_sort[b]
          corr <- htest_corr[b]
          htest_col <- htest_matrix[, s]
          
          c <- 1
          j <- length(htest_col)
          
          for (c in 1:j) {
            
            n <- htest_col[c]
            v <- as.numeric(htest_frame[c, corr])
            dow <- as.character(htest_frame[c, 23])
            hhh <- as.numeric(htest_frame[c, 6])
            if (is.na(n) == TRUE) { n <- "Charles" }
            
            prime_switch <- 1
            if (s < 6 & r > 1) { prime_switch <- 0 }
            if (s >= 11 & s <= 15 & r > 9) { prime_switch <- 0 }
            
            if (n == r & a < t & c %in% htest_fill == FALSE & v > 0.07 & prime_switch > 0) {
              a <- a + 1
              htest_fill[a] <- c
              htest_type[a] <- s
              htest_rank[a] <- r
              htest_value[a] <- v
              htest_ct[a] <- corr
              htest_hr[a] <- hhh
              htest_dow[a] <- dow
            }
          }
        }
        r <- r + 1
      }
      
      a <- 1
      g <- length(htest_fill)
      htest_df <- matrix(nrow = 10, ncol = 10)
      
      for (a in 1:g) {
        
        rno <- htest_fill[a]
        corr <- htest_ct[a]
        hty <- htest_type[a]
        hhr <- htest_hr[a]
        
        prg <- htest_frame[rno, 3]
        val <- htest_value[a]
        dmo <- act_demo
        dsc <- corr_desc(corr)
        rnk <- htest_rank[a]
        typ <- "Prime"
        rkt <- sort_desc(hty)
        wk <- htest_dow[a]
        
        op <- c(act_geo, prg, dmo, typ, dsc, val, rnk, rkt, hhr, wk)
        htest_df[a,] <- op
      }
      
      master_prame <- rbind(master_prame, htest_df)
    }
  }
}

#---------------
#STRIP HIERARCHY
#---------------

master_srame <- matrix(nrow = 1, ncol = 9)

aa <- 1
gg <- nrow(geo_vector)

for (aa in 1:gg) {
  
  act_geo <- as.character(geo_vector[aa, 1])
  
  bb <- 1
  hh <- length(demo_vector)
  
  for (bb in 1:hh) {
    
    act_demo <- as.character(demo_vector[bb])
    
    htest_frame <- test_frame %>%
      filter(GEOGRAPHY == act_geo, SCRIPPS == 1, DAYPART != 8, GENRE != "LN - LOCAL NEWS", GENRE != "SE - SPORTS EVENT",
             DEMO == act_demo, GENRE != "SC - SPORTS COMMENTARY", GENRE != 	"S - SPORTS EVENTS & COMMENTARY",
             GENRE != "SA - SPORTS ANTHOLOGY", GENRE != "SN - SPORTS NEWS")
    
    htest_matrix <- as.matrix(htest_frame[, 36:55])
    htest_sort <- c(16, 17, 19, 18, 20, 6, 7, 9, 8, 10, 1, 2, 4, 3, 5)
    htest_corr <- c(20, 30, 32, 33, 35, 20, 30, 32, 33, 35, 20, 30, 32, 33, 35)
    htest_fill <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_type <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_rank <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_value <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_ct <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    htest_hr <- c(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    
    
    a <- 0
    o <- 0
    r <- 1
    t <- 10
    
    if (nrow(htest_frame) < 10) { 
      print("STRIP LOWER THAN 10:")
      print(act_geo)
      a <- 100 
      o <- 1
    }
    
    if (o == 0) {
      
      while (a < t) {
        
        b <- 1
        h <- length(htest_sort)
        
        for (b in 1:h) {
          
          s <- htest_sort[b]
          corr <- htest_corr[b]
          htest_col <- htest_matrix[, s]
          
          c <- 1
          j <- length(htest_col)
          
          for (c in 1:j) {
            
            n <- htest_col[c]
            v <- as.numeric(htest_frame[c, corr])
            hhh <- as.numeric(htest_frame[c, 6])
            if (is.na(n) == TRUE) { n <- "Charles" }
            if (n == r & a < t & c %in% htest_fill == FALSE & v > 0.07) {
              a <- a + 1
              htest_fill[a] <- c
              htest_type[a] <- s
              htest_rank[a] <- r
              htest_value[a] <- v
              htest_ct[a] <- corr
              htest_hr[a] <- hhh
            }
          }
        }
        r <- r + 1
      }
      
      a <- 1
      g <- length(htest_fill)
      htest_df <- matrix(nrow = 10, ncol = 9)
      
      for (a in 1:g) {
        
        rno <- htest_fill[a]
        corr <- htest_ct[a]
        hty <- htest_type[a]
        hhr <- htest_hr[a]
        
        prg <- htest_frame[rno, 3]
        val <- htest_value[a]
        dmo <- act_demo
        dsc <- corr_desc(corr)
        rnk <- htest_rank[a]
        typ <- "Strip"
        rkt <- sort_desc(hty)
        
        op <- c(act_geo, prg, dmo, typ, dsc, val, rnk, rkt, hhr)
        htest_df[a,] <- op
      }
      
      master_srame <- rbind(master_srame, htest_df)
    }
  }
}

master_prame <- master_prame[-1, ]
master_frame <- master_frame[-1, ]
master_srame <- master_srame[-1, ]
colnames(master_prame) <- c("GEO", "PROGRAM", "DEMO", "TYPE", "METRIC", "VALUE", "DMA RANK", "RANK TYPE", "HOUR", "DOW")
colnames(master_frame) <- c("GEO", "PROGRAM", "DEMO", "TYPE", "METRIC", "VALUE", "DMA RANK", "RANK TYPE", "HOUR")
colnames(master_srame) <- c("GEO", "PROGRAM", "DEMO", "TYPE", "METRIC", "VALUE", "DMA RANK", "RANK TYPE", "HOUR")

#------------
#EXCEL EXPORT
#------------

#Final Export Contains 4 tabs
#THEGRID is the structured all purpose file that has all needed data.
#NEWS 10 are the 10 biggest local news success stories of the report for each market.
#PRIME 10 are the 10 biggest prime programming success stories of the report for each market.
#STRIP 10 are the 10 biggest strip programming success stories of the report for each market.

wb <- createWorkbook()
addWorksheet(wb, "THE GRID")
addWorksheet(wb, "NEWS 10")
addWorksheet(wb, "PRIME 10")
addWorksheet(wb, "STRIP 10")
writeData(wb, sheet = "THE GRID", x = test_frame)
writeData(wb, sheet = "NEWS 10", x = master_frame)
writeData(wb, sheet = "PRIME 10", x = master_prame)
writeData(wb, sheet = "STRIP 10", x = master_srame)
saveWorkbook(wb, 'THEGRID.xlsx', overwrite = TRUE)
