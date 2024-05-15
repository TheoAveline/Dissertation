setwd("C:/Users/Theo/OneDrive - University College London/Politics/Year 3 (Dissertation)//")
library(haven)
library(tidyverse)
library(weights)
library(openxlsx)
library(survey)
library(mfx)
library(arm)
library(boot)
library(car)
library(pscl)
data <- theo_data <- read_sav("C:/Users/Theo/Downloads/theo_data.sav")

#subsetting the data for complete cases
complete <- subset(data, w1_w2_w3_w4x == 1)

#making a derived variable of sat16 only including conservative, labour, lib dem
complete$sat15_dv4 <- ifelse(complete$sat15_dv2 %in% c(1, 2, 3), complete$sat15_dv2, NA)
complete$a_sat16_dv4 <- ifelse(complete$a_sat16_dv2 %in% c(1, 2, 3), complete$a_sat16_dv2, NA)
complete$b_sat16_dv4 <- ifelse(complete$b_sat16_dv2 %in% c(1, 2, 3), complete$b_sat16_dv2, NA)
complete$c_sat16_dv4 <- ifelse(complete$c_sat16_dv2 %in% c(1, 2, 3), complete$c_sat16_dv2, NA)
complete$d_sat16_dv4 <- ifelse(complete$d_sat16_dv2 %in% c(1, 2, 3), complete$d_sat16_dv2, NA)

#making a new binary variable for regressions based on voting not labour in 2019
complete$vote_lib_con_19 <- complete$sat15_dv2
complete$vote_lib_con_19[complete$sat15_dv2 == 1 | complete$sat15_dv2 == 3] <- 1
complete$vote_lib_con_19[complete$sat15_dv2 == 2] <- 0

#making a new subset of the data for regressions just including those who vote con/lib dem 2019
data_vote_lib_con_19 <- complete[complete$vote_lib_con_19 == 1, ]

#making binary variables for those who switch to labour in each wave 
data_vote_lib_con_19$a_switch_lab <- ifelse(data_vote_lib_con_19$a_sat16_dv2 %in% c(1, 3), 0, ifelse(data_vote_lib_con_19$a_sat16_dv4 == 2, 1, NA))
data_vote_lib_con_19$b_switch_lab <- ifelse(data_vote_lib_con_19$b_sat16_dv2 %in% c(1, 3), 0, ifelse(data_vote_lib_con_19$b_sat16_dv4 == 2, 1, NA))
data_vote_lib_con_19$c_switch_lab <- ifelse(data_vote_lib_con_19$c_sat16_dv2 %in% c(1, 3), 0, ifelse(data_vote_lib_con_19$c_sat16_dv4 == 2, 1, NA))
data_vote_lib_con_19$d_switch_lab <- ifelse(data_vote_lib_con_19$d_sat16_dv2 %in% c(1, 3), 0, ifelse(data_vote_lib_con_19$d_sat16_dv4 == 2, 1, NA))


#doing the same again but including others to see if this affects regression results 
#making a new binary variable for regressions based on voting not labour in 2019
complete$vote_against_lab_19 <- complete$sat15_dv2
complete$vote_against_lab_19[complete$sat15_dv2 == 1 | complete$sat15_dv2 == 3| complete$sat15_dv2 == 999] <- 1
complete$vote_against_lab_19[complete$sat15_dv2 == 2] <- 0

#making a new subset of the data for regressions just including those who vote con/lib dem 2019
data_vote_against_lab_19 <- complete[complete$vote_against_lab_19 == 1, ]

#making binary variables for those who switch to labour in each wave 
data_vote_against_lab_19$a_switch_lab <- ifelse(data_vote_against_lab_19$a_sat16_dv2 %in% c(1, 3, 999), 0, ifelse(data_vote_against_lab_19$a_sat16_dv4 == 2, 1, NA))
data_vote_against_lab_19$b_switch_lab <- ifelse(data_vote_against_lab_19$b_sat16_dv2 %in% c(1, 3, 999), 0, ifelse(data_vote_against_lab_19$b_sat16_dv4 == 2, 1, NA))
data_vote_against_lab_19$c_switch_lab <- ifelse(data_vote_against_lab_19$c_sat16_dv2 %in% c(1, 3, 999), 0, ifelse(data_vote_against_lab_19$c_sat16_dv4 == 2, 1, NA))
data_vote_against_lab_19$d_switch_lab <- ifelse(data_vote_against_lab_19$d_sat16_dv2 %in% c(1, 3, 999), 0, ifelse(data_vote_against_lab_19$d_sat16_dv4 == 2, 1, NA))


#checking tables for voting variables
table(complete$sat15_dv2)
table(complete$a_sat16_dv2)
table(complete$b_sat16_dv2)
table(complete$c_sat16_dv2)
table(complete$d_sat16_dv2)

table(complete$sat15_dv)
table(complete$sat15_dv2)

#checking NAs for voting variables
sum(is.na(complete$sat15_dv2))
sum(is.na(complete$a_sat16_dv2))
sum(is.na(complete$b_sat16_dv2))
sum(is.na(complete$c_sat16_dv2))
sum(is.na(complete$d_sat16_dv2))

#making a new subset for regression analysis - only labour, con, lib dem, other for each vote variable
ultra_complete <- complete[
  complete$sat15_dv2 %in% c(1, 2, 3, 999) &
    complete$a_sat16_dv2 %in% c(1, 2, 3, 999) &
    complete$b_sat16_dv2 %in% c(1, 2, 3, 999) &
    complete$c_sat16_dv2 %in% c(1, 2, 3, 999) &
    complete$d_sat16_dv2 %in% c(1, 2, 3, 999), 
]

#checking it worked - it did 
sum(is.na(ultra_complete$sat15_dv2))
sum(is.na(ultra_complete$a_sat16_dv2))
sum(is.na(ultra_complete$b_sat16_dv2))
sum(is.na(ultra_complete$c_sat16_dv2))
sum(is.na(ultra_complete$d_sat16_dv2))

#checking voting variables for ultra complete subset - tells a slightly different narrative
#perhaps conservative voters are less likely to put undecided down
table(ultra_complete$sat15_dv2)
table(ultra_complete$a_sat16_dv2)
table(ultra_complete$b_sat16_dv2)
table(ultra_complete$c_sat16_dv2)
table(ultra_complete$d_sat16_dv2)


#doing a test regression of 2019 election voting on demographic variables
testmodel <- lm(sat15_dv2 ~ dem3_dv + dem7_1_dv + age_4_dv + denomination_dv3, data = ultra_complete)
summary(testmodel)

#making a new variable for voting conservative in the 2019 election
data$vote_con_19 <- ifelse(data$sat15_dv2 == 1, 1, 0)
complete$vote_con_19 <- ifelse(complete$sat15_dv2 == 1, 1, 0)
ultra_complete$vote_con_19 <- ifelse(ultra_complete$sat15_dv2 == 1, 1, 0)

#doing a test logistic regression on this new binary variable
logitmodel <- glm(vote_con_19 ~ as.factor(dem3_dv) + as.factor(dem7_1_dv) + as.factor(age_4_dv), data = ultra_complete, family = "binomial")
summary(logitmodel)

#running a test logistic regression and trying to include the weights
logitmodel2 <- glm(vote_con_19 ~ as.factor(dem3_dv) + as.factor(dem7_1_dv) + as.factor(age_4_dv) + as.factor(region_dv2) + as.factor(denomination_dv3), data = ultra_complete, family = "binomial", weights = longitudinal_weight_1_w1_w2_w3_w4)
summary(logitmodel2)

#doing another logistic regression using complete instead of ultra complete
logitmodel3 <- glm(vote_con_19 ~ as.factor(dem3_dv) + as.factor(dem7_1_dv) + as.factor(age_4_dv) + as.factor(region_dv2) + as.factor(denomination_dv3), data = complete, family = "binomial", weights = longitudinal_weight_1_w1_w2_w3_w4)
summary(logitmodel3)


#pre-regression changing the reference category of denomination to strictly orthodox
data_vote_against_lab_19$denomination_dv3 <- factor(data_vote_against_lab_19$denomination_dv3)
data_vote_against_lab_19$denomination_dv3 <- relevel(data_vote_against_lab_19$denomination_dv3, ref = "1")

#pre-regression changing the reference category of region to north-west
data_vote_against_lab_19$region_dv2 <- factor(data_vote_against_lab_19$region_dv2)
data_vote_against_lab_19$region_dv2 <- relevel(data_vote_against_lab_19$region_dv2, ref = "2")

#trying a regression on wave 1 
logitmodel_aswitch <- glm(a_switch_lab ~ as.factor(dem7_1_dv) + as.factor(dem3_dv) + as.factor(region_dv2) + as.factor(age_4_dv) + as.factor(denomination_dv3), data = data_vote_lib_con_19, family = "binomial", weights = longitudinal_weight_1_w1_w2_w3_w4)
summary(logitmodel_aswitch)
logitmfx(logitmodel_aswitch,atmean=F,data=data_vote_lib_con_19)

#repeating for waves 2,3 and 4
logitmodel_bswitch <- glm(b_switch_lab ~ as.factor(dem7_1_dv) + as.factor(dem3_dv) + as.factor(region_dv2) + as.factor(age_4_dv) + as.factor(denomination_dv3), data = data_vote_lib_con_19, family = "binomial", weights = longitudinal_weight_1_w1_w2_w3_w4)
summary(logitmodel_bswitch)
logitmfx(logitmodel_bswitch,atmean=F,data=data_vote_lib_con_19)

logitmodel_cswitch <- glm(c_switch_lab ~ as.factor(dem7_1_dv) + as.factor(dem3_dv) + as.factor(region_dv2) + as.factor(age_4_dv) + as.factor(denomination_dv3), data = data_vote_lib_con_19, family = "binomial", weights = longitudinal_weight_1_w1_w2_w3_w4)
summary(logitmodel_cswitch)
logitmfx(logitmodel_cswitch,atmean=F,data=data_vote_lib_con_19)

logitmodel_dswitch <- glm(d_switch_lab ~ as.factor(dem7_1_dv) + as.factor(dem3_dv) + as.factor(region_dv2) + as.factor(age_4_dv) + as.factor(denomination_dv3), data = data_vote_lib_con_19, family = "binomial", weights = longitudinal_weight_1_w1_w2_w3_w4)
summary(logitmodel_dswitch)
logitmfx(logitmodel_dswitch,atmean=F,data=data_vote_lib_con_19)

#running the regressions again including others
logitmodel_aswitch2 <- glm(a_switch_lab ~ as.factor(dem7_1_dv) + as.factor(dem3_dv) + as.factor(region_dv2) + as.factor(age_4_dv) + as.factor(denomination_dv3), data = data_vote_against_lab_19, family = "binomial", weights = longitudinal_weight_1_w1_w2_w3_w4)
summary(logitmodel_aswitch2)
logitmfx(logitmodel_aswitch2,atmean=F,data=data_vote_against_lab_19)
vif(logitmodel_aswitch2)


logitmodel_bswitch2 <- glm(b_switch_lab ~ as.factor(dem7_1_dv) + as.factor(dem3_dv) + as.factor(region_dv2) + as.factor(age_4_dv) + as.factor(denomination_dv3), data = data_vote_against_lab_19, family = "binomial", weights = longitudinal_weight_1_w1_w2_w3_w4)
summary(logitmodel_bswitch2)
logitmfx(logitmodel_bswitch2,atmean=F,data=data_vote_against_lab_19)
vif(logitmodel_bswitch2)


logitmodel_cswitch2 <- glm(c_switch_lab ~ as.factor(dem7_1_dv) + as.factor(dem3_dv) + as.factor(region_dv2) + as.factor(age_4_dv) + as.factor(denomination_dv3), data = data_vote_against_lab_19, family = "binomial", weights = longitudinal_weight_1_w1_w2_w3_w4)
summary(logitmodel_cswitch2)
logitmfx(logitmodel_cswitch2,atmean=F,data=data_vote_against_lab_19)
vif(logitmodel_cswitch2)


logitmodel_dswitch2 <- glm(d_switch_lab ~ as.factor(dem7_1_dv) + as.factor(dem3_dv) + as.factor(region_dv2) + as.factor(age_4_dv) + as.factor(denomination_dv3), data = data_vote_against_lab_19, family = "binomial", weights = longitudinal_weight_1_w1_w2_w3_w4)
summary(logitmodel_dswitch2)
logitmfx(logitmodel_dswitch2,atmean=F,data=data_vote_against_lab_19)
vif(logitmodel_dswitch2)


#summary stats for diss appendix
table(complete$dem7_1_dv)
table(complete$dem3_dv)
table(complete$region_dv2)
table(complete$age_4_dv)
table(complete$denomination_dv3)

table(data_vote_against_lab_19$dem7_1_dv)
table(data_vote_against_lab_19$dem3_dv)
table(data_vote_against_lab_19$region_dv2)
table(data_vote_against_lab_19$age_4_dv)
table(data_vote_against_lab_19$denomination_dv3)
