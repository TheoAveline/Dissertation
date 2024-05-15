setwd("C:/Users/Theo/OneDrive - University College London/Politics/Year 3 (Dissertation)//")
library(haven)
library(tidyverse)
library(weights)
library(openxlsx)
library(survey)
data <- theo_data <- read_sav("C:/Users/Theo/Downloads/theo_data.sav")
#subsetting the data for complete cases
complete <- subset(data, w1_w2_w3_w4x == 1)

#making a derived variable 4 of sat16 only including conservative, labour, lib dem
complete$sat15_dv4 <- ifelse(complete$sat15_dv2 %in% c(1, 2, 3), complete$sat15_dv2, NA)
complete$a_sat16_dv4 <- ifelse(complete$a_sat16_dv2 %in% c(1, 2, 3), complete$a_sat16_dv2, NA)
complete$b_sat16_dv4 <- ifelse(complete$b_sat16_dv2 %in% c(1, 2, 3), complete$b_sat16_dv2, NA)
complete$c_sat16_dv4 <- ifelse(complete$c_sat16_dv2 %in% c(1, 2, 3), complete$c_sat16_dv2, NA)
complete$d_sat16_dv4 <- ifelse(complete$d_sat16_dv2 %in% c(1, 2, 3), complete$d_sat16_dv2, NA)

#making a derived variable 5 only including conservative, labour, lib dem and other
complete$sat15_dv5 <- ifelse(complete$sat15_dv2 %in% c(1, 2, 3, 999), complete$sat15_dv2, NA)
complete$a_sat16_dv5 <- ifelse(complete$a_sat16_dv2 %in% c(1, 2, 3, 999), complete$a_sat16_dv2, NA)
complete$b_sat16_dv5 <- ifelse(complete$b_sat16_dv2 %in% c(1, 2, 3, 999), complete$b_sat16_dv2, NA)
complete$c_sat16_dv5 <- ifelse(complete$c_sat16_dv2 %in% c(1, 2, 3, 999), complete$c_sat16_dv2, NA)
complete$d_sat16_dv5 <- ifelse(complete$d_sat16_dv2 %in% c(1, 2, 3, 999), complete$d_sat16_dv2, NA)


#all 5 waves voting variable tables in excel
#sat15 dv4 table in excel
weightsat15table <- wtd.table(complete$sat15_dv4, weights = complete$longitudinal_weight_1_w1_w2_w3_w4)
weightsat15tablesum <- sum(weightsat15table$sum.of.weights)
weightsat15table$percentage <- (weightsat15table$sum.of.weights / weightsat15tablesum) * 100
weightsat15tabledf <- as.data.frame(weightsat15table)
sat15table <- table(complete$sat15_dv4)
sat15tabledf <- as.data.frame(sat15table)
weightsat15tabledf <- data.frame(Frequency = sat15tabledf$Freq, Percentage = weightsat15tabledf$percentage)
wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Voting Tables dv4", x = weightsat15tabledf, startCol = 2, startRow = 3)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 1 dv4 table in excel
weightsat16atable <- wtd.table(complete$a_sat16_dv4, weights = complete$longitudinal_weight_1_w1_w2_w3_w4)
weightsat16atablesum <- sum(weightsat16atable$sum.of.weights)
weightsat16atable$percentage <- (weightsat16atable$sum.of.weights / weightsat16atablesum) * 100
weightsat16atabledf <- as.data.frame(weightsat16atable)
sat16atable <- table(complete$a_sat16_dv4)
sat16atabledf <- as.data.frame(sat16atable)
weightsat16atabledf <- data.frame(Frequency = sat16atabledf$Freq, Percentage = weightsat16atabledf$percentage)
wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Voting Tables dv4", x = weightsat16atabledf, startCol = 5, startRow = 3)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 2 dv4 table in excel
weightsat16btable <- wtd.table(complete$b_sat16_dv4, weights = complete$longitudinal_weight_1_w1_w2_w3_w4)
weightsat16btablesum <- sum(weightsat16btable$sum.of.weights)
weightsat16btable$percentage <- (weightsat16btable$sum.of.weights / weightsat16btablesum) * 100
weightsat16btabledf <- as.data.frame(weightsat16btable)
sat16btable <- table(complete$b_sat16_dv4)
sat16btabledf <- as.data.frame(sat16btable)
weightsat16btabledf <- data.frame(Frequency = sat16btabledf$Freq, Percentage = weightsat16btabledf$percentage)
wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Voting Tables dv4", x = weightsat16btabledf, startCol = 8, startRow = 3)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 3 dv4 table in excel
weightsat16ctable <- wtd.table(complete$c_sat16_dv4, weights = complete$longitudinal_weight_1_w1_w2_w3_w4)
weightsat16ctablesum <- sum(weightsat16ctable$sum.of.weights)
weightsat16ctable$percentage <- (weightsat16ctable$sum.of.weights / weightsat16ctablesum) * 100
weightsat16ctabledf <- as.data.frame(weightsat16ctable)
sat16ctable <- table(complete$c_sat16_dv4)
sat16ctabledf <- as.data.frame(sat16ctable)
weightsat16ctabledf <- data.frame(Frequency = sat16ctabledf$Freq, Percentage = weightsat16ctabledf$percentage)
wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Voting Tables dv4", x = weightsat16ctabledf, startCol = 11, startRow = 3)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 4 dv4 table in excel
weightsat16dtable <- wtd.table(complete$d_sat16_dv4, weights = complete$longitudinal_weight_1_w1_w2_w3_w4)
weightsat16dtablesum <- sum(weightsat16dtable$sum.of.weights)
weightsat16dtable$percentage <- (weightsat16dtable$sum.of.weights / weightsat16dtablesum) * 100
weightsat16dtabledf <- as.data.frame(weightsat16dtable)
sat16dtable <- table(complete$d_sat16_dv4)
sat16dtabledf <- as.data.frame(sat16dtable)
weightsat16dtabledf <- data.frame(Frequency = sat16dtabledf$Freq, Percentage = weightsat16dtabledf$percentage)
wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Voting Tables dv4", x = weightsat16dtabledf, startCol = 14, startRow = 3)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)


#doing the exact same thing but including others (using dv5) for comparison
#seeing whether including others has a significant impact on the graphs/tables
weightsat15table <- wtd.table(complete$sat15_dv5, weights = complete$longitudinal_weight_1_w1_w2_w3_w4)
weightsat15tablesum <- sum(weightsat15table$sum.of.weights)
weightsat15table$percentage <- (weightsat15table$sum.of.weights / weightsat15tablesum) * 100
weightsat15tabledf <- as.data.frame(weightsat15table)
sat15table <- table(complete$sat15_dv5)
sat15tabledf <- as.data.frame(sat15table)
weightsat15tabledf <- data.frame(Frequency = sat15tabledf$Freq, Percentage = weightsat15tabledf$percentage)
wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Voting Tables dv5 for Appendix", x = weightsat15tabledf, startCol = 2, startRow = 3)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 1 dv5 table in excel
weightsat16atable <- wtd.table(complete$a_sat16_dv5, weights = complete$longitudinal_weight_1_w1_w2_w3_w4)
weightsat16atablesum <- sum(weightsat16atable$sum.of.weights)
weightsat16atable$percentage <- (weightsat16atable$sum.of.weights / weightsat16atablesum) * 100
weightsat16atabledf <- as.data.frame(weightsat16atable)
sat16atable <- table(complete$a_sat16_dv5)
sat16atabledf <- as.data.frame(sat16atable)
weightsat16atabledf <- data.frame(Frequency = sat16atabledf$Freq, Percentage = weightsat16atabledf$percentage)
wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Voting Tables dv5 for Appendix", x = weightsat16atabledf, startCol = 5, startRow = 3)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 2 dv5 table in excel
weightsat16btable <- wtd.table(complete$b_sat16_dv5, weights = complete$longitudinal_weight_1_w1_w2_w3_w4)
weightsat16btablesum <- sum(weightsat16btable$sum.of.weights)
weightsat16btable$percentage <- (weightsat16btable$sum.of.weights / weightsat16btablesum) * 100
weightsat16btabledf <- as.data.frame(weightsat16btable)
sat16btable <- table(complete$b_sat16_dv5)
sat16btabledf <- as.data.frame(sat16btable)
weightsat16btabledf <- data.frame(Frequency = sat16btabledf$Freq, Percentage = weightsat16btabledf$percentage)
wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Voting Tables dv5 for Appendix", x = weightsat16btabledf, startCol = 8, startRow = 3)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 3 dv5 table in excel
weightsat16ctable <- wtd.table(complete$c_sat16_dv5, weights = complete$longitudinal_weight_1_w1_w2_w3_w4)
weightsat16ctablesum <- sum(weightsat16ctable$sum.of.weights)
weightsat16ctable$percentage <- (weightsat16ctable$sum.of.weights / weightsat16ctablesum) * 100
weightsat16ctabledf <- as.data.frame(weightsat16ctable)
sat16ctable <- table(complete$c_sat16_dv5)
sat16ctabledf <- as.data.frame(sat16ctable)
weightsat16ctabledf <- data.frame(Frequency = sat16ctabledf$Freq, Percentage = weightsat16ctabledf$percentage)
wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Voting Tables dv5 for Appendix", x = weightsat16ctabledf, startCol = 11, startRow = 3)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 4 dv5 table in excel
weightsat16dtable <- wtd.table(complete$d_sat16_dv5, weights = complete$longitudinal_weight_1_w1_w2_w3_w4)
weightsat16dtablesum <- sum(weightsat16dtable$sum.of.weights)
weightsat16dtable$percentage <- (weightsat16dtable$sum.of.weights / weightsat16dtablesum) * 100
weightsat16dtabledf <- as.data.frame(weightsat16dtable)
sat16dtable <- table(complete$d_sat16_dv5)
sat16dtabledf <- as.data.frame(sat16dtable)
weightsat16dtabledf <- data.frame(Frequency = sat16dtabledf$Freq, Percentage = weightsat16dtabledf$percentage)
wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Voting Tables dv5 for Appendix", x = weightsat16dtabledf, startCol = 14, startRow = 3)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

