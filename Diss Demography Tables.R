setwd("C:/Users/Theo/OneDrive - University College London/Politics/Year 3 (Dissertation)//")
library(haven)
library(tidyverse)
library(weights)
library(openxlsx)
library(survey)
data <- theo_data <- read_sav("C:/Users/Theo/Downloads/theo_data.sav")
#subsetting the data for complete cases
complete <- subset(data, w1_w2_w3_w4x == 1)

#making a derived variable of sat16 only including conservative, labour, lib dem
complete$sat15_dv4 <- ifelse(complete$sat15_dv2 %in% c(1, 2, 3), complete$sat15_dv2, NA)
complete$a_sat16_dv4 <- ifelse(complete$a_sat16_dv2 %in% c(1, 2, 3), complete$a_sat16_dv2, NA)
complete$b_sat16_dv4 <- ifelse(complete$b_sat16_dv2 %in% c(1, 2, 3), complete$b_sat16_dv2, NA)
complete$c_sat16_dv4 <- ifelse(complete$c_sat16_dv2 %in% c(1, 2, 3), complete$c_sat16_dv2, NA)
complete$d_sat16_dv4 <- ifelse(complete$d_sat16_dv2 %in% c(1, 2, 3), complete$d_sat16_dv2, NA)


#making demography dv4 tables in excel
#education dv4 table
#2019 election
weightsat15education1table1 <- wtd.table(complete$sat15_dv4[complete$dem7_1_dv == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem7_1_dv == 1])
weightsat15education1sum <- sum(weightsat15education1table1$sum.of.weights)
weightsat15education1table1$percentage <- (weightsat15education1table1$sum.of.weights / weightsat15education1sum) * 100
weightsat15education1table1df <- as.data.frame(weightsat15education1table1)
sat15education1table <- table(complete$sat15_dv4[complete$dem7_1_dv == 1])
sat15education1tabledf <- as.data.frame(sat15education1table)
sat15education1sum <- sum(sat15education1tabledf$Freq)
sat15education1tabledf$Percentage <- (sat15education1tabledf$Freq / sat15education1sum) * 100
weightsat15education1tabledf <- data.frame(Party = weightsat15education1table1$x, Frequency = sat15education1tabledf$Freq, Percentage_Weighted = weightsat15education1table1$percentage, Percentage_Unweighted = sat15education1tabledf$Percentage)

weightsat15education2table1 <- wtd.table(complete$sat15_dv4[complete$dem7_1_dv == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem7_1_dv == 2])
weightsat15education2sum <- sum(weightsat15education2table1$sum.of.weights)
weightsat15education2table1$percentage <- (weightsat15education2table1$sum.of.weights / weightsat15education2sum) * 100
weightsat15education2table1df <- as.data.frame(weightsat15education2table1)
sat15education2table <- table(complete$sat15_dv4[complete$dem7_1_dv == 2])
sat15education2tabledf <- as.data.frame(sat15education2table)
sat15education2sum <- sum(sat15education2tabledf$Freq)
sat15education2tabledf$Percentage <- (sat15education2tabledf$Freq / sat15education2sum) * 100
weightsat15education2tabledf <- data.frame(Frequency = sat15education2tabledf$Freq, Percentage_Weighted = weightsat15education2table1$percentage, Percentage_Unweighted = sat15education2tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Education dv4 Tables", x = weightsat15education1tabledf, startCol = 1, startRow = 4)
writeData(wb, sheet = "Education dv4 Tables", x = weightsat15education2tabledf, startCol = 6, startRow = 4)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 1
weightsat16aeducation1table1 <- wtd.table(complete$a_sat16_dv4[complete$dem7_1_dv == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem7_1_dv == 1])
weightsat16aeducation1sum <- sum(weightsat16aeducation1table1$sum.of.weights)
weightsat16aeducation1table1$percentage <- (weightsat16aeducation1table1$sum.of.weights / weightsat16aeducation1sum) * 100
weightsat16aeducation1table1df <- as.data.frame(weightsat16aeducation1table1)
sat16aeducation1table <- table(complete$a_sat16_dv4[complete$dem7_1_dv == 1])
sat16aeducation1tabledf <- as.data.frame(sat16aeducation1table)
sat16aeducation1sum <- sum(sat16aeducation1tabledf$Freq)
sat16aeducation1tabledf$Percentage <- (sat16aeducation1tabledf$Freq / sat16aeducation1sum) * 100
weightsat16aeducation1tabledf <- data.frame(Party = weightsat16aeducation1table1$x, Frequency = sat16aeducation1tabledf$Freq, Percentage_Weighted = weightsat16aeducation1table1$percentage, Percentage_Unweighted = sat16aeducation1tabledf$Percentage)

weightsat16aeducation2table1 <- wtd.table(complete$a_sat16_dv4[complete$dem7_1_dv == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem7_1_dv == 2])
weightsat16aeducation2sum <- sum(weightsat16aeducation2table1$sum.of.weights)
weightsat16aeducation2table1$percentage <- (weightsat16aeducation2table1$sum.of.weights / weightsat16aeducation2sum) * 100
weightsat16aeducation2table1df <- as.data.frame(weightsat16aeducation2table1)
sat16aeducation2table <- table(complete$a_sat16_dv4[complete$dem7_1_dv == 2])
sat16aeducation2tabledf <- as.data.frame(sat16aeducation2table)
sat16aeducation2sum <- sum(sat16aeducation2tabledf$Freq)

sat16aeducation2tabledf$Percentage <- (sat16aeducation2tabledf$Freq / sat16aeducation2sum) * 100
weightsat16aeducation2tabledf <- data.frame(Frequency = sat16aeducation2tabledf$Freq, Percentage_Weighted = weightsat16aeducation2table1$percentage, Percentage_Unweighted = sat16aeducation2tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Education dv4 Tables", x = weightsat16aeducation1tabledf, startCol = 1, startRow = 14)
writeData(wb, sheet = "Education dv4 Tables", x = weightsat16aeducation2tabledf, startCol = 6, startRow = 14)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 2
weightsat16beducation1table1 <- wtd.table(complete$b_sat16_dv4[complete$dem7_1_dv == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem7_1_dv == 1])
weightsat16beducation1sum <- sum(weightsat16beducation1table1$sum.of.weights)
weightsat16beducation1table1$percentage <- (weightsat16beducation1table1$sum.of.weights / weightsat16beducation1sum) * 100
weightsat16beducation1table1df <- as.data.frame(weightsat16beducation1table1)
sat16beducation1table <- table(complete$b_sat16_dv4[complete$dem7_1_dv == 1])
sat16beducation1tabledf <- as.data.frame(sat16beducation1table)
sat16beducation1sum <- sum(sat16beducation1tabledf$Freq)
sat16beducation1tabledf$Percentage <- (sat16beducation1tabledf$Freq / sat16beducation1sum) * 100
weightsat16beducation1tabledf <- data.frame(Party = weightsat16beducation1table1$x, Frequency = sat16beducation1tabledf$Freq, Percentage_Weighted = weightsat16beducation1table1$percentage, Percentage_Unweighted = sat16beducation1tabledf$Percentage)

weightsat16beducation2table1 <- wtd.table(complete$b_sat16_dv4[complete$dem7_1_dv == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem7_1_dv == 2])
weightsat16beducation2sum <- sum(weightsat16beducation2table1$sum.of.weights)
weightsat16beducation2table1$percentage <- (weightsat16beducation2table1$sum.of.weights / weightsat16beducation2sum) * 100
weightsat16beducation2table1df <- as.data.frame(weightsat16beducation2table1)
sat16beducation2table <- table(complete$b_sat16_dv4[complete$dem7_1_dv == 2])
sat16beducation2tabledf <- as.data.frame(sat16beducation2table)
sat16beducation2sum <- sum(sat16beducation2tabledf$Freq)
sat16beducation2tabledf$Percentage <- (sat16beducation2tabledf$Freq / sat16beducation2sum) * 100
weightsat16beducation2tabledf <- data.frame(Frequency = sat16beducation2tabledf$Freq, Percentage_Weighted = weightsat16beducation2table1$percentage, Percentage_Unweighted = sat16beducation2tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Education dv4 Tables", x = weightsat16beducation1tabledf, startCol = 1, startRow = 24)
writeData(wb, sheet = "Education dv4 Tables", x = weightsat16beducation2tabledf, startCol = 6, startRow = 24)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 3
weightsat16ceducation1table1 <- wtd.table(complete$c_sat16_dv4[complete$dem7_1_dv == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem7_1_dv == 1])
weightsat16ceducation1sum <- sum(weightsat16ceducation1table1$sum.of.weights)
weightsat16ceducation1table1$percentage <- (weightsat16ceducation1table1$sum.of.weights / weightsat16ceducation1sum) * 100
weightsat16ceducation1table1df <- as.data.frame(weightsat16ceducation1table1)
sat16ceducation1table <- table(complete$c_sat16_dv4[complete$dem7_1_dv == 1])
sat16ceducation1tabledf <- as.data.frame(sat16ceducation1table)
sat16ceducation1sum <- sum(sat16ceducation1tabledf$Freq)
sat16ceducation1tabledf$Percentage <- (sat16ceducation1tabledf$Freq / sat16ceducation1sum) * 100
weightsat16ceducation1tabledf <- data.frame(Party = weightsat16ceducation1table1$x, Frequency = sat16ceducation1tabledf$Freq, Percentage_Weighted = weightsat16ceducation1table1$percentage, Percentage_Unweighted = sat16ceducation1tabledf$Percentage)

weightsat16ceducation2table1 <- wtd.table(complete$c_sat16_dv4[complete$dem7_1_dv == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem7_1_dv == 2])
weightsat16ceducation2sum <- sum(weightsat16ceducation2table1$sum.of.weights)
weightsat16ceducation2table1$percentage <- (weightsat16ceducation2table1$sum.of.weights / weightsat16ceducation2sum) * 100
weightsat16ceducation2table1df <- as.data.frame(weightsat16ceducation2table1)
sat16ceducation2table <- table(complete$c_sat16_dv4[complete$dem7_1_dv == 2])
sat16ceducation2tabledf <- as.data.frame(sat16ceducation2table)
sat16ceducation2sum <- sum(sat16ceducation2tabledf$Freq)
sat16ceducation2tabledf$Percentage <- (sat16ceducation2tabledf$Freq / sat16ceducation2sum) * 100
weightsat16ceducation2tabledf <- data.frame(Frequency = sat16ceducation2tabledf$Freq, Percentage_Weighted = weightsat16ceducation2table1$percentage, Percentage_Unweighted = sat16ceducation2tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Education dv4 Tables", x = weightsat16ceducation1tabledf, startCol = 1, startRow = 34)
writeData(wb, sheet = "Education dv4 Tables", x = weightsat16ceducation2tabledf, startCol = 6, startRow = 34)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 4
weightsat16deducation1table1 <- wtd.table(complete$d_sat16_dv4[complete$dem7_1_dv == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem7_1_dv == 1])
weightsat16deducation1sum <- sum(weightsat16deducation1table1$sum.of.weights)
weightsat16deducation1table1$percentage <- (weightsat16deducation1table1$sum.of.weights / weightsat16deducation1sum) * 100
weightsat16deducation1table1df <- as.data.frame(weightsat16deducation1table1)
sat16deducation1table <- table(complete$d_sat16_dv4[complete$dem7_1_dv == 1])
sat16deducation1tabledf <- as.data.frame(sat16deducation1table)
sat16deducation1sum <- sum(sat16deducation1tabledf$Freq)
sat16deducation1tabledf$Percentage <- (sat16deducation1tabledf$Freq / sat16deducation1sum) * 100
weightsat16deducation1tabledf <- data.frame(Party = weightsat16deducation1table1$x, Frequency = sat16deducation1tabledf$Freq, Percentage_Weighted = weightsat16deducation1table1$percentage, Percentage_Unweighted = sat16deducation1tabledf$Percentage)

weightsat16deducation2table1 <- wtd.table(complete$d_sat16_dv4[complete$dem7_1_dv == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem7_1_dv == 2])
weightsat16deducation2sum <- sum(weightsat16deducation2table1$sum.of.weights)
weightsat16deducation2table1$percentage <- (weightsat16deducation2table1$sum.of.weights / weightsat16deducation2sum) * 100
weightsat16deducation2table1df <- as.data.frame(weightsat16deducation2table1)
sat16deducation2table <- table(complete$d_sat16_dv4[complete$dem7_1_dv == 2])
sat16deducation2tabledf <- as.data.frame(sat16deducation2table)
sat16deducation2sum <- sum(sat16deducation2tabledf$Freq)
sat16deducation2tabledf$Percentage <- (sat16deducation2tabledf$Freq / sat16deducation2sum) * 100
weightsat16deducation2tabledf <- data.frame(Frequency = sat16deducation2tabledf$Freq, Percentage_Weighted = weightsat16deducation2table1$percentage, Percentage_Unweighted = sat16deducation2tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Education dv4 Tables", x = weightsat16deducation1tabledf, startCol = 1, startRow = 44)
writeData(wb, sheet = "Education dv4 Tables", x = weightsat16deducation2tabledf, startCol = 6, startRow = 44)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#sex dv4 table
#2019 election
weightsat15sex1table1 <- wtd.table(complete$sat15_dv4[complete$dem3_dv == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem3_dv == 1])
weightsat15sex1sum <- sum(weightsat15sex1table1$sum.of.weights)
weightsat15sex1table1$percentage <- (weightsat15sex1table1$sum.of.weights / weightsat15sex1sum) * 100
weightsat15sex1table1df <- as.data.frame(weightsat15sex1table1)
sat15sex1table <- table(complete$sat15_dv4[complete$dem3_dv == 1])
sat15sex1tabledf <- as.data.frame(sat15sex1table)
sat15sex1sum <- sum(sat15sex1tabledf$Freq)
sat15sex1tabledf$Percentage <- (sat15sex1tabledf$Freq / sat15sex1sum) * 100
weightsat15sex1tabledf <- data.frame(Party = weightsat15sex1table1$x, Frequency = sat15sex1tabledf$Freq, Percentage_Weighted = weightsat15sex1table1$percentage, Percentage_Unweighted = sat15sex1tabledf$Percentage)

weightsat15sex2table1 <- wtd.table(complete$sat15_dv4[complete$dem3_dv == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem3_dv == 2])
weightsat15sex2sum <- sum(weightsat15sex2table1$sum.of.weights)
weightsat15sex2table1$percentage <- (weightsat15sex2table1$sum.of.weights / weightsat15sex2sum) * 100
weightsat15sex2table1df <- as.data.frame(weightsat15sex2table1)
sat15sex2table <- table(complete$sat15_dv4[complete$dem3_dv == 2])
sat15sex2tabledf <- as.data.frame(sat15sex2table)
sat15sex2sum <- sum(sat15sex2tabledf$Freq)
sat15sex2tabledf$Percentage <- (sat15sex2tabledf$Freq / sat15sex2sum) * 100
weightsat15sex2tabledf <- data.frame(Frequency = sat15sex2tabledf$Freq, Percentage_Weighted = weightsat15sex2table1$percentage, Percentage_Unweighted = sat15sex2tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Sex dv4 Tables", x = weightsat15sex1tabledf, startCol = 1, startRow = 4)
writeData(wb, sheet = "Sex dv4 Tables", x = weightsat15sex2tabledf, startCol = 6, startRow = 4)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 1
weightsat16asex1table1 <- wtd.table(complete$a_sat16_dv4[complete$dem3_dv == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem3_dv == 1])
weightsat16asex1sum <- sum(weightsat16asex1table1$sum.of.weights)
weightsat16asex1table1$percentage <- (weightsat16asex1table1$sum.of.weights / weightsat16asex1sum) * 100
weightsat16asex1table1df <- as.data.frame(weightsat16asex1table1)
sat16asex1table <- table(complete$a_sat16_dv4[complete$dem3_dv == 1])
sat16asex1tabledf <- as.data.frame(sat16asex1table)
sat16asex1sum <- sum(sat16asex1tabledf$Freq)
sat16asex1tabledf$Percentage <- (sat16asex1tabledf$Freq / sat16asex1sum) * 100
weightsat16asex1tabledf <- data.frame(Party = weightsat16asex1table1$x, Frequency = sat16asex1tabledf$Freq, Percentage_Weighted = weightsat16asex1table1$percentage, Percentage_Unweighted = sat16asex1tabledf$Percentage)

weightsat16asex2table1 <- wtd.table(complete$a_sat16_dv4[complete$dem3_dv == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem3_dv == 2])
weightsat16asex2sum <- sum(weightsat16asex2table1$sum.of.weights)
weightsat16asex2table1$percentage <- (weightsat16asex2table1$sum.of.weights / weightsat16asex2sum) * 100
weightsat16asex2table1df <- as.data.frame(weightsat16asex2table1)
sat16asex2table <- table(complete$a_sat16_dv4[complete$dem3_dv == 2])
sat16asex2tabledf <- as.data.frame(sat16asex2table)
sat16asex2sum <- sum(sat16asex2tabledf$Freq)

sat16asex2tabledf$Percentage <- (sat16asex2tabledf$Freq / sat16asex2sum) * 100
weightsat16asex2tabledf <- data.frame(Frequency = sat16asex2tabledf$Freq, Percentage_Weighted = weightsat16asex2table1$percentage, Percentage_Unweighted = sat16asex2tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Sex dv4 Tables", x = weightsat16asex1tabledf, startCol = 1, startRow = 14)
writeData(wb, sheet = "Sex dv4 Tables", x = weightsat16asex2tabledf, startCol = 6, startRow = 14)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 2
weightsat16bsex1table1 <- wtd.table(complete$b_sat16_dv4[complete$dem3_dv == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem3_dv == 1])
weightsat16bsex1sum <- sum(weightsat16bsex1table1$sum.of.weights)
weightsat16bsex1table1$percentage <- (weightsat16bsex1table1$sum.of.weights / weightsat16bsex1sum) * 100
weightsat16bsex1table1df <- as.data.frame(weightsat16bsex1table1)
sat16bsex1table <- table(complete$b_sat16_dv4[complete$dem3_dv == 1])
sat16bsex1tabledf <- as.data.frame(sat16bsex1table)
sat16bsex1sum <- sum(sat16bsex1tabledf$Freq)
sat16bsex1tabledf$Percentage <- (sat16bsex1tabledf$Freq / sat16bsex1sum) * 100
weightsat16bsex1tabledf <- data.frame(Party = weightsat16bsex1table1$x, Frequency = sat16bsex1tabledf$Freq, Percentage_Weighted = weightsat16bsex1table1$percentage, Percentage_Unweighted = sat16bsex1tabledf$Percentage)

weightsat16bsex2table1 <- wtd.table(complete$b_sat16_dv4[complete$dem3_dv == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem3_dv == 2])
weightsat16bsex2sum <- sum(weightsat16bsex2table1$sum.of.weights)
weightsat16bsex2table1$percentage <- (weightsat16bsex2table1$sum.of.weights / weightsat16bsex2sum) * 100
weightsat16bsex2table1df <- as.data.frame(weightsat16bsex2table1)
sat16bsex2table <- table(complete$b_sat16_dv4[complete$dem3_dv == 2])
sat16bsex2tabledf <- as.data.frame(sat16bsex2table)
sat16bsex2sum <- sum(sat16bsex2tabledf$Freq)
sat16bsex2tabledf$Percentage <- (sat16bsex2tabledf$Freq / sat16bsex2sum) * 100
weightsat16bsex2tabledf <- data.frame(Frequency = sat16bsex2tabledf$Freq, Percentage_Weighted = weightsat16bsex2table1$percentage, Percentage_Unweighted = sat16bsex2tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Sex dv4 Tables", x = weightsat16bsex1tabledf, startCol = 1, startRow = 24)
writeData(wb, sheet = "Sex dv4 Tables", x = weightsat16bsex2tabledf, startCol = 6, startRow = 24)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 3
weightsat16csex1table1 <- wtd.table(complete$c_sat16_dv4[complete$dem3_dv == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem3_dv == 1])
weightsat16csex1sum <- sum(weightsat16csex1table1$sum.of.weights)
weightsat16csex1table1$percentage <- (weightsat16csex1table1$sum.of.weights / weightsat16csex1sum) * 100
weightsat16csex1table1df <- as.data.frame(weightsat16csex1table1)
sat16csex1table <- table(complete$c_sat16_dv4[complete$dem3_dv == 1])
sat16csex1tabledf <- as.data.frame(sat16csex1table)
sat16csex1sum <- sum(sat16csex1tabledf$Freq)
sat16csex1tabledf$Percentage <- (sat16csex1tabledf$Freq / sat16csex1sum) * 100
weightsat16csex1tabledf <- data.frame(Party = weightsat16csex1table1$x, Frequency = sat16csex1tabledf$Freq, Percentage_Weighted = weightsat16csex1table1$percentage, Percentage_Unweighted = sat16csex1tabledf$Percentage)

weightsat16csex2table1 <- wtd.table(complete$c_sat16_dv4[complete$dem3_dv == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem3_dv == 2])
weightsat16csex2sum <- sum(weightsat16csex2table1$sum.of.weights)
weightsat16csex2table1$percentage <- (weightsat16csex2table1$sum.of.weights / weightsat16csex2sum) * 100
weightsat16csex2table1df <- as.data.frame(weightsat16csex2table1)
sat16csex2table <- table(complete$c_sat16_dv4[complete$dem3_dv == 2])
sat16csex2tabledf <- as.data.frame(sat16csex2table)
sat16csex2sum <- sum(sat16csex2tabledf$Freq)
sat16csex2tabledf$Percentage <- (sat16csex2tabledf$Freq / sat16csex2sum) * 100
weightsat16csex2tabledf <- data.frame(Frequency = sat16csex2tabledf$Freq, Percentage_Weighted = weightsat16csex2table1$percentage, Percentage_Unweighted = sat16csex2tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Sex dv4 Tables", x = weightsat16csex1tabledf, startCol = 1, startRow = 34)
writeData(wb, sheet = "Sex dv4 Tables", x = weightsat16csex2tabledf, startCol = 6, startRow = 34)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 4
weightsat16dsex1table1 <- wtd.table(complete$d_sat16_dv4[complete$dem3_dv == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem3_dv == 1])
weightsat16dsex1sum <- sum(weightsat16dsex1table1$sum.of.weights)
weightsat16dsex1table1$percentage <- (weightsat16dsex1table1$sum.of.weights / weightsat16dsex1sum) * 100
weightsat16dsex1table1df <- as.data.frame(weightsat16dsex1table1)
sat16dsex1table <- table(complete$d_sat16_dv4[complete$dem3_dv == 1])
sat16dsex1tabledf <- as.data.frame(sat16dsex1table)
sat16dsex1sum <- sum(sat16dsex1tabledf$Freq)
sat16dsex1tabledf$Percentage <- (sat16dsex1tabledf$Freq / sat16dsex1sum) * 100
weightsat16dsex1tabledf <- data.frame(Party = weightsat16dsex1table1$x, Frequency = sat16dsex1tabledf$Freq, Percentage_Weighted = weightsat16dsex1table1$percentage, Percentage_Unweighted = sat16dsex1tabledf$Percentage)

weightsat16dsex2table1 <- wtd.table(complete$d_sat16_dv4[complete$dem3_dv == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$dem3_dv == 2])
weightsat16dsex2sum <- sum(weightsat16dsex2table1$sum.of.weights)
weightsat16dsex2table1$percentage <- (weightsat16dsex2table1$sum.of.weights / weightsat16dsex2sum) * 100
weightsat16dsex2table1df <- as.data.frame(weightsat16dsex2table1)
sat16dsex2table <- table(complete$d_sat16_dv4[complete$dem3_dv == 2])
sat16dsex2tabledf <- as.data.frame(sat16dsex2table)
sat16dsex2sum <- sum(sat16dsex2tabledf$Freq)
sat16dsex2tabledf$Percentage <- (sat16dsex2tabledf$Freq / sat16dsex2sum) * 100
weightsat16dsex2tabledf <- data.frame(Frequency = sat16dsex2tabledf$Freq, Percentage_Weighted = weightsat16dsex2table1$percentage, Percentage_Unweighted = sat16dsex2tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Sex dv4 Tables", x = weightsat16dsex1tabledf, startCol = 1, startRow = 44)
writeData(wb, sheet = "Sex dv4 Tables", x = weightsat16dsex2tabledf, startCol = 6, startRow = 44)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#region dv4 tables
#2019 election
weightsat15region1table1 <- wtd.table(complete$sat15_dv4[complete$region_dv2 == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$region_dv2 == 1])
weightsat15region1sum <- sum(weightsat15region1table1$sum.of.weights)
weightsat15region1table1$percentage <- (weightsat15region1table1$sum.of.weights / weightsat15region1sum) * 100
weightsat15region1table1df <- as.data.frame(weightsat15region1table1)
sat15region1table <- table(complete$sat15_dv4[complete$region_dv2 == 1])
sat15region1tabledf <- as.data.frame(sat15region1table)
sat15region1sum <- sum(sat15region1tabledf$Freq)
sat15region1tabledf$Percentage <- (sat15region1tabledf$Freq / sat15region1sum) * 100
weightsat15region1tabledf <- data.frame(Party = weightsat15region1table1$x, Frequency = sat15region1tabledf$Freq, Percentage_Weighted = weightsat15region1table1$percentage, Percentage_Unweighted = sat15region1tabledf$Percentage)

weightsat15region2table1 <- wtd.table(complete$sat15_dv4[complete$region_dv2 == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$region_dv2 == 2])
weightsat15region2sum <- sum(weightsat15region2table1$sum.of.weights)
weightsat15region2table1$percentage <- (weightsat15region2table1$sum.of.weights / weightsat15region2sum) * 100
weightsat15region2table1df <- as.data.frame(weightsat15region2table1)
sat15region2table <- table(complete$sat15_dv4[complete$region_dv2 == 2])
sat15region2tabledf <- as.data.frame(sat15region2table)
sat15region2sum <- sum(sat15region2tabledf$Freq)
sat15region2tabledf$Percentage <- (sat15region2tabledf$Freq / sat15region2sum) * 100
weightsat15region2tabledf <- data.frame(Frequency = sat15region2tabledf$Freq, Percentage_Weighted = weightsat15region2table1$percentage, Percentage_Unweighted = sat15region2tabledf$Percentage)

weightsat15region999table1 <- wtd.table(complete$sat15_dv4[complete$region_dv2 == 999], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$region_dv2 == 999])
weightsat15region999sum <- sum(weightsat15region999table1$sum.of.weights)
weightsat15region999table1$percentage <- (weightsat15region999table1$sum.of.weights / weightsat15region999sum) * 100
weightsat15region999table1df <- as.data.frame(weightsat15region999table1)
sat15region999table <- table(complete$sat15_dv4[complete$region_dv2 == 999])
sat15region999tabledf <- as.data.frame(sat15region999table)
sat15region999sum <- sum(sat15region999tabledf$Freq)
sat15region999tabledf$Percentage <- (sat15region999tabledf$Freq / sat15region999sum) * 100
weightsat15region999tabledf <- data.frame(Frequency = sat15region999tabledf$Freq, Percentage_Weighted = weightsat15region999table1$percentage, Percentage_Unweighted = sat15region999tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Region dv4 Tables", x = weightsat15region1tabledf, startCol = 1, startRow = 4)
writeData(wb, sheet = "Region dv4 Tables", x = weightsat15region2tabledf, startCol = 6, startRow = 4)
writeData(wb, sheet = "Region dv4 Tables", x = weightsat15region999tabledf, startCol = 10, startRow = 4)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 1
weightsat16aregion1table1 <- wtd.table(complete$a_sat16_dv4[complete$region_dv2 == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$region_dv2 == 1])
weightsat16aregion1sum <- sum(weightsat16aregion1table1$sum.of.weights)
weightsat16aregion1table1$percentage <- (weightsat16aregion1table1$sum.of.weights / weightsat16aregion1sum) * 100
weightsat16aregion1table1df <- as.data.frame(weightsat16aregion1table1)
sat16aregion1table <- table(complete$a_sat16_dv4[complete$region_dv2 == 1])
sat16aregion1tabledf <- as.data.frame(sat16aregion1table)
sat16aregion1sum <- sum(sat16aregion1tabledf$Freq)
sat16aregion1tabledf$Percentage <- (sat16aregion1tabledf$Freq / sat16aregion1sum) * 100
weightsat16aregion1tabledf <- data.frame(Party = weightsat16aregion1table1$x, Frequency = sat16aregion1tabledf$Freq, Percentage_Weighted = weightsat16aregion1table1$percentage, Percentage_Unweighted = sat16aregion1tabledf$Percentage)

weightsat16aregion2table1 <- wtd.table(complete$a_sat16_dv4[complete$region_dv2 == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$region_dv2 == 2])
weightsat16aregion2sum <- sum(weightsat16aregion2table1$sum.of.weights)
weightsat16aregion2table1$percentage <- (weightsat16aregion2table1$sum.of.weights / weightsat16aregion2sum) * 100
weightsat16aregion2table1df <- as.data.frame(weightsat16aregion2table1)
sat16aregion2table <- table(complete$a_sat16_dv4[complete$region_dv2 == 2])
sat16aregion2tabledf <- as.data.frame(sat16aregion2table)
sat16aregion2sum <- sum(sat16aregion2tabledf$Freq)
sat16aregion2tabledf$Percentage <- (sat16aregion2tabledf$Freq / sat16aregion2sum) * 100
weightsat16aregion2tabledf <- data.frame(Frequency = sat16aregion2tabledf$Freq, Percentage_Weighted = weightsat16aregion2table1$percentage, Percentage_Unweighted = sat16aregion2tabledf$Percentage)

weightsat16aregion999table1 <- wtd.table(complete$a_sat16_dv4[complete$region_dv2 == 999], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$region_dv2 == 999])
weightsat16aregion999sum <- sum(weightsat16aregion999table1$sum.of.weights)
weightsat16aregion999table1$percentage <- (weightsat16aregion999table1$sum.of.weights / weightsat16aregion999sum) * 100
weightsat16aregion999table1df <- as.data.frame(weightsat16aregion999table1)
sat16aregion999table <- table(complete$a_sat16_dv4[complete$region_dv2 == 999])
sat16aregion999tabledf <- as.data.frame(sat16aregion999table)
sat16aregion999sum <- sum(sat16aregion999tabledf$Freq)
sat16aregion999tabledf$Percentage <- (sat16aregion999tabledf$Freq / sat16aregion999sum) * 100
weightsat16aregion999tabledf <- data.frame(Frequency = sat16aregion999tabledf$Freq, Percentage_Weighted = weightsat16aregion999table1$percentage, Percentage_Unweighted = sat16aregion999tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Region dv4 Tables", x = weightsat16aregion1tabledf, startCol = 1, startRow = 14)
writeData(wb, sheet = "Region dv4 Tables", x = weightsat16aregion2tabledf, startCol = 6, startRow = 14)
writeData(wb, sheet = "Region dv4 Tables", x = weightsat16aregion999tabledf, startCol = 10, startRow = 14)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 2
weightsat16bregion1table1 <- wtd.table(complete$b_sat16_dv4[complete$region_dv2 == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$region_dv2 == 1])
weightsat16bregion1sum <- sum(weightsat16bregion1table1$sum.of.weights)
weightsat16bregion1table1$percentage <- (weightsat16bregion1table1$sum.of.weights / weightsat16bregion1sum) * 100
weightsat16bregion1table1df <- as.data.frame(weightsat16bregion1table1)
sat16bregion1table <- table(complete$b_sat16_dv4[complete$region_dv2 == 1])
sat16bregion1tabledf <- as.data.frame(sat16bregion1table)
sat16bregion1sum <- sum(sat16bregion1tabledf$Freq)
sat16bregion1tabledf$Percentage <- (sat16bregion1tabledf$Freq / sat16bregion1sum) * 100
weightsat16bregion1tabledf <- data.frame(Party = weightsat16bregion1table1$x, Frequency = sat16bregion1tabledf$Freq, Percentage_Weighted = weightsat16bregion1table1$percentage, Percentage_Unweighted = sat16bregion1tabledf$Percentage)

weightsat16bregion2table1 <- wtd.table(complete$b_sat16_dv4[complete$region_dv2 == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$region_dv2 == 2])
weightsat16bregion2sum <- sum(weightsat16bregion2table1$sum.of.weights)
weightsat16bregion2table1$percentage <- (weightsat16bregion2table1$sum.of.weights / weightsat16bregion2sum) * 100
weightsat16bregion2table1df <- as.data.frame(weightsat16bregion2table1)
sat16bregion2table <- table(complete$b_sat16_dv4[complete$region_dv2 == 2])
sat16bregion2tabledf <- as.data.frame(sat16bregion2table)
sat16bregion2sum <- sum(sat16bregion2tabledf$Freq)
sat16bregion2tabledf$Percentage <- (sat16bregion2tabledf$Freq / sat16bregion2sum) * 100
weightsat16bregion2tabledf <- data.frame(Frequency = sat16bregion2tabledf$Freq, Percentage_Weighted = weightsat16bregion2table1$percentage, Percentage_Unweighted = sat16bregion2tabledf$Percentage)

weightsat16bregion999table1 <- wtd.table(complete$b_sat16_dv4[complete$region_dv2 == 999], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$region_dv2 == 999])
weightsat16bregion999sum <- sum(weightsat16bregion999table1$sum.of.weights)
weightsat16bregion999table1$percentage <- (weightsat16bregion999table1$sum.of.weights / weightsat16bregion999sum) * 100
weightsat16bregion999table1df <- as.data.frame(weightsat16bregion999table1)
sat16bregion999table <- table(complete$b_sat16_dv4[complete$region_dv2 == 999])
sat16bregion999tabledf <- as.data.frame(sat16bregion999table)
sat16bregion999sum <- sum(sat16bregion999tabledf$Freq)
sat16bregion999tabledf$Percentage <- (sat16bregion999tabledf$Freq / sat16bregion999sum) * 100
weightsat16bregion999tabledf <- data.frame(Frequency = sat16bregion999tabledf$Freq, Percentage_Weighted = weightsat16bregion999table1$percentage, Percentage_Unweighted = sat16bregion999tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Region dv4 Tables", x = weightsat16bregion1tabledf, startCol = 1, startRow = 24)
writeData(wb, sheet = "Region dv4 Tables", x = weightsat16bregion2tabledf, startCol = 6, startRow = 24)
writeData(wb, sheet = "Region dv4 Tables", x = weightsat16bregion999tabledf, startCol = 10, startRow = 24)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 3
weightsat16cregion1table1 <- wtd.table(complete$c_sat16_dv4[complete$region_dv2 == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$region_dv2 == 1])
weightsat16cregion1sum <- sum(weightsat16cregion1table1$sum.of.weights)
weightsat16cregion1table1$percentage <- (weightsat16cregion1table1$sum.of.weights / weightsat16cregion1sum) * 100
weightsat16cregion1table1df <- as.data.frame(weightsat16cregion1table1)
sat16cregion1table <- table(complete$c_sat16_dv4[complete$region_dv2 == 1])
sat16cregion1tabledf <- as.data.frame(sat16cregion1table)
sat16cregion1sum <- sum(sat16cregion1tabledf$Freq)
sat16cregion1tabledf$Percentage <- (sat16cregion1tabledf$Freq / sat16cregion1sum) * 100
weightsat16cregion1tabledf <- data.frame(Party = weightsat16cregion1table1$x, Frequency = sat16cregion1tabledf$Freq, Percentage_Weighted = weightsat16cregion1table1$percentage, Percentage_Unweighted = sat16cregion1tabledf$Percentage)

weightsat16cregion2table1 <- wtd.table(complete$c_sat16_dv4[complete$region_dv2 == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$region_dv2 == 2])
weightsat16cregion2sum <- sum(weightsat16cregion2table1$sum.of.weights)
weightsat16cregion2table1$percentage <- (weightsat16cregion2table1$sum.of.weights / weightsat16cregion2sum) * 100
weightsat16cregion2table1df <- as.data.frame(weightsat16cregion2table1)
sat16cregion2table <- table(complete$c_sat16_dv4[complete$region_dv2 == 2])
sat16cregion2tabledf <- as.data.frame(sat16cregion2table)
sat16cregion2sum <- sum(sat16cregion2tabledf$Freq)
sat16cregion2tabledf$Percentage <- (sat16cregion2tabledf$Freq / sat16cregion2sum) * 100
weightsat16cregion2tabledf <- data.frame(Frequency = sat16cregion1tabledf$Freq, Percentage_Weighted = weightsat16cregion2table1$percentage, Percentage_Unweighted = sat16cregion2tabledf$Percentage)

weightsat16cregion999table1 <- wtd.table(complete$c_sat16_dv4[complete$region_dv2 == 999], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$region_dv2 == 999])
weightsat16cregion999sum <- sum(weightsat16cregion999table1$sum.of.weights)
weightsat16cregion999table1$percentage <- (weightsat16cregion999table1$sum.of.weights / weightsat16cregion999sum) * 100
weightsat16cregion999table1df <- as.data.frame(weightsat16cregion999table1)
sat16cregion999table <- table(complete$c_sat16_dv4[complete$region_dv2 == 999])
sat16cregion999tabledf <- as.data.frame(sat16cregion999table)
sat16cregion999sum <- sum(sat16cregion999tabledf$Freq)
sat16cregion999tabledf$Percentage <- (sat16cregion999tabledf$Freq / sat16cregion999sum) * 100
weightsat16cregion999tabledf <- data.frame(Frequency = sat16cregion999tabledf$Freq, Percentage_Weighted = weightsat16cregion999table1$percentage, Percentage_Unweighted = sat16cregion999tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Region dv4 Tables", x = weightsat16cregion1tabledf, startCol = 1, startRow = 34)
writeData(wb, sheet = "Region dv4 Tables", x = weightsat16cregion2tabledf, startCol = 6, startRow = 34)
writeData(wb, sheet = "Region dv4 Tables", x = weightsat16cregion999tabledf, startCol = 10, startRow = 34)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 4
weightsat16dregion1table1 <- wtd.table(complete$d_sat16_dv4[complete$region_dv2 == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$region_dv2 == 1])
weightsat16dregion1sum <- sum(weightsat16dregion1table1$sum.of.weights)
weightsat16dregion1table1$percentage <- (weightsat16dregion1table1$sum.of.weights / weightsat16dregion1sum) * 100
weightsat16dregion1table1df <- as.data.frame(weightsat16dregion1table1)
sat16dregion1table <- table(complete$d_sat16_dv4[complete$region_dv2 == 1])
sat16dregion1tabledf <- as.data.frame(sat16dregion1table)
sat16dregion1sum <- sum(sat16dregion1tabledf$Freq)
sat16dregion1tabledf$Percentage <- (sat16dregion1tabledf$Freq / sat16dregion1sum) * 100
weightsat16dregion1tabledf <- data.frame(Party = weightsat16dregion1table1$x, Frequency = sat16dregion1tabledf$Freq, Percentage_Weighted = weightsat16dregion1table1$percentage, Percentage_Unweighted = sat16dregion1tabledf$Percentage)

weightsat16dregion2table1 <- wtd.table(complete$d_sat16_dv4[complete$region_dv2 == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$region_dv2 == 2])
weightsat16dregion2sum <- sum(weightsat16dregion2table1$sum.of.weights)
weightsat16dregion2table1$percentage <- (weightsat16dregion2table1$sum.of.weights / weightsat16dregion2sum) * 100
weightsat16dregion2table1df <- as.data.frame(weightsat16dregion2table1)
sat16dregion2table <- table(complete$d_sat16_dv4[complete$region_dv2 == 2])
sat16dregion2tabledf <- as.data.frame(sat16dregion2table)
sat16dregion2sum <- sum(sat16dregion2tabledf$Freq)
sat16dregion2tabledf$Percentage <- (sat16dregion2tabledf$Freq / sat16dregion2sum) * 100
weightsat16dregion2tabledf <- data.frame(Frequency = sat16dregion2tabledf$Freq, Percentage_Weighted = weightsat16dregion2table1$percentage, Percentage_Unweighted = sat16dregion2tabledf$Percentage)

weightsat16dregion999table1 <- wtd.table(complete$d_sat16_dv4[complete$region_dv2 == 999], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$region_dv2 == 999])
weightsat16dregion999sum <- sum(weightsat16dregion999table1$sum.of.weights)
weightsat16dregion999table1$percentage <- (weightsat16dregion999table1$sum.of.weights / weightsat16dregion999sum) * 100
weightsat16dregion999table1df <- as.data.frame(weightsat16dregion999table1)
sat16dregion999table <- table(complete$d_sat16_dv4[complete$region_dv2 == 999])
sat16dregion999tabledf <- as.data.frame(sat16dregion999table)
sat16dregion999sum <- sum(sat16dregion999tabledf$Freq)
sat16dregion999tabledf$Percentage <- (sat16dregion999tabledf$Freq / sat16dregion999sum) * 100
weightsat16dregion999tabledf <- data.frame(Frequency = sat16dregion999tabledf$Freq, Percentage_Weighted = weightsat16dregion999table1$percentage, Percentage_Unweighted = sat16dregion999tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Region dv4 Tables", x = weightsat16dregion1tabledf, startCol = 1, startRow = 44)
writeData(wb, sheet = "Region dv4 Tables", x = weightsat16dregion2tabledf, startCol = 6, startRow = 44)
writeData(wb, sheet = "Region dv4 Tables", x = weightsat16dregion999tabledf, startCol = 10, startRow = 44)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#age dv4 excel tables
#2019 election
weightsat15age1table1 <- wtd.table(complete$sat15_dv4[complete$a_age_4_dv == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_age_4_dv == 1])
weightsat15age1sum <- sum(weightsat15age1table1$sum.of.weights)
weightsat15age1table1$percentage <- (weightsat15age1table1$sum.of.weights / weightsat15age1sum) * 100
weightsat15age1table1df <- as.data.frame(weightsat15age1table1)
sat15age1table <- table(complete$sat15_dv4[complete$a_age_4_dv == 1])
sat15age1tabledf <- as.data.frame(sat15age1table)
sat15age1sum <- sum(sat15age1tabledf$Freq)
sat15age1tabledf$Percentage <- (sat15age1tabledf$Freq / sat15age1sum) * 100
weightsat15age1tabledf <- data.frame(Party = weightsat15age1table1$x, Frequency = sat15age1tabledf$Freq, Percentage_Weighted = weightsat15age1table1$percentage, Percentage_Unweighted = sat15age1tabledf$Percentage)

weightsat15age2table1 <- wtd.table(complete$sat15_dv4[complete$a_age_4_dv == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_age_4_dv == 2])
weightsat15age2sum <- sum(weightsat15age2table1$sum.of.weights)
weightsat15age2table1$percentage <- (weightsat15age2table1$sum.of.weights / weightsat15age2sum) * 100
weightsat15age2table1df <- as.data.frame(weightsat15age2table1)
sat15age2table <- table(complete$sat15_dv4[complete$a_age_4_dv == 2])
sat15age2tabledf <- as.data.frame(sat15age2table)
sat15age2sum <- sum(sat15age2tabledf$Freq)
sat15age2tabledf$Percentage <- (sat15age2tabledf$Freq / sat15age2sum) * 100
weightsat15age2tabledf <- data.frame(Frequency = sat15age2tabledf$Freq, Percentage_Weighted = weightsat15age2table1$percentage, Percentage_Unweighted = sat15age2tabledf$Percentage)

weightsat15age3table1 <- wtd.table(complete$sat15_dv4[complete$a_age_4_dv == 3], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_age_4_dv == 3])
weightsat15age3sum <- sum(weightsat15age3table1$sum.of.weights)
weightsat15age3table1$percentage <- (weightsat15age3table1$sum.of.weights / weightsat15age3sum) * 100
weightsat15age3table1df <- as.data.frame(weightsat15age3table1)
sat15age3table <- table(complete$sat15_dv4[complete$a_age_4_dv == 3])
sat15age3tabledf <- as.data.frame(sat15age3table)
sat15age3sum <- sum(sat15age3tabledf$Freq)
sat15age3tabledf$Percentage <- (sat15age3tabledf$Freq / sat15age3sum) * 100
weightsat15age3tabledf <- data.frame(Frequency = sat15age3tabledf$Freq, Percentage_Weighted = weightsat15age3table1$percentage, Percentage_Unweighted = sat15age3tabledf$Percentage)

weightsat15age4table1 <- wtd.table(complete$sat15_dv4[complete$a_age_4_dv == 4], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_age_4_dv == 4])
weightsat15age4sum <- sum(weightsat15age4table1$sum.of.weights)
weightsat15age4table1$percentage <- (weightsat15age4table1$sum.of.weights / weightsat15age4sum) * 100
weightsat15age4table1df <- as.data.frame(weightsat15age4table1)
sat15age4table <- table(complete$sat15_dv4[complete$a_age_4_dv == 4])
sat15age4tabledf <- as.data.frame(sat15age4table)
sat15age4sum <- sum(sat15age4tabledf$Freq)
sat15age4tabledf$Percentage <- (sat15age4tabledf$Freq / sat15age4sum) * 100
weightsat15age4tabledf <- data.frame(Frequency = sat15age4tabledf$Freq, Percentage_Weighted = weightsat15age4table1$percentage, Percentage_Unweighted = sat15age4tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Age dv4 Tables", x = weightsat15age1tabledf, startCol = 1, startRow = 4)
writeData(wb, sheet = "Age dv4 Tables", x = weightsat15age2tabledf, startCol = 6, startRow = 4)
writeData(wb, sheet = "Age dv4 Tables", x = weightsat15age3tabledf, startCol = 10, startRow = 4)
writeData(wb, sheet = "Age dv4 Tables", x = weightsat15age4tabledf, startCol = 14, startRow = 4)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 1
weightsat16aage1table1 <- wtd.table(complete$a_sat16_dv4[complete$a_age_4_dv == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_age_4_dv == 1])
weightsat16aage1sum <- sum(weightsat16aage1table1$sum.of.weights)
weightsat16aage1table1$percentage <- (weightsat16aage1table1$sum.of.weights / weightsat16aage1sum) * 100
weightsat16aage1table1df <- as.data.frame(weightsat16aage1table1)
sat16aage1table <- table(complete$a_sat16_dv4[complete$a_age_4_dv == 1])
sat16aage1tabledf <- as.data.frame(sat16aage1table)
sat16aage1sum <- sum(sat16aage1tabledf$Freq)
sat16aage1tabledf$Percentage <- (sat16aage1tabledf$Freq / sat16aage1sum) * 100
weightsat16aage1tabledf <- data.frame(Party = weightsat16aage1table1$x, Frequency = sat16aage1tabledf$Freq, Percentage_Weighted = weightsat16aage1table1$percentage, Percentage_Unweighted = sat16aage1tabledf$Percentage)

weightsat16aage2table1 <- wtd.table(complete$a_sat16_dv4[complete$a_age_4_dv == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_age_4_dv == 2])
weightsat16aage2sum <- sum(weightsat16aage2table1$sum.of.weights)
weightsat16aage2table1$percentage <- (weightsat16aage2table1$sum.of.weights / weightsat16aage2sum) * 100
weightsat16aage2table1df <- as.data.frame(weightsat16aage2table1)
sat16aage2table <- table(complete$a_sat16_dv4[complete$a_age_4_dv == 2])
sat16aage2tabledf <- as.data.frame(sat16aage2table)
sat16aage2sum <- sum(sat16aage2tabledf$Freq)
sat16aage2tabledf$Percentage <- (sat16aage2tabledf$Freq / sat16aage2sum) * 100
weightsat16aage2tabledf <- data.frame(Frequency = sat16aage2tabledf$Freq, Percentage_Weighted = weightsat16aage2table1$percentage, Percentage_Unweighted = sat16aage2tabledf$Percentage)

weightsat16aage3table1 <- wtd.table(complete$a_sat16_dv4[complete$a_age_4_dv == 3], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_age_4_dv == 3])
weightsat16aage3sum <- sum(weightsat16aage3table1$sum.of.weights)
weightsat16aage3table1$percentage <- (weightsat16aage3table1$sum.of.weights / weightsat16aage3sum) * 100
weightsat16aage3table1df <- as.data.frame(weightsat16aage3table1)
sat16aage3table <- table(complete$a_sat16_dv4[complete$a_age_4_dv == 3])
sat16aage3tabledf <- as.data.frame(sat16aage3table)
sat16aage3sum <- sum(sat16aage3tabledf$Freq)
sat16aage3tabledf$Percentage <- (sat16aage3tabledf$Freq / sat16aage3sum) * 100
weightsat16aage3tabledf <- data.frame(Frequency = sat16aage3tabledf$Freq, Percentage_Weighted = weightsat16aage3table1$percentage, Percentage_Unweighted = sat16aage3tabledf$Percentage)

weightsat16aage4table1 <- wtd.table(complete$a_sat16_dv4[complete$a_age_4_dv == 4], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_age_4_dv == 4])
weightsat16aage4sum <- sum(weightsat16aage4table1$sum.of.weights)
weightsat16aage4table1$percentage <- (weightsat16aage4table1$sum.of.weights / weightsat16aage4sum) * 100
weightsat16aage4table1df <- as.data.frame(weightsat16aage4table1)
sat16aage4table <- table(complete$a_sat16_dv4[complete$a_age_4_dv == 4])
sat16aage4tabledf <- as.data.frame(sat16aage4table)
sat16aage4sum <- sum(sat16aage4tabledf$Freq)
sat16aage4tabledf$Percentage <- (sat16aage4tabledf$Freq / sat16aage4sum) * 100
weightsat16aage4tabledf <- data.frame(Frequency = sat16aage4tabledf$Freq, Percentage_Weighted = weightsat16aage4table1$percentage, Percentage_Unweighted = sat16aage4tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16aage1tabledf, startCol = 1, startRow = 14)
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16aage2tabledf, startCol = 6, startRow = 14)
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16aage3tabledf, startCol = 10, startRow = 14)
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16aage4tabledf, startCol = 14, startRow = 14)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 2
weightsat16bage1table1 <- wtd.table(complete$b_sat16_dv4[complete$b_age_4_dv == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$b_age_4_dv == 1])
weightsat16bage1sum <- sum(weightsat16bage1table1$sum.of.weights)
weightsat16bage1table1$percentage <- (weightsat16bage1table1$sum.of.weights / weightsat16bage1sum) * 100
weightsat16bage1table1df <- as.data.frame(weightsat16bage1table1)
sat16bage1table <- table(complete$b_sat16_dv4[complete$b_age_4_dv == 1])
sat16bage1tabledf <- as.data.frame(sat16bage1table)
sat16bage1sum <- sum(sat16bage1tabledf$Freq)
sat16bage1tabledf$Percentage <- (sat16bage1tabledf$Freq / sat16bage1sum) * 100
weightsat16bage1tabledf <- data.frame(Party = weightsat16bage1table1$x, Frequency = sat16bage1tabledf$Freq, Percentage_Weighted = weightsat16bage1table1$percentage, Percentage_Unweighted = sat16bage1tabledf$Percentage)

weightsat16bage2table1 <- wtd.table(complete$b_sat16_dv4[complete$b_age_4_dv == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$b_age_4_dv == 2])
weightsat16bage2sum <- sum(weightsat16bage2table1$sum.of.weights)
weightsat16bage2table1$percentage <- (weightsat16bage2table1$sum.of.weights / weightsat16bage2sum) * 100
weightsat16bage2table1df <- as.data.frame(weightsat16bage2table1)
sat16bage2table <- table(complete$b_sat16_dv4[complete$b_age_4_dv == 2])
sat16bage2tabledf <- as.data.frame(sat16bage2table)
sat16bage2sum <- sum(sat16bage2tabledf$Freq)
sat16bage2tabledf$Percentage <- (sat16bage2tabledf$Freq / sat16bage2sum) * 100
weightsat16bage2tabledf <- data.frame(Frequency = sat16bage2tabledf$Freq, Percentage_Weighted = weightsat16bage2table1$percentage, Percentage_Unweighted = sat16bage2tabledf$Percentage)

weightsat16bage3table1 <- wtd.table(complete$b_sat16_dv4[complete$b_age_4_dv == 3], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$b_age_4_dv == 3])
weightsat16bage3sum <- sum(weightsat16bage3table1$sum.of.weights)
weightsat16bage3table1$percentage <- (weightsat16bage3table1$sum.of.weights / weightsat16bage3sum) * 100
weightsat16bage3table1df <- as.data.frame(weightsat16bage3table1)
sat16bage3table <- table(complete$b_sat16_dv4[complete$b_age_4_dv == 3])
sat16bage3tabledf <- as.data.frame(sat16bage3table)
sat16bage3sum <- sum(sat16bage3tabledf$Freq)
sat16bage3tabledf$Percentage <- (sat16bage3tabledf$Freq / sat16bage3sum) * 100
weightsat16bage3tabledf <- data.frame(Frequency = sat16bage3tabledf$Freq, Percentage_Weighted = weightsat16bage3table1$percentage, Percentage_Unweighted = sat16bage3tabledf$Percentage)

weightsat16bage4table1 <- wtd.table(complete$b_sat16_dv4[complete$b_age_4_dv == 4], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$b_age_4_dv == 4])
weightsat16bage4sum <- sum(weightsat16bage4table1$sum.of.weights)
weightsat16bage4table1$percentage <- (weightsat16bage4table1$sum.of.weights / weightsat16bage4sum) * 100
weightsat16bage4table1df <- as.data.frame(weightsat16bage4table1)
sat16bage4table <- table(complete$b_sat16_dv4[complete$b_age_4_dv == 4])
sat16bage4tabledf <- as.data.frame(sat16bage4table)
sat16bage4sum <- sum(sat16bage4tabledf$Freq)
sat16bage4tabledf$Percentage <- (sat16bage4tabledf$Freq / sat16bage4sum) * 100
weightsat16bage4tabledf <- data.frame(Frequency = sat16bage4tabledf$Freq, Percentage_Weighted = weightsat16bage4table1$percentage, Percentage_Unweighted = sat16bage4tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16bage1tabledf, startCol = 1, startRow = 24)
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16bage2tabledf, startCol = 6, startRow = 24)
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16bage3tabledf, startCol = 10, startRow = 24)
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16bage4tabledf, startCol = 14, startRow = 24)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 3
weightsat16cage1table1 <- wtd.table(complete$c_sat16_dv4[complete$c_age_4_dv == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$c_age_4_dv == 1])
weightsat16cage1sum <- sum(weightsat16cage1table1$sum.of.weights)
weightsat16cage1table1$percentage <- (weightsat16cage1table1$sum.of.weights / weightsat16cage1sum) * 100
weightsat16cage1table1df <- as.data.frame(weightsat16cage1table1)
sat16cage1table <- table(complete$c_sat16_dv4[complete$c_age_4_dv == 1])
sat16cage1tabledf <- as.data.frame(sat16cage1table)
sat16cage1sum <- sum(sat16cage1tabledf$Freq)
sat16cage1tabledf$Percentage <- (sat16cage1tabledf$Freq / sat16cage1sum) * 100
weightsat16cage1tabledf <- data.frame(Party = weightsat16cage1table1$x, Frequency = sat16cage1tabledf$Freq, Percentage_Weighted = weightsat16cage1table1$percentage, Percentage_Unweighted = sat16cage1tabledf$Percentage)

weightsat16cage2table1 <- wtd.table(complete$c_sat16_dv4[complete$c_age_4_dv == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$c_age_4_dv == 2])
weightsat16cage2sum <- sum(weightsat16cage2table1$sum.of.weights)
weightsat16cage2table1$percentage <- (weightsat16cage2table1$sum.of.weights / weightsat16cage2sum) * 100
weightsat16cage2table1df <- as.data.frame(weightsat16cage2table1)
sat16cage2table <- table(complete$c_sat16_dv4[complete$c_age_4_dv == 2])
sat16cage2tabledf <- as.data.frame(sat16cage2table)
sat16cage2sum <- sum(sat16cage2tabledf$Freq)
sat16cage2tabledf$Percentage <- (sat16cage2tabledf$Freq / sat16cage2sum) * 100
weightsat16cage2tabledf <- data.frame(Frequency = sat16cage2tabledf$Freq, Percentage_Weighted = weightsat16cage2table1$percentage, Percentage_Unweighted = sat16cage2tabledf$Percentage)

weightsat16cage3table1 <- wtd.table(complete$c_sat16_dv4[complete$c_age_4_dv == 3], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$c_age_4_dv == 3])
weightsat16cage3sum <- sum(weightsat16cage3table1$sum.of.weights)
weightsat16cage3table1$percentage <- (weightsat16cage3table1$sum.of.weights / weightsat16cage3sum) * 100
weightsat16cage3table1df <- as.data.frame(weightsat16cage3table1)
sat16cage3table <- table(complete$c_sat16_dv4[complete$c_age_4_dv == 3])
sat16cage3tabledf <- as.data.frame(sat16cage3table)
sat16cage3sum <- sum(sat16cage3tabledf$Freq)
sat16cage3tabledf$Percentage <- (sat16cage3tabledf$Freq / sat16cage3sum) * 100
weightsat16cage3tabledf <- data.frame(Frequency = sat16cage3tabledf$Freq, Percentage_Weighted = weightsat16cage3table1$percentage, Percentage_Unweighted = sat16cage3tabledf$Percentage)

weightsat16cage4table1 <- wtd.table(complete$c_sat16_dv4[complete$c_age_4_dv == 4], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$c_age_4_dv == 4])
weightsat16cage4sum <- sum(weightsat16cage4table1$sum.of.weights)
weightsat16cage4table1$percentage <- (weightsat16cage4table1$sum.of.weights / weightsat16cage4sum) * 100
weightsat16cage4table1df <- as.data.frame(weightsat16cage4table1)
sat16cage4table <- table(complete$c_sat16_dv4[complete$c_age_4_dv == 4])
sat16cage4tabledf <- as.data.frame(sat16cage4table)
sat16cage4sum <- sum(sat16cage4tabledf$Freq)
sat16cage4tabledf$Percentage <- (sat16cage4tabledf$Freq / sat16cage4sum) * 100
weightsat16cage4tabledf <- data.frame(Frequency = sat16cage4tabledf$Freq, Percentage_Weighted = weightsat16cage4table1$percentage, Percentage_Unweighted = sat16cage4tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16cage1tabledf, startCol = 1, startRow = 34)
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16cage2tabledf, startCol = 6, startRow = 34)
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16cage3tabledf, startCol = 10, startRow = 34)
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16cage4tabledf, startCol = 14, startRow = 34)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 4
weightsat16dage1table1 <- wtd.table(complete$d_sat16_dv4[complete$d_age_4_dv == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$d_age_4_dv == 1])
weightsat16dage1sum <- sum(weightsat16dage1table1$sum.of.weights)
weightsat16dage1table1$percentage <- (weightsat16dage1table1$sum.of.weights / weightsat16dage1sum) * 100
weightsat16dage1table1df <- as.data.frame(weightsat16dage1table1)
sat16dage1table <- table(complete$d_sat16_dv4[complete$d_age_4_dv == 1])
sat16dage1tabledf <- as.data.frame(sat16dage1table)
sat16dage1sum <- sum(sat16dage1tabledf$Freq)
sat16dage1tabledf$Percentage <- (sat16dage1tabledf$Freq / sat16dage1sum) * 100
weightsat16dage1tabledf <- data.frame(Party = weightsat16dage1table1$x, Frequency = sat16dage1tabledf$Freq, Percentage_Weighted = weightsat16dage1table1$percentage, Percentage_Unweighted = sat16dage1tabledf$Percentage)

weightsat16dage2table1 <- wtd.table(complete$d_sat16_dv4[complete$d_age_4_dv == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$d_age_4_dv == 2])
weightsat16dage2sum <- sum(weightsat16dage2table1$sum.of.weights)
weightsat16dage2table1$percentage <- (weightsat16dage2table1$sum.of.weights / weightsat16dage2sum) * 100
weightsat16dage2table1df <- as.data.frame(weightsat16dage2table1)
sat16dage2table <- table(complete$d_sat16_dv4[complete$d_age_4_dv == 2])
sat16dage2tabledf <- as.data.frame(sat16dage2table)
sat16dage2sum <- sum(sat16dage2tabledf$Freq)
sat16dage2tabledf$Percentage <- (sat16dage2tabledf$Freq / sat16dage2sum) * 100
weightsat16dage2tabledf <- data.frame(Frequency = sat16dage2tabledf$Freq, Percentage_Weighted = weightsat16dage2table1$percentage, Percentage_Unweighted = sat16dage2tabledf$Percentage)

weightsat16dage3table1 <- wtd.table(complete$d_sat16_dv4[complete$d_age_4_dv == 3], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$d_age_4_dv == 3])
weightsat16dage3sum <- sum(weightsat16dage3table1$sum.of.weights)
weightsat16dage3table1$percentage <- (weightsat16dage3table1$sum.of.weights / weightsat16dage3sum) * 100
weightsat16dage3table1df <- as.data.frame(weightsat16dage3table1)
sat16dage3table <- table(complete$d_sat16_dv4[complete$d_age_4_dv == 3])
sat16dage3tabledf <- as.data.frame(sat16dage3table)
sat16dage3sum <- sum(sat16dage3tabledf$Freq)
sat16dage3tabledf$Percentage <- (sat16dage3tabledf$Freq / sat16dage3sum) * 100
weightsat16dage3tabledf <- data.frame(Frequency = sat16dage3tabledf$Freq, Percentage_Weighted = weightsat16dage3table1$percentage, Percentage_Unweighted = sat16dage3tabledf$Percentage)

weightsat16dage4table1 <- wtd.table(complete$d_sat16_dv4[complete$d_age_4_dv == 4], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$d_age_4_dv == 4])
weightsat16dage4sum <- sum(weightsat16dage4table1$sum.of.weights)
weightsat16dage4table1$percentage <- (weightsat16dage4table1$sum.of.weights / weightsat16dage4sum) * 100
weightsat16dage4table1df <- as.data.frame(weightsat16dage4table1)
sat16dage4table <- table(complete$d_sat16_dv4[complete$d_age_4_dv == 4])
sat16dage4tabledf <- as.data.frame(sat16dage4table)
sat16dage4sum <- sum(sat16dage4tabledf$Freq)
sat16dage4tabledf$Percentage <- (sat16dage4tabledf$Freq / sat16dage4sum) * 100
weightsat16dage4tabledf <- data.frame(Frequency = sat16dage4tabledf$Freq, Percentage_Weighted = weightsat16dage4table1$percentage, Percentage_Unweighted = sat16dage4tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16dage1tabledf, startCol = 1, startRow = 44)
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16dage2tabledf, startCol = 6, startRow = 44)
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16dage3tabledf, startCol = 10, startRow = 44)
writeData(wb, sheet = "Age dv4 Tables", x = weightsat16dage4tabledf, startCol = 14, startRow = 44)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#denomination dv4 excel tables
#2019 election
weightsat15denom0table1 <- wtd.table(complete$sat15_dv4[complete$a_denomination_dv3 == 0], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_denomination_dv3 == 0])
weightsat15denom0sum <- sum(weightsat15denom0table1$sum.of.weights)
weightsat15denom0table1$percentage <- (weightsat15denom0table1$sum.of.weights / weightsat15denom0sum) * 100
weightsat15denom0table1df <- as.data.frame(weightsat15denom0table1)
sat15denom0table <- table(complete$sat15_dv4[complete$a_denomination_dv3 == 0])
sat15denom0tabledf <- as.data.frame(sat15denom0table)
sat15denom0sum <- sum(sat15denom0tabledf$Freq)
sat15denom0tabledf$Percentage <- (sat15denom0tabledf$Freq / sat15denom0sum) * 100
weightsat15denom0tabledf <- data.frame(Party = weightsat15denom0table1$x, Frequency = sat15denom0tabledf$Freq, Percentage_Weighted = weightsat15denom0table1$percentage, Percentage_Unweighted = sat15denom0tabledf$Percentage)

weightsat15denom1table1 <- wtd.table(complete$sat15_dv4[complete$a_denomination_dv3 == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_denomination_dv3 == 1])
weightsat15denom1sum <- sum(weightsat15denom1table1$sum.of.weights)
weightsat15denom1table1$percentage <- (weightsat15denom1table1$sum.of.weights / weightsat15denom1sum) * 100
weightsat15denom1table1df <- as.data.frame(weightsat15denom1table1)
sat15denom1table <- table(complete$sat15_dv4[complete$a_denomination_dv3 == 1])
sat15denom1tabledf <- as.data.frame(sat15denom1table)
sat15denom1sum <- sum(sat15denom1tabledf$Freq)
sat15denom1tabledf$Percentage <- (sat15denom1tabledf$Freq / sat15denom1sum) * 100
weightsat15denom1tabledf <- data.frame(Frequency = sat15denom1tabledf$Freq, Percentage_Weighted = weightsat15denom1table1$percentage, Percentage_Unweighted = sat15denom1tabledf$Percentage)

weightsat15denom2table1 <- wtd.table(complete$sat15_dv4[complete$a_denomination_dv3 == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_denomination_dv3 == 2])
weightsat15denom2sum <- sum(weightsat15denom2table1$sum.of.weights)
weightsat15denom2table1$percentage <- (weightsat15denom2table1$sum.of.weights / weightsat15denom2sum) * 100
weightsat15denom2table1df <- as.data.frame(weightsat15denom2table1)
sat15denom2table <- table(complete$sat15_dv4[complete$a_denomination_dv3 == 2])
sat15denom2tabledf <- as.data.frame(sat15denom2table)
sat15denom2sum <- sum(sat15denom2tabledf$Freq)
sat15denom2tabledf$Percentage <- (sat15denom2tabledf$Freq / sat15denom2sum) * 100
weightsat15denom2tabledf <- data.frame(Frequency = sat15denom2tabledf$Freq, Percentage_Weighted = weightsat15denom2table1$percentage, Percentage_Unweighted = sat15denom2tabledf$Percentage)

weightsat15denom3table1 <- wtd.table(complete$sat15_dv4[complete$a_denomination_dv3 == 3], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_denomination_dv3 == 3])
weightsat15denom3sum <- sum(weightsat15denom3table1$sum.of.weights)
weightsat15denom3table1$percentage <- (weightsat15denom3table1$sum.of.weights / weightsat15denom3sum) * 100
weightsat15denom3table1df <- as.data.frame(weightsat15denom3table1)
sat15denom3table <- table(complete$sat15_dv4[complete$a_denomination_dv3 == 3])
sat15denom3tabledf <- as.data.frame(sat15denom3table)
sat15denom3sum <- sum(sat15denom3tabledf$Freq)
sat15denom3tabledf$Percentage <- (sat15denom3tabledf$Freq / sat15denom3sum) * 100
weightsat15denom3tabledf <- data.frame(Frequency = sat15denom3tabledf$Freq, Percentage_Weighted = weightsat15denom3table1$percentage, Percentage_Unweighted = sat15denom3tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat15denom0tabledf, startCol = 1, startRow = 4)
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat15denom1tabledf, startCol = 6, startRow = 4)
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat15denom2tabledf, startCol = 10, startRow = 4)
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat15denom3tabledf, startCol = 14, startRow = 4)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 1
weightsat16adenom0table1 <- wtd.table(complete$a_sat16_dv4[complete$a_denomination_dv3 == 0], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_denomination_dv3 == 0])
weightsat16adenom0sum <- sum(weightsat16adenom0table1$sum.of.weights)
weightsat16adenom0table1$percentage <- (weightsat16adenom0table1$sum.of.weights / weightsat16adenom0sum) * 100
weightsat16adenom0table1df <- as.data.frame(weightsat16adenom0table1)
sat16adenom0table <- table(complete$a_sat16_dv4[complete$a_denomination_dv3 == 0])
sat16adenom0tabledf <- as.data.frame(sat16adenom0table)
sat16adenom0sum <- sum(sat16adenom0tabledf$Freq)
sat16adenom0tabledf$Percentage <- (sat16adenom0tabledf$Freq / sat16adenom0sum) * 100
weightsat16adenom0tabledf <- data.frame(Party = weightsat16adenom0table1$x, Frequency = sat16adenom0tabledf$Freq, Percentage_Weighted = weightsat16adenom0table1$percentage, Percentage_Unweighted = sat16adenom0tabledf$Percentage)

weightsat16adenom1table1 <- wtd.table(complete$a_sat16_dv4[complete$a_denomination_dv3 == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_denomination_dv3 == 1])
weightsat16adenom1sum <- sum(weightsat16adenom1table1$sum.of.weights)
weightsat16adenom1table1$percentage <- (weightsat16adenom1table1$sum.of.weights / weightsat16adenom1sum) * 100
weightsat16adenom1table1df <- as.data.frame(weightsat16adenom1table1)
sat16adenom1table <- table(complete$a_sat16_dv4[complete$a_denomination_dv3 == 1])
sat16adenom1tabledf <- as.data.frame(sat16adenom1table)
sat16adenom1sum <- sum(sat16adenom1tabledf$Freq)
sat16adenom1tabledf$Percentage <- (sat16adenom1tabledf$Freq / sat16adenom1sum) * 100
weightsat16adenom1tabledf <- data.frame(Frequency = sat16adenom1tabledf$Freq, Percentage_Weighted = weightsat16adenom1table1$percentage, Percentage_Unweighted = sat16adenom1tabledf$Percentage)

weightsat16adenom2table1 <- wtd.table(complete$a_sat16_dv4[complete$a_denomination_dv3 == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_denomination_dv3 == 2])
weightsat16adenom2sum <- sum(weightsat16adenom2table1$sum.of.weights)
weightsat16adenom2table1$percentage <- (weightsat16adenom2table1$sum.of.weights / weightsat16adenom2sum) * 100
weightsat16adenom2table1df <- as.data.frame(weightsat16adenom2table1)
sat16adenom2table <- table(complete$a_sat16_dv4[complete$a_denomination_dv3 == 2])
sat16adenom2tabledf <- as.data.frame(sat16adenom2table)
sat16adenom2sum <- sum(sat16adenom2tabledf$Freq)
sat16adenom2tabledf$Percentage <- (sat16adenom2tabledf$Freq / sat16adenom2sum) * 100
weightsat16adenom2tabledf <- data.frame(Frequency = sat16adenom2tabledf$Freq, Percentage_Weighted = weightsat16adenom2table1$percentage, Percentage_Unweighted = sat16adenom2tabledf$Percentage)

weightsat16adenom3table1 <- wtd.table(complete$a_sat16_dv4[complete$a_denomination_dv3 == 3], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$a_denomination_dv3 == 3])
weightsat16adenom3sum <- sum(weightsat16adenom3table1$sum.of.weights)
weightsat16adenom3table1$percentage <- (weightsat16adenom3table1$sum.of.weights / weightsat16adenom3sum) * 100
weightsat16adenom3table1df <- as.data.frame(weightsat16adenom3table1)
sat16adenom3table <- table(complete$a_sat16_dv4[complete$a_denomination_dv3 == 3])
sat16adenom3tabledf <- as.data.frame(sat16adenom3table)
sat16adenom3sum <- sum(sat16adenom3tabledf$Freq)
sat16adenom3tabledf$Percentage <- (sat16adenom3tabledf$Freq / sat16adenom3sum) * 100
weightsat16adenom3tabledf <- data.frame(Frequency = sat16adenom3tabledf$Freq, Percentage_Weighted = weightsat16adenom3table1$percentage, Percentage_Unweighted = sat16adenom3tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16adenom0tabledf, startCol = 1, startRow = 14)
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16adenom1tabledf, startCol = 6, startRow = 14)
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16adenom2tabledf, startCol = 10, startRow = 14)
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16adenom3tabledf, startCol = 14, startRow = 14)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 2
weightsat16bdenom0table1 <- wtd.table(complete$b_sat16_dv4[complete$b_denomination_dv3 == 0], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$b_denomination_dv3 == 0])
weightsat16bdenom0sum <- sum(weightsat16bdenom0table1$sum.of.weights)
weightsat16bdenom0table1$percentage <- (weightsat16bdenom0table1$sum.of.weights / weightsat16bdenom0sum) * 100
weightsat16bdenom0table1df <- as.data.frame(weightsat16bdenom0table1)
sat16bdenom0table <- table(complete$b_sat16_dv4[complete$b_denomination_dv3 == 0])
sat16bdenom0tabledf <- as.data.frame(sat16bdenom0table)
sat16bdenom0sum <- sum(sat16bdenom0tabledf$Freq)
sat16bdenom0tabledf$Percentage <- (sat16bdenom0tabledf$Freq / sat16bdenom0sum) * 100
weightsat16bdenom0tabledf <- data.frame(Party = weightsat16bdenom0table1$x, Frequency = sat16bdenom0tabledf$Freq, Percentage_Weighted = weightsat16bdenom0table1$percentage, Percentage_Unweighted = sat16bdenom0tabledf$Percentage)

weightsat16bdenom1table1 <- wtd.table(complete$b_sat16_dv4[complete$b_denomination_dv3 == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$b_denomination_dv3 == 1])
weightsat16bdenom1sum <- sum(weightsat16bdenom1table1$sum.of.weights)
weightsat16bdenom1table1$percentage <- (weightsat16bdenom1table1$sum.of.weights / weightsat16bdenom1sum) * 100
weightsat16bdenom1table1df <- as.data.frame(weightsat16bdenom1table1)
sat16bdenom1table <- table(complete$b_sat16_dv4[complete$b_denomination_dv3 == 1])
sat16bdenom1tabledf <- as.data.frame(sat16bdenom1table)
sat16bdenom1sum <- sum(sat16bdenom1tabledf$Freq)
sat16bdenom1tabledf$Percentage <- (sat16bdenom1tabledf$Freq / sat16bdenom1sum) * 100
weightsat16bdenom1tabledf <- data.frame(Frequency = sat16bdenom1tabledf$Freq, Percentage_Weighted = weightsat16bdenom1table1$percentage, Percentage_Unweighted = sat16bdenom1tabledf$Percentage)

weightsat16bdenom2table1 <- wtd.table(complete$b_sat16_dv4[complete$b_denomination_dv3 == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$b_denomination_dv3 == 2])
weightsat16bdenom2sum <- sum(weightsat16bdenom2table1$sum.of.weights)
weightsat16bdenom2table1$percentage <- (weightsat16bdenom2table1$sum.of.weights / weightsat16bdenom2sum) * 100
weightsat16bdenom2table1df <- as.data.frame(weightsat16bdenom2table1)
sat16bdenom2table <- table(complete$b_sat16_dv4[complete$b_denomination_dv3 == 2])
sat16bdenom2tabledf <- as.data.frame(sat16bdenom2table)
sat16bdenom2sum <- sum(sat16bdenom2tabledf$Freq)
sat16bdenom2tabledf$Percentage <- (sat16bdenom2tabledf$Freq / sat16bdenom2sum) * 100
weightsat16bdenom2tabledf <- data.frame(Frequency = sat16bdenom2tabledf$Freq, Percentage_Weighted = weightsat16bdenom2table1$percentage, Percentage_Unweighted = sat16bdenom2tabledf$Percentage)

weightsat16bdenom3table1 <- wtd.table(complete$b_sat16_dv4[complete$b_denomination_dv3 == 3], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$b_denomination_dv3 == 3])
weightsat16bdenom3sum <- sum(weightsat16bdenom3table1$sum.of.weights)
weightsat16bdenom3table1$percentage <- (weightsat16bdenom3table1$sum.of.weights / weightsat16bdenom3sum) * 100
weightsat16bdenom3table1df <- as.data.frame(weightsat16bdenom3table1)
sat16bdenom3table <- table(complete$b_sat16_dv4[complete$b_denomination_dv3 == 3])
sat16bdenom3tabledf <- as.data.frame(sat16bdenom3table)
sat16bdenom3sum <- sum(sat16bdenom3tabledf$Freq)
sat16bdenom3tabledf$Percentage <- (sat16bdenom3tabledf$Freq / sat16bdenom3sum) * 100
weightsat16bdenom3tabledf <- data.frame(Frequency = sat16bdenom3tabledf$Freq, Percentage_Weighted = weightsat16bdenom3table1$percentage, Percentage_Unweighted = sat16bdenom3tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16bdenom0tabledf, startCol = 1, startRow = 24)
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16bdenom1tabledf, startCol = 6, startRow = 24)
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16bdenom2tabledf, startCol = 10, startRow = 24)
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16bdenom3tabledf, startCol = 14, startRow = 24)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 3
weightsat16cdenom0table1 <- wtd.table(complete$c_sat16_dv4[complete$c_denomination_dv3 == 0], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$c_denomination_dv3 == 0])
weightsat16cdenom0sum <- sum(weightsat16cdenom0table1$sum.of.weights)
weightsat16cdenom0table1$percentage <- (weightsat16cdenom0table1$sum.of.weights / weightsat16cdenom0sum) * 100
weightsat16cdenom0table1df <- as.data.frame(weightsat16cdenom0table1)
sat16cdenom0table <- table(complete$c_sat16_dv4[complete$c_denomination_dv3 == 0])
sat16cdenom0tabledf <- as.data.frame(sat16cdenom0table)
sat16cdenom0sum <- sum(sat16cdenom0tabledf$Freq)
sat16cdenom0tabledf$Percentage <- (sat16cdenom0tabledf$Freq / sat16cdenom0sum) * 100
weightsat16cdenom0tabledf <- data.frame(Party = weightsat16cdenom0table1$x, Frequency = sat16cdenom0tabledf$Freq, Percentage_Weighted = weightsat16cdenom0table1$percentage, Percentage_Unweighted = sat16cdenom0tabledf$Percentage)

weightsat16cdenom1table1 <- wtd.table(complete$c_sat16_dv4[complete$c_denomination_dv3 == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$c_denomination_dv3 == 1])
weightsat16cdenom1sum <- sum(weightsat16cdenom1table1$sum.of.weights)
weightsat16cdenom1table1$percentage <- (weightsat16cdenom1table1$sum.of.weights / weightsat16cdenom1sum) * 100
weightsat16cdenom1table1df <- as.data.frame(weightsat16cdenom1table1)
sat16cdenom1table <- table(complete$c_sat16_dv4[complete$c_denomination_dv3 == 1])
sat16cdenom1tabledf <- as.data.frame(sat16cdenom1table)
sat16cdenom1sum <- sum(sat16cdenom1tabledf$Freq)
sat16cdenom1tabledf$Percentage <- (sat16cdenom1tabledf$Freq / sat16cdenom1sum) * 100
weightsat16cdenom1tabledf <- data.frame(Frequency = sat16cdenom1tabledf$Freq, Percentage_Weighted = weightsat16cdenom1table1$percentage, Percentage_Unweighted = sat16cdenom1tabledf$Percentage)

weightsat16cdenom2table1 <- wtd.table(complete$c_sat16_dv4[complete$c_denomination_dv3 == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$c_denomination_dv3 == 2])
weightsat16cdenom2sum <- sum(weightsat16cdenom2table1$sum.of.weights)
weightsat16cdenom2table1$percentage <- (weightsat16cdenom2table1$sum.of.weights / weightsat16cdenom2sum) * 100
weightsat16cdenom2table1df <- as.data.frame(weightsat16cdenom2table1)
sat16cdenom2table <- table(complete$c_sat16_dv4[complete$c_denomination_dv3 == 2])
sat16cdenom2tabledf <- as.data.frame(sat16cdenom2table)
sat16cdenom2sum <- sum(sat16cdenom2tabledf$Freq)
sat16cdenom2tabledf$Percentage <- (sat16cdenom2tabledf$Freq / sat16cdenom2sum) * 100
weightsat16cdenom2tabledf <- data.frame(Frequency = sat16cdenom2tabledf$Freq, Percentage_Weighted = weightsat16cdenom2table1$percentage, Percentage_Unweighted = sat16cdenom2tabledf$Percentage)

weightsat16cdenom3table1 <- wtd.table(complete$c_sat16_dv4[complete$c_denomination_dv3 == 3], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$c_denomination_dv3 == 3])
weightsat16cdenom3sum <- sum(weightsat16cdenom3table1$sum.of.weights)
weightsat16cdenom3table1$percentage <- (weightsat16cdenom3table1$sum.of.weights / weightsat16cdenom3sum) * 100
weightsat16cdenom3table1df <- as.data.frame(weightsat16cdenom3table1)
sat16cdenom3table <- table(complete$c_sat16_dv4[complete$c_denomination_dv3 == 3])
sat16cdenom3tabledf <- as.data.frame(sat16cdenom3table)
sat16cdenom3sum <- sum(sat16cdenom3tabledf$Freq)
sat16cdenom3tabledf$Percentage <- (sat16cdenom3tabledf$Freq / sat16cdenom3sum) * 100
weightsat16cdenom3tabledf <- data.frame(Frequency = sat16cdenom3tabledf$Freq, Percentage_Weighted = weightsat16cdenom3table1$percentage, Percentage_Unweighted = sat16cdenom3tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16cdenom0tabledf, startCol = 1, startRow = 34)
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16cdenom1tabledf, startCol = 6, startRow = 34)
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16cdenom2tabledf, startCol = 10, startRow = 34)
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16cdenom3tabledf, startCol = 14, startRow = 34)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

#wave 4
weightsat16ddenom0table1 <- wtd.table(complete$d_sat16_dv4[complete$d_denomination_dv3 == 0], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$d_denomination_dv3 == 0])
weightsat16ddenom0sum <- sum(weightsat16ddenom0table1$sum.of.weights)
weightsat16ddenom0table1$percentage <- (weightsat16ddenom0table1$sum.of.weights / weightsat16ddenom0sum) * 100
weightsat16ddenom0table1df <- as.data.frame(weightsat16ddenom0table1)
sat16ddenom0table <- table(complete$d_sat16_dv4[complete$d_denomination_dv3 == 0])
sat16ddenom0tabledf <- as.data.frame(sat16ddenom0table)
sat16ddenom0sum <- sum(sat16ddenom0tabledf$Freq)
sat16ddenom0tabledf$Percentage <- (sat16ddenom0tabledf$Freq / sat16ddenom0sum) * 100
weightsat16ddenom0tabledf <- data.frame(Party = weightsat16ddenom0table1$x, Frequency = sat16ddenom0tabledf$Freq, Percentage_Weighted = weightsat16ddenom0table1$percentage, Percentage_Unweighted = sat16ddenom0tabledf$Percentage)

weightsat16ddenom1table1 <- wtd.table(complete$d_sat16_dv4[complete$d_denomination_dv3 == 1], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$d_denomination_dv3 == 1])
weightsat16ddenom1sum <- sum(weightsat16ddenom1table1$sum.of.weights)
weightsat16ddenom1table1$percentage <- (weightsat16ddenom1table1$sum.of.weights / weightsat16ddenom1sum) * 100
weightsat16ddenom1table1df <- as.data.frame(weightsat16ddenom1table1)
sat16ddenom1table <- table(complete$d_sat16_dv4[complete$d_denomination_dv3 == 1])
sat16ddenom1tabledf <- as.data.frame(sat16ddenom1table)
sat16ddenom1sum <- sum(sat16ddenom1tabledf$Freq)
sat16ddenom1tabledf$Percentage <- (sat16ddenom1tabledf$Freq / sat16ddenom1sum) * 100
weightsat16ddenom1tabledf <- data.frame(Frequency = sat16ddenom1tabledf$Freq, Percentage_Weighted = weightsat16ddenom1table1$percentage, Percentage_Unweighted = sat16ddenom1tabledf$Percentage)

weightsat16ddenom2table1 <- wtd.table(complete$d_sat16_dv4[complete$d_denomination_dv3 == 2], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$d_denomination_dv3 == 2])
weightsat16ddenom2sum <- sum(weightsat16ddenom2table1$sum.of.weights)
weightsat16ddenom2table1$percentage <- (weightsat16ddenom2table1$sum.of.weights / weightsat16ddenom2sum) * 100
weightsat16ddenom2table1df <- as.data.frame(weightsat16ddenom2table1)
sat16ddenom2table <- table(complete$d_sat16_dv4[complete$d_denomination_dv3 == 2])
sat16ddenom2tabledf <- as.data.frame(sat16ddenom2table)
sat16ddenom2sum <- sum(sat16ddenom2tabledf$Freq)
sat16ddenom2tabledf$Percentage <- (sat16ddenom2tabledf$Freq / sat16ddenom2sum) * 100
weightsat16ddenom2tabledf <- data.frame(Frequency = sat16ddenom2tabledf$Freq, Percentage_Weighted = weightsat16ddenom2table1$percentage, Percentage_Unweighted = sat16ddenom2tabledf$Percentage)

weightsat16ddenom3table1 <- wtd.table(complete$d_sat16_dv4[complete$d_denomination_dv3 == 3], weights = complete$longitudinal_weight_1_w1_w2_w3_w4[complete$d_denomination_dv3 == 3])
weightsat16ddenom3sum <- sum(weightsat16ddenom3table1$sum.of.weights)
weightsat16ddenom3table1$percentage <- (weightsat16ddenom3table1$sum.of.weights / weightsat16ddenom3sum) * 100
weightsat16ddenom3table1df <- as.data.frame(weightsat16ddenom3table1)
sat16ddenom3table <- table(complete$d_sat16_dv4[complete$d_denomination_dv3 == 3])
sat16ddenom3tabledf <- as.data.frame(sat16ddenom3table)
sat16ddenom3sum <- sum(sat16ddenom3tabledf$Freq)
sat16ddenom3tabledf$Percentage <- (sat16ddenom3tabledf$Freq / sat16ddenom3sum) * 100
weightsat16ddenom3tabledf <- data.frame(Frequency = sat16ddenom3tabledf$Freq, Percentage_Weighted = weightsat16ddenom3table1$percentage, Percentage_Unweighted = sat16ddenom3tabledf$Percentage)

wb <- loadWorkbook("Dissertation Tables.xlsx")
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16ddenom0tabledf, startCol = 1, startRow = 44)
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16ddenom1tabledf, startCol = 6, startRow = 44)
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16ddenom2tabledf, startCol = 10, startRow = 44)
writeData(wb, sheet = "Denomination dv4 Tables", x = weightsat16ddenom3tabledf, startCol = 14, startRow = 44)
saveWorkbook(wb, "Dissertation Tables.xlsx", overwrite = TRUE)

