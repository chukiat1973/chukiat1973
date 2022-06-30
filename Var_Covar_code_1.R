

####################### Manipulate data for VaR Optimization  ################
library("readxl")
library("mvtnorm")
library("graphics")


####### Import data from excel file ######
dat<- read_excel("C:/Users/Mac/Desktop/Close_price_2020_2022.xlsx", sheet = 6)
dat

#############  Plot time series by r command very good ##############################
### https://stack overflow.com/questions/26076262/issues-plotting-multivariate-time-series-in-r
require(graphics)
library(zoo)
library(ggplot2)
matplot(dat, type = c("l"),pch=1,col = 1:16, lwd = 1,lty = 1) #plot
legend("topright", legend = 1:16, col=1:16, pch=1) # optional legend


###########################################################################

SP500<-mean(dat$SP500)
DJI<-mean(dat$DJI)
NASDAQ<-mean(dat$NASDAQ)
NYSE<-mean(dat$NYSE)
BUK100P<-mean(dat$BUK100P)
CAC40<-mean(dat$CAC40)
ESTX<-mean(dat$ESTX)
FTSE100<-mean(dat$FTSE100)
Nikkei225<-mean(dat$Nikkei225)
Hang<-mean(dat$Hang)
SP_Boom<-mean(dat$SP_Bom)
Shenz<-mean(dat$Shenz)
SET<-mean(dat$SET)
FTSE_Sin<-mean(dat$FTSE_Sin)
FTSE_Ma<-mean(dat$FTSE_Ma)
JKSE<-mean(dat$JKSE)

SP500_v<-var(dat$SP500)
DJI_v<-var(dat$DJI)
NASDAQ_v<-var(dat$NASDAQ)
NYSE_v<-var(dat$NYSE)
BUK100P_v<-var(dat$BUK100P)
CAC40_v<-var(dat$CAC40)
ESTX_v<-var(dat$ESTX)
FTSE100_v<-var(dat$FTSE100)
Nikkei225_v<-var(dat$Nikkei225)
Hang_v<-var(dat$Hang)
SP_Boom_v<-var(dat$SP_Bom)
Shenz_v<-var(dat$Shenz)
SET_v<-var(dat$SET)
FTSE_Sin_v<-var(dat$FTSE_Sin)
FTSE_Ma_v<-var(dat$FTSE_Ma)
JKSE_v<-var(dat$JKSE)


mu<-data.frame(SP500, DJI, NASDAQ, NYSE, BUK100P, CAC40, ESTX, FTSE100, Nikkei225, 
               Hang, SP_Boom, Shenz, SET, FTSE_Sin, FTSE_Ma, JKSE)
mu

var<-data.frame(SP500_v, DJI_v, NASDAQ_v, NYSE_v, BUK100P_v, CAC40_v, ESTX_v
                , FTSE100_v, Nikkei225_v, Hang_v, SP_Boom_v, Shenz_v, SET_v, 
                FTSE_Sin_v, FTSE_Ma_v, JKSE_v)
var

mu_sm     <- colMeans(dat)
Sigma_scm <- cov(dat)

covar <- data.frame(Sigma_scm)
##### export Mean and co-variance to excel ##########################################
### Export data to sheet of Excel ###
library(openxlsx)
write.xlsx(mu, file = "mu.xlsx", colNames = TRUE, borders = "columns")
write.xlsx(var, file = "var.xlsx", colNames = TRUE, borders = "columns")
write.xlsx(covar, file = "cov.xlsx", colNames = TRUE, borders = "columns")

##################  stock Exchange 0.00 ##########################################
####### Import data from excel file ######
dat<- read_excel("C:/Users/Mac/Desktop/Close_price_2020_2022.xlsx", sheet = 5)
dat
SP500_M<-mean(dat$SP500)
DJI_M<-mean(dat$DJI)
NASDAQ_M<-mean(dat$NASDAQ)
NYSE_M<-mean(dat$NYSE)




mu<-data.frame(SP500_M,DJI_M,NASDAQ_M,NYSE_M)
mu

dat1<-data.frame(dat$SP500, dat$DJI, dat$NASDAQ, dat$NYSE)
dat1
mu_sm     <- colMeans(dat1)
Sigma_scm <- cov(dat1)

covar <- data.frame(Sigma_scm)
##### export Mean and co-variance to excel #####################################
### Export data to sheet of Excel ###
library(openxlsx)
write.xlsx(mu, file = "mu.xlsx", colNames = TRUE, borders = "columns")
write.xlsx(covar, file = "cov.xlsx", colNames = TRUE, borders = "columns")

#####################  stock Exchange 0.25 #######################################
####### Import data from excel file ######
dat<- read_excel("C:/Users/Mac/Desktop/Close_price_2020_2022.xlsx", sheet = 5)
dat
Hang_M<-mean(dat$Hang)
Shenz_M<-mean(dat$Shenz)
BUK100P_M<-mean(dat$BUK100P)
FTSE100_M<-mean(dat$FTSE100)

mu<-data.frame(Hang_M, Shenz_M, BUK100P_M, FTSE100_M)
mu

dat1<-data.frame(dat$Hang, dat$Shenz, dat$BUK100P, dat$FTSE100)
dat1
mu_sm     <- colMeans(dat1)
Sigma_scm <- cov(dat1)

covar <- data.frame(Sigma_scm)
##### export Mean and co-variance to excel #####################################
### Export data to sheet of Excel ###
library(openxlsx)
write.xlsx(mu, file = "mu.xlsx", colNames = TRUE, borders = "columns")
write.xlsx(covar, file = "cov.xlsx", colNames = TRUE, borders = "columns")


#####################  stock Exchange 0.50 #######################################
####### Import data from excel file ######
dat<- read_excel("C:/Users/Mac/Desktop/Close_price_2020_2022.xlsx", sheet = 5)
dat
Hang_M<-mean(dat$Hang)
BUK100P_M<-mean(dat$BUK100P)
FTSE100_M<-mean(dat$FTSE100)
FTSE_Ma_M<-mean(dat$FTSE_Ma)


mu<-data.frame(Hang_M, BUK100P_M, FTSE100_M,FTSE_Ma_M)
mu

dat1<-data.frame(dat$Hang, dat$BUK100P, dat$FTSE100, dat$FTSE_Ma)
dat1
mu_sm     <- colMeans(dat1)
Sigma_scm <- cov(dat1)

covar <- data.frame(Sigma_scm)
##### export Mean and co-variance to excel #####################################
### Export data to sheet of Excel ###
library(openxlsx)
write.xlsx(mu, file = "mu.xlsx", colNames = TRUE, borders = "columns")
write.xlsx(covar, file = "cov.xlsx", colNames = TRUE, borders = "columns")

#####################  stock Exchange 0.75 #######################################
####### Import data from excel file ######
dat<- read_excel("C:/Users/Mac/Desktop/Close_price_2020_2022.xlsx", sheet = 5)
dat
Hang_M<-mean(dat$Hang)
FTSE_Ma_M<-mean(dat$FTSE_Ma)
BUK100P_M<-mean(dat$BUK100P)
FTSE100_M<-mean(dat$FTSE100)



mu<-data.frame(Hang_M, FTSE_Ma_M, BUK100P_M, FTSE100_M)
mu

dat1<-data.frame(dat$Hang, dat$FTSE_Ma, dat$BUK100P, dat$FTSE100)
dat1
mu_sm     <- colMeans(dat1)
Sigma_scm <- cov(dat1)

covar <- data.frame(Sigma_scm)
##### export Mean and co-variance to excel #####################################
### Export data to sheet of Excel ###
library(openxlsx)
write.xlsx(mu, file = "mu.xlsx", colNames = TRUE, borders = "columns")
write.xlsx(covar, file = "cov.xlsx", colNames = TRUE, borders = "columns")

#####################  stock Exchange 1.00 #######################################
####### Import data from excel file ######
dat<- read_excel("C:/Users/Mac/Desktop/Close_price_2020_2022.xlsx", sheet = 5)
dat
Hang_M<-mean(dat$Hang)
FTSE_Ma_M<-mean(dat$FTSE_Ma)
BUK100P_M<-mean(dat$BUK100P)
FTSE100_M<-mean(dat$FTSE100)



mu<-data.frame(Hang_M, FTSE_Ma_M, BUK100P_M, FTSE100_M)
mu

dat1<-data.frame(dat$Hang, dat$FTSE_Ma, dat$BUK100P, dat$FTSE100)
dat1
mu_sm     <- colMeans(dat1)
Sigma_scm <- cov(dat1)

covar <- data.frame(Sigma_scm)
##### export Mean and co-variance to excel #####################################
### Export data to sheet of Excel ###
library(openxlsx)
write.xlsx(mu, file = "mu.xlsx", colNames = TRUE, borders = "columns")
write.xlsx(covar, file = "cov.xlsx", colNames = TRUE, borders = "columns")



