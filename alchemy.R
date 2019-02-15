#SET WORKING DIRECTORY AND LOAD LIBRARIES
setwd("/Users/larsnygaard/Tabell_alchemy")
library(eikonapir)
library(rmarkdown)
library(tidyverse)
library(DT)

#READ XLS
library(xlsx)
file_bjorn <- read.xlsx("_Aggregation_200_Universe.xlsx",2, colIndex=2)
ticker_char <- as.character(file_bjorn$Updated.at.14.24.27)
ticker_char <- substring(ticker_char, 2)
set_app_id('F3AEA95748854609FA504255')

data_alc = get_data(as.list(ticker_char), list("TR.CommonName", "TR.TRBCEconomicSector", "TR.CompanyMarketCap","TR.ExchangeCountry",
                                                 "TR.PriceMoCountryRank","TR.Volatility5D","TR.Volatility10D","TR.Volatility20D","TR.Volatility30D",
                                                 "TR.Volatility40D","TR.Volatility50D","TR.Volatility60D","TR.Volatility80D","TR.Volatility100D",
                                                 "TR.Volatility120D","TR.Volatility150D","TR.Volatility180D","TR.Volatility240D","TR.PriceAvg5D",
                                                 "TR.PriceAvg10D","TR.PriceAvg20D","TR.PriceAvg30D","TR.PriceAvg40D","TR.PriceAvg50D","TR.PriceAvg60D",
                                                 "TR.PriceAvg80D","TR.PriceAvg100D","TR.PriceAvg120D","TR.PriceAvg150D","TR.PriceAvg180D",
                                                 "TR.PriceAvg200D","TR.PriceAvg240D","TR.PricePctChgOver50DayAvg","TR.DirMovIdxDiMinus",
                                                 "TR.DirMovIdxDiPlus", "TR.AvgDirMovIdxRating14D","TR.BollingerUpBand","TR.BollingerMidBand",
                                                 "TR.BollingerLowBand","TR.MovAvgCDSignal","TR.PriceClose","TR.PriceAvgNetDiff50D",
                                                 "TR.PriceAvgNetDiff200D"))

data_alc[c(4,6:33,35:ncol(data_alc))] <- apply(data_alc[c(4,6:33,35:ncol(data_alc))], 2, as.numeric )
data_alc <- as.tibble(data_alc)
data_alc["5D<10D VOL"] <- data_alc$`Volatility - 5 days` < data_alc$`Volatility - 10 days`
data_alc["10D<20D VOL"] <- data_alc$`Volatility - 10 days` < data_alc$`Volatility - 20 days`
data_alc["20D<30D VOL"] <- data_alc$`Volatility - 20 days` < data_alc$`Volatility - 30 days`
data_alc["30D<40D VOL"] <- data_alc$`Volatility - 30 days` < data_alc$`Volatility - 40 days`
data_alc["40D<50D VOL"] <- data_alc$`Volatility - 40 days` < data_alc$`Volatility - 50 days`
data_alc["50D<60D VOL"] <- data_alc$`Volatility - 50 days` < data_alc$`Volatility - 60 days`
data_alc["60D<80D VOL"] <- data_alc$`Volatility - 60 days` < data_alc$`Volatility - 80 days`
data_alc["80D<100D VOL"] <- data_alc$`Volatility - 80 days` < data_alc$`Volatility - 100 days`
data_alc["100D<120D VOL"] <- data_alc$`Volatility - 100 days` < data_alc$`Volatility - 120 days`
data_alc["120D<150D VOL"] <- data_alc$`Volatility - 120 days` < data_alc$`Volatility - 150 days`
data_alc["150D<180D VOL"] <- data_alc$`Volatility - 150 days` < data_alc$`Volatility - 180 days`
data_alc["180D<240D VOL"] <- data_alc$`Volatility - 180 days` < data_alc$`Volatility - 240 days`

data_alc["5D<10D SMA"] <- data_alc$`5-day SMA` < data_alc$`10-day SMA`
data_alc["10D<20D SMA"] <- data_alc$`10-day SMA` < data_alc$`20-day SMA`
data_alc["20D<30D SMA"] <- data_alc$`20-day SMA` < data_alc$`30-day SMA`
data_alc["30D<40D SMA"] <- data_alc$`30-day SMA` < data_alc$`40-day SMA`
data_alc["40D<50D SMA"] <- data_alc$`40-day SMA` < data_alc$`50-day SMA`
data_alc["50D<60D SMA"] <- data_alc$`50-day SMA` < data_alc$`60-day SMA`
data_alc["60D<80D SMA"] <- data_alc$`60-day SMA` < data_alc$`80-day SMA`
data_alc["80D<100D SMA"] <- data_alc$`80-day SMA` < data_alc$`100-day SMA`
data_alc["100D<120D SMA"] <- data_alc$`100-day SMA` < data_alc$`120-day SMA`
data_alc["120D<150D SMA"] <- data_alc$`120-day SMA` < data_alc$`150-day SMA`
data_alc["150D<180D SMA"] <- data_alc$`150-day SMA` < data_alc$`180-day SMA`
data_alc["180D<240D SMA"] <- data_alc$`180-day SMA` < data_alc$`240-day SMA`
  
sub_vol <- data_alc[c("5D<10D VOL","10D<20D VOL","20D<30D VOL","30D<40D VOL","40D<50D VOL","50D<60D VOL","60D<80D VOL","80D<100D VOL","100D<120D VOL","120D<150D VOL","150D<180D VOL","180D<240D VOL")]
data_alc["VOL DECREASE"] <- apply(sub_vol, 1, function(x) length(x[x==FALSE]))
data_alc["VOL INCREASE"] <- apply(sub_vol, 1, function(x) length(x[x==TRUE]))
data_alc["VOL IMPROVEMENT"] <- data_alc["VOL DECREASE"] - data_alc["VOL INCREASE"]

sub_sma <- data_alc[c("5D<10D SMA","10D<20D SMA","20D<30D SMA","30D<40D SMA","40D<50D SMA","50D<60D SMA","60D<80D SMA","80D<100D SMA","100D<120D SMA","120D<150D SMA","150D<180D SMA","180D<240D SMA")]
data_alc["SMA DECREASE"] <- apply(sub_sma, 1, function(x) length(x[x==FALSE]))
data_alc["SMA INCREASE"] <- apply(sub_sma, 1, function(x) length(x[x==TRUE]))
data_alc["SMA IMPROVEMENT"] <- data_alc["SMA DECREASE"] - data_alc["SMA INCREASE"]
data_alc["COMBO"] <- data_alc["VOL IMPROVEMENT"] + data_alc["SMA IMPROVEMENT"]
data_alc["HIGH VOL REGIME"] <- 2*data_alc["VOL IMPROVEMENT"] + data_alc["SMA IMPROVEMENT"]
data_alc["LOW VOL REGIME"] <- data_alc["VOL IMPROVEMENT"] + 2*data_alc["SMA IMPROVEMENT"]

data_alc["NET DMI"] <- data_alc["DMI- Positive Directional Indicator"] + data_alc["DMI- Negative Directional Indicator"]
data_alc["ADX<DMI"] <- data_alc["ADX Rating - 14 Day"] + data_alc["DMI- Negative Directional Indicator"]
data_alc["BOLL>CLS"] <- data_alc["Bollinger Middle Band"] > data_alc["Price Close"]
data_alc["CLS>50D"] <- data_alc["Close Price vs 50-day SMA Net Change"] > 0
data_alc["CLS>200D"] <- data_alc["Close Price vs 200-day SMA Net Change"] > 0
sub_ma_count <- data_alc[c("CLS>50D","CLS>200D")]
data_alc["MA COUNT"] <- apply(sub_ma_count, 1, function(x) length(x[x==TRUE]))


stats <- data.frame()
#Universe
stats[1,1] <- mean(data_alc$"SMA IMPROVEMENT")
stats[2,1] <- mean(data_alc$"VOL IMPROVEMENT")
stats[3,1] <- mean(data_alc$"COMBO")
stats[4,1] <- mean(data_alc$"HIGH VOL REGIME")
stats[5,1] <- mean(data_alc$"LOW VOL REGIME")
stats[6,1] <- mean(data_alc$"NET DMI")
stats[7,1] <- mean(data_alc$"MA COUNT")
row.names(stats) <- c("SMA IMPROVEMENT","VOL IMPROVEMENT","COMBO","HIGH VOL REGIME","LOW VOL REGIME","NET DMI","MA COUNT")


sub_universe_countries <- c("Denmark","Finland","Norway","Sweden")
j=2
for(i in sub_universe_countries){
  print(i)
  data_sub_universe <- subset(data_alc, data_alc$`Country of Exchange` == i)
  stats[1,j] <- mean(data_sub_universe$"SMA IMPROVEMENT")
  stats[2,j] <- mean(data_sub_universe$"VOL IMPROVEMENT")
  stats[3,j] <- mean(data_sub_universe$"COMBO")
  stats[4,j] <- mean(data_sub_universe$"HIGH VOL REGIME")
  stats[5,j] <- mean(data_sub_universe$"LOW VOL REGIME")
  stats[6,j] <- mean(data_sub_universe$"NET DMI")
  stats[7,j] <- mean(data_sub_universe$"MA COUNT")
j=j+1
}

sub_universe_sectors <- c("Healthcare","Consumer Cyclicals","Energy","Financials","Industrials","Consumer Non-Cyclicals","Utilities","Basic Materials","Telecommunications Services","Technology")
#Universe
for(i in sub_universe_sectors){
  print(i)
  data_sub_universe <- subset(data_alc, data_alc$`TRBC Economic Sector Name` == i)
  stats[1,j] <- mean(data_sub_universe$"SMA IMPROVEMENT")
  stats[2,j] <- mean(data_sub_universe$"VOL IMPROVEMENT")
  stats[3,j] <- mean(data_sub_universe$"COMBO")
  stats[4,j] <- mean(data_sub_universe$"HIGH VOL REGIME")
  stats[5,j] <- mean(data_sub_universe$"LOW VOL REGIME")
  stats[6,j] <- mean(data_sub_universe$"NET DMI")
  stats[7,j] <- mean(data_sub_universe$"MA COUNT")
  j=j+1
}
colnames(stats) <- c("Universe",sub_universe_countries,sub_universe_sectors)





###########################################################################
##############################  RENDER SITE  ##############################
###########################################################################
rmarkdown::render_site()
#END
