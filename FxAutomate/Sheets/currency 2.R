install.packages("quantmod")
install.packages("xlsx")
install.packages("dplyr")
install.packages("tibble")
install.packages("jsonlite")
install.packages("rJava")
install.packages("openxlsx")



library(quantmod)
library(xlsx)
library(dplyr)
library(tibble)
library(jsonlite)
library(rJava)
library(openxlsx)





####### source: Oanda mid-day rates 180 days history only ###########
CUR<-c('USD/GBP','USD/JPY','USD/CAD','USD/AUD','USD/BRL','USD/MXN','USD/EUR','USD/TWD','USD/CNY','USD/INR','USD/KRW','USD/RUB','USD/TRY','EUR/GBP')
getFX(Currencies=CUR,from='2020-06-01')
df<-merge(USDGBP,USDJPY,USDCAD,USDAUD,USDBRL,USDMXN,USDEUR,USDTWD,USDCNY,USDINR,USDKRW,USDRUB,USDTRY,EURGBP)
df<-as.data.frame(df)
df<-add_column(df, date = rownames(df), .before = 1)



# append to history data
table1<-read.xlsx("F:/MBR/MISC/fx.xlsx",sheetName = "fx",stringsAsFactors = FALSE)
table1<-table1[,-1]
df<-rbind(df,table1)
df<-distinct(df)



# save
write.xlsx(df,"F:/MBR/MISC/fx.xlsx",sheetName="fx",col.names = TRUE, row.names = TRUE, append = FALSE)
