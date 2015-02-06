## Automating the download of DOR Taxes

##Where I want the files
setwd("C:\\Users\\dramirez\\Desktop\\Projects\\Datasets\\DOR Retail")

fileLink <- "http://www.dornc.com/publications/monthly_sales_"

# The beginning of the file link is the same, and the file is the same format
# monthly_sales_m-yy.xls
# the month is single for 1-9 and double 10-12

# Here are all the years, incase you need to subset a few.
# years <- c("00","01","02","03","04","05","06","07","08","09","10","11","12","13","14")
years <- c("02","03","04","05","06","07","08","09","10","11","12","13","14")
months <- as.character(c(1:12))

## This is the function to download the files.  It's easier to see and edit if it's
## separate.
getExcel <- function(mth,yr){
  URL <- paste(fileLink,mth,"-",yr,".xls",sep="")
  destFile <- paste("monthly_sales_",yr,"_",mth,".xls",sep="")
  download.file(URL,destFile,mode="wb")
}

## Loops through the excel files and puts the results in a list.
results <- lapply(years,function(x) lapply(months,function(y) try(getExcel(y,x))))

