# Load Packages

library(googledrive)
library(openxlsx)
library(readxl)
library(zoo)
library(stringr)
library(RDCOMClient)
library(xtable)
library(plyr)
library(tibble)

# Load function script
options(scipen = 999)
setwd('C:/File/MCC script')
source('AMJ Functions.R')

# Authenticate google drive API
drive_auth()


## Re-organize Amazon files
Amazon <- as.vector(unlist(drive_ls('1. Buy Details/Amazon/AMJ',recursive = T, pattern = "xlsm")[2]))
setwd('C:/File/Buy Details/Amazon')
do.call(file.remove, list(list.files()))
sapply(1:length(Amazon), function(x) drive_download(as_id(Amazon[x]), overwrite = TRUE))
Amazon.upload <- list.files()
Amazon.upload <- Amazon.upload[!grepl('Archive',Amazon.upload)]
sapply(1:length(Amazon.upload), function(x) drive_upload(Amazon.upload[x],as_dribble('For Analytics - New Format')$path))

## Re-organize Amazon Cross Quarter files
Amazon <- as.vector(unlist(drive_ls('1. Buy Details/Amazon/Cross-Quarter',recursive = T, pattern = "xlsm")[2]))
setwd('C:/File/Buy Details/Amazon Cross Quarter New')
do.call(file.remove, list(list.files()))
sapply(1:length(Amazon), function(x) drive_download(as_id(Amazon[x]), overwrite = TRUE))
Amazon.upload <- list.files()
Amazon.upload <- Amazon.upload[!grepl('Archive',Amazon.upload)]
sapply(1:length(Amazon.upload), function(x) drive_upload(Amazon.upload[x],as_dribble('For Analytics - Cross Quarter')$path))

# Download All buy details
GD.HVR <- '4. AMJ FY20/HVR'
GD.Clorox <- '4. AMJ FY20/Clorox Studio'
GD.CES <- '4. AMJ FY20/CES'
GD.Burts <- '4. AMJ FY20/Burts Bees'
GD.Kingsford <- '4. AMJ FY20/Kingsford'
GD.Cats <- '4. AMJ FY20/Cats'
GD.PPD <- '4. AMJ FY20/PPD'
GD.RL <- '4. AMJ FY20/Renew Life'
GD.Glad <- '4. AMJ FY20/Glad'
GD.Amazon <- 'For Analytics - New Format'
GD.Nutranext <- '4. AMJ FY20/NXT'
GD.CQ <- 'For Analytics - Cross Quarter'

all.brand<- c(GD.Kingsford,GD.Burts,GD.Cats,GD.Clorox,GD.Glad,GD.HVR,GD.Nutranext,GD.PPD,GD.RL,GD.Amazon,GD.CES)
all.brand.name<-c('Kingsford','Burts','Cats','Clorox Studio','Glad','HVR','Nutranext','PPD','Renew Life','Amazon','CES')

for (i in 1:length(all.brand)){
  if(i==1){
    reference.info <- brand.reference(all.brand.name[i],all.brand[i])
  } else {
    temp <- brand.reference(all.brand.name[i],all.brand[i])
    reference.info <- rbind(reference.info, temp)}
}
colnames(reference.info) <- c('Buy Details','Brand')

setwd('C:/File/Buy Details/AMJ/')
do.call(file.remove, list(list.files()))
sapply(1:length(all.brand), function(x) buy.detail.download(all.brand[x]))

# Combine files and map brand info
all.files <- list.files(pattern = '.xlsm')
all.buy.details.AMJ <- table.merge(all.files)
all.buy.details.AMJ.copy <- all.buy.details.AMJ
all.buy.details.AMJ <- join(all.buy.details.AMJ, reference.info, by = "Buy Details")

# Generate placement name
for (i in 1:nrow(all.buy.details.AMJ)){
  if((all.buy.details.AMJ$`Placement Type`[i] == 'Package') | (all.buy.details.AMJ$Width[i] == 'PKG')){
    all.buy.details.AMJ$Size[i] <- 'PKG'
  } else if(tolower(all.buy.details.AMJ$Width[i]) == 'vast'){
    all.buy.details.AMJ$Size[i] <- '0 x 0'
  } 
  else {
    all.buy.details.AMJ$Size[i] <- paste(all.buy.details.AMJ$Width[i], all.buy.details.AMJ$Height[i], sep = ' x ')
  }
}
all.buy.details.AMJ$`DCM Placement Name`<- paste(
  all.buy.details.AMJ$`Line Item`,
  all.buy.details.AMJ$Geo,
  all.buy.details.AMJ$`Site Name`,
  all.buy.details.AMJ$`audience + placement name`,
  all.buy.details.AMJ$Vehicle,
  all.buy.details.AMJ$`Cost Structure`,
  all.buy.details.AMJ$`Campaign ID`,
  all.buy.details.AMJ$`Inventory Source`,
  all.buy.details.AMJ$`Targetin WHO`,
  all.buy.details.AMJ$Size,
  all.buy.details.AMJ$`Site Served or Dart`,
  sep = '|'
)

# Data clean up
all.buy.details.AMJ$`Start Date` <- do.call(date.convert, list(all.buy.details.AMJ$`Start Date`))
all.buy.details.AMJ$`End Date` <- do.call(date.convert, list(all.buy.details.AMJ$`End Date`))
all.buy.details.AMJ$`Campaign ID` <- gsub("\\D", "", all.buy.details.AMJ$`Campaign ID`)
all.buy.details.AMJ$`Note - Missing Campaign ID`<-''
all.buy.details.AMJ$`Note - Missing Placement ID`<-''
all.buy.details.AMJ$`Note - Incorrect Start Date`<-''
all.buy.details.AMJ$`Note - Missing Tag Type`<-''
all.buy.details.AMJ$`Note - Zero IO Rate`<-''
all.buy.details.AMJ$`Note - Missing Placement ID`[which(all.buy.details.AMJ$`DCM Placement ID`=="")]<-'x'
all.buy.details.AMJ$`DCM Placement ID`[which(all.buy.details.AMJ$`DCM Placement ID`=="")]<-'Missing Placement ID'
all.buy.details.AMJ$`Note - Missing Campaign ID`[which(all.buy.details.AMJ$`Campaign ID`=="")]<-'x'
all.buy.details.AMJ$`Campaign ID`[which(all.buy.details.AMJ$`Campaign ID`=="")]<-'Missing Campaign ID'
all.buy.details.AMJ$`Note - Incorrect Start Date`[which((all.buy.details.AMJ$`Start Date` > '2020-06-29' |
                                                           all.buy.details.AMJ$`Start Date` < '2020-03-30') &
                                                          all.buy.details.AMJ$Brand != 'Amazon Cross Quarter')]<-'x'
all.buy.details.AMJ$`Note - Missing Tag Type`[which(all.buy.details.AMJ$`Adserving Fees - Tag Type`=='' & 
                                                      all.buy.details.AMJ$Width != '2')]<-'x'
all.buy.details.AMJ$`Adserving Fees - Tag Type`[which(all.buy.details.AMJ$`Adserving Fees - Tag Type`=='' &
                                                        all.buy.details.AMJ$Width != '2')]<-'Missing Tag Type'
all.buy.details.AMJ$`Note - Zero IO Rate`[which(
  (all.buy.details.AMJ$`Net/Gross Rate` == '' & (tolower(all.buy.details.AMJ$`Cost Structure`) != 'vadd' & tolower(all.buy.details.AMJ$`Cost Structure`) != 'flat rate - impressions') & all.buy.details.AMJ$`Placement Type` == 'Package') |
    (all.buy.details.AMJ$`Net/Gross Rate` == '0' & (tolower(all.buy.details.AMJ$`Cost Structure`) != 'vadd' & tolower(all.buy.details.AMJ$`Cost Structure`) != 'flat rate - impressions') & all.buy.details.AMJ$`Placement Type` == 'Package'))]<-'x'
all.buy.details.AMJ$`Net/Gross Rate`[which((all.buy.details.AMJ$`Net/Gross Rate` == '' & (tolower(all.buy.details.AMJ$`Cost Structure`) != 'vadd' & tolower(all.buy.details.AMJ$`Cost Structure`) != 'flat rate - impressions') & all.buy.details.AMJ$`Placement Type` == 'Package') |
                                             (all.buy.details.AMJ$`Net/Gross Rate` == '0' & (tolower(all.buy.details.AMJ$`Cost Structure`) != 'vadd' & tolower(all.buy.details.AMJ$`Cost Structure`) != 'flat rate - impressions') & all.buy.details.AMJ$`Placement Type` == 'Package'))]<-'Zero Cost Type'
all.buy.details.AMJ$`Cost Structure`[which(all.buy.details.AMJ$`Cost Structure` == 'CPE')]<-'CPA'                                              (all.buy.details.AMJ$`Net/Gross Rate` == '0' & all.buy.details.AMJ$`Cost Structure` != 'VADD')) ,]
problematic.rows <- all.buy.details.AMJ[which(
  all.buy.details.AMJ$`Note - Incorrect Start Date` == 'x' |
    all.buy.details.AMJ$`Note - Missing Campaign ID` == 'x' |
    all.buy.details.AMJ$`Note - Missing Placement ID` == 'x' |
    all.buy.details.AMJ$`Note - Missing Tag Type` == 'x' |
    all.buy.details.AMJ$`Note - Zero IO Rate` == 'x'
),]

# Generate problematic data
problematic.rows$`GROSS UNITS`[which(problematic.rows$`GROSS UNITS`=='')]<-0
problematic.rows$`TOTAL COST`[which(problematic.rows$`TOTAL COST`=='')]<-0
problematic.rows$`GROSS UNITS`<-as.numeric(problematic.rows$`GROSS UNITS`)
problematic.rows$`TOTAL COST`<-as.numeric(problematic.rows$`TOTAL COST`)
                                         (all.buy.details.AMJ$`Net/Gross Rate` == '0' & (all.buy.details.AMJ$`Cost Structure` != 'VADD' | all.buy.details.AMJ$`Cost Structure` != 'Flat Rate - Impressions'))),]
nrow(problematic.rows)

# output
output <- all.buy.details.AMJ
problematic.files <- problematic.rows
setwd('C:/File/Buy Details Output/AMJ/')
file.remove(list.files())
write.csv(output, 'Buy Detail Master File.csv', row.names = F)
write.csv(problematic.files, 'Problematic Files.csv', row.names = F)

# send to Datorama
OutApp <- COMCreate("Outlook.Application")
outMail = OutApp$CreateItem(0)
outMail[["To"]] = 'xxxx@xxxx.com'
outMail[["subject"]] = "Clorox AMJ Buy Details Upload"
outMail[["body"]] = "Buy Details Upload"
outMail[["Attachments"]]$Add('C:\\File\\Buy Details Output\\AMJ\\Buy Detail Master File.csv')
outMail$Send()

# Send diagnosis email with problematic data, if any
if(nrow(problematic.rows)>0){
  problematic.files <- join(problematic.rows, reference.info, by = "Buy Details")
  bad.buy.details <- unique(problematic.files$`Buy Details`)
  email.body <- data.frame(Index = index(1:length(bad.buy.details)), `Buy Details` = bad.buy.details)
  MCC <- paste("people1@xxx.com","people2@xxx.com", sep = ";", collapse=NULL)
  OutApp <- COMCreate("Outlook.Application")
  outMail = OutApp$CreateItem(0)
  outMail[["To"]] = MCC
  outMail[["subject"]] = "AMJ Buy Details Issues"
  y <- print(xtable(email.body), type="html", print.results=FALSE, include.rownames=FALSE)
  outMail[["HTMLbody"]] = paste0("Below are AMJ Buy Details with issues. Please check columns (BR - BV) respectively from the attached file.","<html>", y, "</html>")
  outMail[["Attachments"]]$Add('C:\\File\\Buy Details Output\\AMJ\\Problematic Files.csv')
  
  outMail$Send()
}

