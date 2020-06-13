buy.detail.download <- function(folder.path, recursive){
  
  excel <- as.vector(unlist(drive_ls(folder.path, recursive = F, pattern = "xlsm")[2]))
  gs <- as.vector(unlist(drive_ls(folder.path, recursive = F, type = "spreadsheet")[2]))
  brand <- c(excel,gs)
  setwd('C:/File/Buy Details/AMJ')
  sapply(1:length(brand), function(x) drive_download(as_id(brand[x]), overwrite = TRUE))
}

table.merge <- function(file.list){
  for (i in 1:length(file.list)){
    print(file.list[i])
    data <- read.xlsx(file.list[i], sheet='Plan', startRow = 6)
    colnames(data)<-gsub("[.]"," ",colnames(data))
    if(('ATD Audience' %in% colnames(data)) == FALSE){
      data <- add_column(data, '', .after = "Salesforce audience")
      names(data)[25]<-'ATD Audience'
    }
    data[is.na(data)]<-''
    
    temp <- data[which(data$`View/Create/Update` != "" & data$`View/Create/Update` != "Total"),]
    temp$`Buy Details` <- file.list[i]
    if (i==1){
      all.data <- temp
    } else {
      all.data <- rbind(all.data,temp)
    }
  }
  all.data$`Campaign ID` <- trimws(all.data$`Campaign ID`)
  return (all.data)
}

date.convert <- function(date){
  return (as.Date(as.numeric(date), origin = "1899-12-30"))
}

brand.reference <- function(folder.name, folder.path){
  all.files <- drive_ls(folder.path)$name
  all.files.no.archive <- all.files[!grepl('Archive',all.files)]
  reference.info <- data.frame(`Buy Details` = all.files.no.archive, Brand = folder.name, check.names = F)
  return (reference.info)
}