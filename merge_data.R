list.of.packages <- c("data.table","openxlsx","stringr")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)
lapply(list.of.packages, require, character.only=T)

setwd("~/git/ubos-census-2014")

starts_with_x = function(x){
  return(tolower(substr(x,1,1))=="x" | is.na(x))
}

saps = list.files(pattern="SAP*")

data_list = list()
data_index = 1
for(sap_filename in saps){
  message(sap_filename)
  region = str_extract(sap_filename,"(?<=_)(.*?)(?=\\.)")
  wb <- loadWorkbook(sap_filename)
  sheet_names = sheets(wb)
  for(sheet_name in sheet_names){
    tmp_dat = read.xlsx(wb, sheet=sheet_name, check.names=F, fillMergedCells=T)
    blank_names = names(tmp_dat)[which(starts_with_x(names(tmp_dat)))]
    pull_row = 1
    while(length(blank_names)>0){
      names(tmp_dat)[which(starts_with_x(names(tmp_dat)))] = tmp_dat[pull_row,which(starts_with_x(names(tmp_dat)))]
      blank_names = names(tmp_dat)[which(starts_with_x(names(tmp_dat)))]
      pull_row = pull_row + 1
    }
    tmp_dat = subset(tmp_dat,!is.na(Subcounty))
    tmp_dat$Region = region
    tmp_dat$Sheet.name = sheet_name
    data_list[[data_index]] = tmp_dat
    data_index = data_index + 1
  }
}

dat = rbindlist(data_list,fill=T)
fwrite(dat,"all_ug_2014.csv")
