list.of.packages <- c("data.table","openxlsx")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)
lapply(list.of.packages, require, character.only=T)

setwd("~/ubos")

starts_with_x = function(x){
  return(tolower(substr(x,1,1))=="x" | is.na(x))
}

saps = list.files(pattern="SAP*")

name_list = list()
name_index = 1
for(sap_filename in saps){
  message(sap_filename)
  wb <- loadWorkbook(sap_filename)
  sheet_names = sheets(wb)
  for(sheet_name in sheet_names){
    tmp_dat_names = read.xlsx(wb, sheet=sheet_name, check.names=F, fillMergedCells=T)
    blank_names = names(tmp_dat_names)[which(starts_with_x(names(tmp_dat_names)))]
    pull_row = 1
    while(length(blank_names)>0){
      names(tmp_dat_names)[which(starts_with_x(names(tmp_dat_names)))] = tmp_dat_names[pull_row,which(starts_with_x(names(tmp_dat_names)))]
      blank_names = names(tmp_dat_names)[which(starts_with_x(names(tmp_dat_names)))]
      pull_row = pull_row + 1
    }
    tmp_dat_names = subset(tmp_dat_names,!is.na(Subcounty))
    colwise_non_na_count = colSums(!is.na(tmp_dat_names))
    tmp_name_df = data.frame(sheet=sheet_name, variable_name=names(tmp_dat_names), parish_count=colwise_non_na_count)
    name_list[[name_index]] = tmp_name_df
    name_index = name_index + 1
  }
}

name_df = rbindlist(name_list)
name_df = name_df[,.(parish_count=sum(parish_count)),by=.(sheet,variable_name)]
fwrite(name_df,"all_census_columns.csv")
