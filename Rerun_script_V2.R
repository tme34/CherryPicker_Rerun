#Checking if packages need to be installed
if(!require(readxl)){
  install.packages("readxl", repos = "https://cran.rstudio.com/")
}
if(!require(openxlsx)){
  install.packages("openxlsx", repos = "https://cran.rstudio.com/")
}
if(!require(tidyverse)){
  install.packages("tidyverse", repos = "https://cran.rstudio.com/")
}

#Loading needed packages
library(readxl)
library(openxlsx)
library(tidyverse)

#Query if script is interactive
print(interactive())

if(interactive()){
  winDialog(type = c("ok"),
            "Select list of rerun entries")
}

#Choosing input file interactively (rerun file)
rerun_list<-read.xlsx(xlsxFile=choose.files(caption = "Select Rerun-list File",
                                            multi=FALSE)
                      ,sheet="Ark1",colNames = FALSE)

#Manually naming column names
colnames(rerun_list)<-c("Numb","Plate","Position","Plate-Barcode","Tube-Barcode")

#Determining number of samples and lists of acquired rerun data.

#samples = number of rows in initial rerun data.
samples<-nrow(rerun_list)
#n_list is the number of filled sheets created.
n_list<-samples %/% 95
#n_list_last is the number of entries (rows ) of the last non-full sheet.
n_list_last<-samples%%95

#If no full sheets in the data is available, set n_list to 1 anyways.

#Splitting data into lists of 95 rows + rest list.
if (samples>95){
  rest_rerun_list<-tail(rerun_list, n_list_last)
  if (n_list_last!=0){
    rerun_list<-head(rerun_list, -n_list_last)
  }
  split_list<-split(rerun_list, rep(1:n_list,each=95))
  if(n_list_last!=0){
    split_list[[(length(split_list)+1)]]<-rest_rerun_list
  }
  
} else{
  split_list<-split(rerun_list, rep(1,each=samples))
}

#n_list changed to length of split_list, since rest list is added in this version.
n_list<-length(split_list)

#Finding number of plates in all splits
new_split_list<-list()
new_plate_list<-list()
name_plate_list<-list()
new_name_list<-list()
new_work_list<-list()
new_rwork_list<-list()

#For-loop determining number of sheets needed for each 95 entries and compiling
#into a final lists of data frames used as rerun list sheets.
#K is used as an enumerator to append new rerun, work and rwork lists.

k<-0
for (i in 1:n_list){
  k<-k+1
  plates<-split_list[[i]]$Plate
  number_of_plates<-length(unique(plates))
  
  if (number_of_plates<20){
    #If there are below 20 plates in the 95 entries
    new_name_list[[k]]<-paste("Rerun_list",i,"A",sep="_")
    new_work_list[[k]]<-paste("Work_list",i,"A",sep="_")
    new_rwork_list[[k]]<-paste("Reverse_work_list",i,"A",sep="_")
    name_plate_list[[k]]<-paste("Plate_list",i,"A",sep="_")
    new_split_list[[k]]<-split_list[[i]]
    new_plate_list[[k]]<-data.frame(Plate=matrix(unique(split_list[[i]]$Plate)), 
                                    Plate_Barcode=matrix(unique(split_list[[i]]$`Plate-Barcode`)))
  }
  
  if (number_of_plates>19 && number_of_plates<39){
    #If there are above 19 plates in the 95 entries
    plate_twenty<-unique(plates)[20]
    
    row_split_B<-which(split_list[[i]]$Plate==plate_twenty)[1]
    
    split_list_A<-split_list[[i]][1:(row_split_B-1),]
    split_list_B<-split_list[[i]][row_split_B:nrow(split_list[[i]]),]
    
    new_name_list[[k]]<-paste("Rerun_list",i,"A",sep="_")
    new_work_list[[k]]<-paste("Work_list",i,"A",sep="_")
    new_rwork_list[[k]]<-paste("Reverse_work_list",i,"A",sep="_")
    name_plate_list[[k]]<-paste("Plate_list",i,"A",sep="_")
    new_split_list[[k]]<-split_list_A
    new_plate_list[[k]]<-data.frame(Plate=matrix(unique(split_list_A$Plate)), 
                                    Plate_Barcode=matrix(unique(split_list_A$`Plate-Barcode`)))
    k<-k+1
    new_name_list[[k]]<-paste("Rerun_list",i,"B",sep="_")
    new_work_list[[k]]<-paste("Work_list",i,"B",sep="_")
    new_rwork_list[[k]]<-paste("Reverse_work_list",i,"B",sep="_")
    name_plate_list[[k]]<-paste("Plate_list",i,"B",sep="_")
    new_split_list[[k]]<-split_list_B
    new_plate_list[[k]]<-data.frame(Plate=matrix(unique(split_list_B$Plate)), 
                                    Plate_Barcode=matrix(unique(split_list_B$`Plate-Barcode`)))
  }
  
  if (number_of_plates>38 && number_of_plates<58){
    #If there are above 38 plates in the 95 entries
    plate_twenty<-unique(plates)[20]
    plate_three_nine<-unique(plates)[39]
    
    row_split_B<-which(split_list[[i]]$Plate==plate_twenty)[1]
    row_split_C<-which(split_list[[i]]$Plate==plate_three_nine)[1]
    
    split_list_A<-split_list[[i]][1:(row_split_B-1),]
    split_list_B<-split_list[[i]][row_split_B:(row_split_C-1),]
    split_list_C<-split_list[[i]][row_split_C:nrow(split_list[[i]]),]
    
    new_name_list[[k]]<-paste("Rerun_list",i,"A",sep="_")
    new_work_list[[k]]<-paste("Work_list",i,"A",sep="_")
    new_rwork_list[[k]]<-paste("Reverse_work_list",i,"A",sep="_")
    name_plate_list[[k]]<-paste("Plate_list",i,"A",sep="_")
    new_split_list[[k]]<-split_list_A
    new_plate_list[[k]]<-data.frame(Plate=matrix(unique(split_list_A$Plate)), 
                                    Plate_Barcode=matrix(unique(split_list_A$`Plate-Barcode`)))
    k<-k+1
    new_name_list[[k]]<-paste("Rerun_list",i,"B",sep="_")
    new_work_list[[k]]<-paste("Work_list",i,"B",sep="_")
    new_rwork_list[[k]]<-paste("Reverse_work_list",i,"B",sep="_")
    name_plate_list[[k]]<-paste("Plate_list",i,"B",sep="_")
    new_split_list[[k]]<-split_list_B
    new_plate_list[[k]]<-data.frame(Plate=matrix(unique(split_list_B$Plate)), 
                                    Plate_Barcode=matrix(unique(split_list_B$`Plate-Barcode`)))
    k<-k+1
    new_name_list[[k]]<-paste("Rerun_list",i,"C",sep="_")
    new_work_list[[k]]<-paste("Work_list",i,"C",sep="_")
    new_rwork_list[[k]]<-paste("Reverse_work_list",i,"C",sep="_")
    name_plate_list[[k]]<-paste("Plate_list",i,"C",sep="_")
    new_split_list[[k]]<-split_list_C
    new_plate_list[[k]]<-data.frame(Plate=matrix(unique(split_list_C$Plate)), 
                                    Plate_Barcode=matrix(unique(split_list_C$`Plate-Barcode`)))
  }
  
  if (number_of_plates>57 && number_of_plates<77){
    #If there are above 57 plates in the 95 entries
    plate_twenty<-unique(plates)[20]
    plate_three_nine<-unique(plates)[39]
    plate_five_eight<-unique(plates)[58]
    
    row_split_B<-which(split_list[[i]]$Plate==plate_twenty)[1]
    row_split_C<-which(split_list[[i]]$Plate==plate_three_nine)[1]
    row_split_D<-which(split_list[[i]]$Plate==plate_five_eight)[1]
    
    split_list_A<-split_list[[i]][1:(row_split_B-1),]
    split_list_B<-split_list[[i]][row_split_B:(row_split_C-1),]
    split_list_C<-split_list[[i]][row_split_C:(row_split_D-1),]
    split_list_D<-split_list[[i]][row_split_D:nrow(split_list[[i]]),]
    
    new_name_list[[k]]<-paste("Rerun_list",i,"A",sep="_")
    new_work_list[[k]]<-paste("Work_list",i,"A",sep="_")
    new_rwork_list[[k]]<-paste("Reverse_work_list",i,"A",sep="_")
    name_plate_list[[k]]<-paste("Plate_list",i,"A",sep="_")
    new_split_list[[k]]<-split_list_A
    new_plate_list[[k]]<-data.frame(Plate=matrix(unique(split_list_A$Plate)), 
                                    Plate_Barcode=matrix(unique(split_list_A$`Plate-Barcode`)))
    k<-k+1
    new_name_list[[k]]<-paste("Rerun_list",i,"B",sep="_")
    new_work_list[[k]]<-paste("Work_list",i,"B",sep="_")
    new_rwork_list[[k]]<-paste("Reverse_work_list",i,"B",sep="_")
    name_plate_list[[k]]<-paste("Plate_list",i,"B",sep="_")
    new_split_list[[k]]<-split_list_B
    new_plate_list[[k]]<-data.frame(Plate=matrix(unique(split_list_B$Plate)), 
                                    Plate_Barcode=matrix(unique(split_list_B$`Plate-Barcode`)))
    k<-k+1
    new_name_list[[k]]<-paste("Rerun_list",i,"C",sep="_")
    new_work_list[[k]]<-paste("Work_list",i,"C",sep="_")
    new_rwork_list[[k]]<-paste("Reverse_work_list",i,"C",sep="_")
    name_plate_list[[k]]<-paste("Plate_list",i,"C",sep="_")
    new_split_list[[k]]<-split_list_C
    new_plate_list[[k]]<-data.frame(Plate=matrix(unique(split_list_C$Plate)), 
                                    Plate_Barcode=matrix(unique(split_list_C$`Plate-Barcode`)))
    k<-k+1
    new_name_list[[k]]<-paste("Rerun_list",i,"D",sep="_")
    new_work_list[[k]]<-paste("Work_list",i,"D",sep="_")
    new_rwork_list[[k]]<-paste("Reverse_work_list",i,"D",sep="_")
    name_plate_list[[k]]<-paste("Plate_list",i,"D",sep="_")
    new_split_list[[k]]<-split_list_D
    new_plate_list[[k]]<-data.frame(Plate=matrix(unique(split_list_D$Plate)), 
                                    Plate_Barcode=matrix(unique(split_list_D$`Plate-Barcode`)))
  }
  
  if (number_of_plates>76){
    #If there are above 96 different plates in the 95 entries
    plate_twenty<-unique(plates)[20]
    plate_three_nine<-unique(plates)[39]
    plate_five_eight<-unique(plates)[58]
    plate_seven_seven<-unique(plates)[77]
    
    row_split_B<-which(split_list[[i]]$Plate==plate_twenty)[1]
    row_split_C<-which(split_list[[i]]$Plate==plate__three_nine)[1]
    row_split_D<-which(split_list[[i]]$Plate==plate__five_eight)[1]
    row_split_E<-which(split_list[[i]]$Plate==plate__seven_seven)[1]
    
    split_list_A<-split_list[[i]][1:(row_split_B-1),]
    split_list_B<-split_list[[i]][row_split_B:(row_split_C-1),]
    split_list_C<-split_list[[i]][row_split_C:(row_split_D-1),]
    split_list_D<-split_list[[i]][row_split_D:(row_split_E-1),]
    split_list_E<-split_list[[i]][row_split_E:nrow(split_list[[i]]),]
    
    new_name_list[[k]]<-paste("Rerun_list",i,"A",sep="_")
    new_work_list[[k]]<-paste("Work_list",i,"A",sep="_")
    new_rwork_list[[k]]<-paste("Reverse_work_list",i,"A",sep="_")
    name_plate_list[[k]]<-paste("Plate_list",i,"A",sep="_")
    new_split_list[[k]]<-split_list_A
    new_plate_list[[k]]<-data.frame(Plate=matrix(unique(split_list_A$Plate)), 
                                    Plate_Barcode=matrix(unique(split_list_A$`Plate-Barcode`)))
    k<-k+1
    new_name_list[[k]]<-paste("Rerun_list",i,"B",sep="_")
    new_work_list[[k]]<-paste("Work_list",i,"B",sep="_")
    new_rwork_list[[k]]<-paste("Reverse_work_list",i,"B",sep="_")
    name_plate_list[[k]]<-paste("Plate_list",i,"B",sep="_")
    new_split_list[[k]]<-split_list_B
    new_plate_list[[k]]<-data.frame(Plate=matrix(unique(split_list_B$Plate)), 
                                    Plate_Barcode=matrix(unique(split_list_B$`Plate-Barcode`)))
    k<-k+1
    new_name_list[[k]]<-paste("Rerun_list",i,"C",sep="_")
    new_work_list[[k]]<-paste("Work_list",i,"C",sep="_")
    new_rwork_list[[k]]<-paste("Reverse_work_list",i,"C",sep="_")
    name_plate_list[[k]]<-paste("Plate_list",i,"C",sep="_")
    new_split_list[[k]]<-split_list_C
    new_plate_list[[k]]<-data.frame(Plate=matrix(unique(split_list_C$Plate)), 
                                    Plate_Barcode=matrix(unique(split_list_C$`Plate-Barcode`)))
    k<-k+1
    new_name_list[[k]]<-paste("Rerun_list",i,"D",sep="_")
    new_work_list[[k]]<-paste("Work_list",i,"D",sep="_")
    new_rwork_list[[k]]<-paste("Reverse_work_list",i,"D",sep="_")
    name_plate_list[[k]]<-paste("Plate_list",i,"D",sep="_")
    new_split_list[[k]]<-split_list_D
    new_plate_list[[k]]<-data.frame(Plate=matrix(unique(split_list_D$Plate)), 
                                    Plate_Barcode=matrix(unique(split_list_D$`Plate-Barcode`)))
    k<-k+1
    new_name_list[[k]]<-paste("Rerun_list",i,"E",sep="_")
    new_work_list[[k]]<-paste("Work_list",i,"E",sep="_")
    new_rwork_list[[k]]<-paste("Reverse_work_list",i,"E",sep="_")
    name_plate_list[[k]]<-paste("Plate_list",i,"E",sep="_")
    new_split_list[[k]]<-split_list_E
    new_plate_list[[k]]<-data.frame(Plate=matrix(unique(split_list_E$Plate)), 
                                    Plate_Barcode=matrix(unique(split_list_E$`Plate-Barcode`)))
  }
}

#Number of sheets needed for rerun, work and reverse work lists each.
new_sheet_number<-length(new_split_list)

#Initiating excel file.
excel_file<-createWorkbook()

#Manually creating "Position at Destination" vector.
pos_at_dest<-c("A01","A02","A03","A04","A05","A06","A07","A08","A09","A10","A11",
               "A12","B01","B02","B03","B04","B05","B06","B07","B08","B09","B10",
               "B11","B12","C01","C02","C03","C04","C05","C06","C07","C08","C09",
               "C10","C11","C12","D01","D02","D03","D04","D05","D06","D07","D08",
               "D09","D10","D11","D12","E01","E02","E03","E04","E05","E06","E07",
               "E08","E09","E10","E11","E12","F01","F02","F03","F04","F05","F06",
               "F07","F08","F09","F10","F11","F12","G01","G02","G03","G04","G05",
               "G06","G07","G08","G09","G10","G11","G12","H01","H02","H03","H04",
               "H05","H06","H07","H08","H09","H10","H11")


#Creating cell color styles for red, yellow and green.
red_c<-createStyle(fontColour="#000000",bgFill="#FF0000")
yellow_c<-createStyle(fontColour="#000000",bgFill="#FFFF00")
lgreen_c<-createStyle(fontColour="#000000",bgFill="#92D050")

centering <- createStyle(halign = "center")

#Writing rerun list to excel sheets.
#Dest_n is the index in the current post_at_dest vector.
#Dest_p is the index in the previous post_at_dest_vector.

dest_n<-0
for (i in 1:new_sheet_number){
  #Resetting index in post_at_dest if 95 entries have been fulfilled.
  if (dest_n>94){
    dest_n<-0
  }
  dest_p<-dest_n
  dest_n<-nrow(new_split_list[[i]])+dest_n
  #Creating temporary variable of names and assigning sheets with the names
  name <- new_name_list[[i]]
  plate_name<- name_plate_list[[i]]
  assign(name,addWorksheet(wb=excel_file,sheetName=new_name_list[[i]]))
  assign(plate_name,addWorksheet(wb=excel_file,sheetName=name_plate_list[[i]]))
  
#CHANGE PLATE BARCODE BELOW - CHANGE PLATE BARCODE BELOW - CHANGE PLATE BARCODE BELOW - CHANGE PLATE BARCODE BELOW (the string)
  new_split_list[[i]]$`Destination-Plate`<-rep("ABXXXXXXXX", nrow(new_split_list[[i]]))
  
  new_split_list[[i]]$`Position-at-Destination`<-pos_at_dest[(dest_p+1):dest_n]
  #Writing rerun and plate sheets to excel file (with color formatting)
  writeData(wb=excel_file,sheet=eval(parse(text = new_name_list[[i]])),x=new_split_list[[i]],
            rowNames = FALSE)
  writeData(wb=excel_file,sheet=eval(parse(text = name_plate_list[[i]])),x=new_plate_list[[i]],
            rowNames = FALSE,colNames = FALSE)
  conditionalFormatting(excel_file,name,cols=1:5,rows=1,rule ="!=0", style=red_c,
                        type="expression")
  conditionalFormatting(excel_file,name,cols=6,rows=1,rule ="!=0", style=yellow_c,
                        type="expression")
  conditionalFormatting(excel_file,name,cols=7,rows=1,rule ="!=0", style=lgreen_c,
                        type="expression")
  setColWidths(excel_file,name,cols=c(2,4,5,6,7),widths=20)
  setColWidths(excel_file,plate_name,cols=1:2,widths=20)
  addStyle(excel_file,name,centering,rows=2:(nrow(new_split_list[[i]])+1),cols=2:7
           ,gridExpand = T)
  addStyle(excel_file,plate_name,centering,rows=1:length(unique(new_split_list[[i]]$Plate))
           ,cols=1:2, gridExpand = T)
}

#Making work-list sheet data.
work_list<-new_split_list
rwork_list<-new_split_list

#Creating work sheets for each rerun list.
for (i in 1:new_sheet_number){
  work_list[[i]]$combined<-paste(work_list[[i]]$`Plate-Barcode`,work_list[[i]]$Position
                                 ,work_list[[i]]$`Tube-Barcode`,work_list[[i]]$`Destination-Plate`
                                 ,work_list[[i]]$`Position-at-Destination`,sep=",")
  rwork_list[[i]]$combined<-paste(rwork_list[[i]]$`Destination-Plate`,rwork_list[[i]]$`Position-at-Destination`
                                 ,rwork_list[[i]]$`Tube-Barcode`,rwork_list[[i]]$`Plate-Barcode`
                                 ,rwork_list[[i]]$Position,sep=",")
  
  work_list[[i]]<-select(work_list[[i]],combined)
  work_list[[i]]$combined<-str_replace(work_list[[i]]$combined,"NA","")
  rwork_list[[i]]<-select(rwork_list[[i]],combined)
  rwork_list[[i]]$combined<-str_replace(rwork_list[[i]]$combined,"NA","")
}

#Creating variable holding the current date.
date<-format(Sys.Date(),"%d-%m-%Y")

#Creating directory/folder to write worklists to.
dir.create(paste(getwd(),paste("Worklists",date,sep="-"),sep="/"))
dir.create(paste(getwd(),paste("Reverse-Worklists",date,sep="-"),sep="/"))

#Creating variable holding the string of the file path.
file_path_work<-paste(getwd(),paste("Worklists",date,sep="-"),sep="/")
file_path_rwork<-paste(getwd(),paste("Reverse-Worklists",date,sep="-"),sep="/")

#Writing work list to excel sheets.
for (i in 1:new_sheet_number){
  work <- new_work_list[[i]]
  assign(work,addWorksheet(wb=excel_file,sheetName = new_work_list[i]))
  writeData(wb=excel_file,sheet=eval(parse(text = new_work_list[[i]])),x=work_list[[i]],
            rowNames = FALSE,colNames = FALSE)
  write.table(work_list[[i]],file=paste(file_path_work,paste(new_work_list[[i]],"csv",sep="."),sep="/"),
              col.names=FALSE,sep=",",quote=FALSE,row.names = FALSE)
}

#Making reverse work list to excel sheets.
for (i in 1:new_sheet_number){
  rwork <-  new_rwork_list[[i]]
  assign(rwork,addWorksheet(wb=excel_file,sheetName =  new_rwork_list[[i]]))
  writeData(wb=excel_file,sheet=eval(parse(text =  new_rwork_list[[i]])),x=rwork_list[[i]],
            rowNames = FALSE,colNames = FALSE)
  write.table(rwork_list[[i]],file=paste(file_path_rwork,paste(new_rwork_list[[i]],"csv",sep="."),sep="/"),
              col.names=FALSE,sep=",",quote=FALSE, row.names = FALSE)
}

if(interactive()){
  winDialog(type = c("ok"),
            "Save the excel file with .xlsx extension (Eksempel: rerun.xlsx)")
}

print("Save file in pop-up window (remember .xlsx extension)")

#Saving and writing final excel file with interactive input dialog.
saveWorkbook(excel_file, file=choose.files(caption="Save As...",filters = c("*.xlsx"))
             ,overwrite=TRUE)

