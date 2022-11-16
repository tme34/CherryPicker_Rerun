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

#Destination Scan interactive message
if(interactive()){
  winDialog(type = c("ok"),
            "Select one worklist file from A-D")
}
#Selecting number of Worklists/How many runs (A,B,C,D,E)/(1,2,3,4,5)
#Load first as A
work_select_A<-choose.files(caption="Select Worklist", multi=FALSE, filters=c("*.csv"))
#Determining what subworklist is selected
work_extract<-str_extract(work_select_A,"list_[0-9]{1,3}_[A-E]{1}")
#Editing subworklist to be A, since the script starts from A-E.
work_select_A<-sub("list_[0-9]{1,3}_[A-E]{1}",
                   paste(substring(work_extract,first=1,last=(nchar(work_extract)-2)),"A",sep="_"),
                   work_select_A)

Work_A<-read.table(file=work_select_A,sep=",",header=FALSE)

#Determining number of rows. When full = 95.
row_count<-nrow(Work_A)

#Determining number of plates in each work-list
plates_A<-length(unique(Work_A$V1))
plates_B<-0
plates_C<-0
plates_D<-0

#B
if(row_count<95 && plates_A==19){
  work_select_B<-sub("_A","_B",work_select_A)
  Work_B<-read.table(file=work_select_B,sep=",",header=FALSE)
  row_count<-row_count+nrow(Work_B)
  Work_A<-rbind(Work_A,Work_B)
  plates_B<-length(unique(Work_B$V1))
}

#C
if(row_count<95 && plates_B==19){
  work_select_C<-sub("_A","_C",work_select_A)
  Work_C<-read.table(file=work_select_C,sep=",",header=FALSE)
  row_count<-row_count+nrow(Work_C)
  Work_A<-rbind(Work_A,Work_C)
  plates_C<-length(unique(Work_C$V1))
}

#D
if(row_count<95 && plates_C==19){
  work_select_D<-sub("_A","_D",work_select_A)
  Work_D<-read.table(file=work_select_D,sep=",",header=FALSE)
  row_count<-row_count+nrow(Work_D)
  Work_A<-rbind(Work_A,Work_D)
  plates_D<-length(unique(Work_D$V1))
}

#E
if(row_count<95 && plates_D==19){
  work_select_E<-sub("_A","_E",work_select_A)
  Work_E<-read.table(file=work_select_E,sep=",",header=FALSE)
  row_count<-row_count+nrow(Work_E)
  Work_A<-rbind(Work_A,Work_E)
}

#Renaming the fully appended list to full
Work_full<-Work_A

#Creating accessible column names
colnames(Work_full)<-c("Plate_Barcode","Position","Tube_Barcode","Destination_Plate","Position_At_Destination")

#Select message for destination scan
if(interactive()){
  winDialog(type = c("ok"),
            "Select destination scan file matching worklists")
}

#Reading Destination scan of 95/96 full plate
destination_scan<-read.table(file=choose.files(caption = "Select Destination Scan File"
                                               , multi = FALSE)
                             , sep=",", header=FALSE) %>%
  arrange(.,V2)

#Removing H12 row, if present in CSV file
destination_scan<-destination_scan %>%
  filter(V2!="H12")
destination_scan<-destination_scan[1:nrow(Work_full),]

#Creating the final data frame of Tjek (Placeholder for Destination scan)
Tjek<-data.frame(Tube_Barcode_Rerun_liste = matrix(Work_full$Tube_Barcode),
                 Tube_Barcode_Scan_Destination=matrix(destination_scan$V3),
                 Position=matrix(destination_scan$V2)) %>%
  mutate(.,Tjek=Tube_Barcode_Rerun_liste-Tube_Barcode_Scan_Destination) %>%
  relocate(.,Tjek,.before=Position)
colnames(Tjek)<-c("Tube-Barcode - Rerun-liste","Tube-Barcode - Scan Destination","Tjek","Position")
#Initiating excel workbook
tjek_file<-createWorkbook()

#Creating excel styles for cell coloring and font
red_c<-createStyle(fontColour="#000000",bgFill="#FF0000")
lgreen_c<-createStyle(fontColour="#000000",bgFill="#92D050")
bold_style<-createStyle(textDecoration = "Bold")
centering <- createStyle(halign = "center")

#Initiating excel-sheet and writing the sheet data to the excel workbook
tjek_sheet<-addWorksheet(wb=tjek_file,sheetName="Tjek")
writeData(wb=tjek_file, sheet=tjek_sheet, x=Tjek, rowNames = FALSE)
conditionalFormatting(tjek_file,tjek_sheet,cols=3,rows=2:(nrow(Tjek)+1),rule ="!=0", style=red_c,
                      type="expression")
conditionalFormatting(tjek_file,tjek_sheet,cols=3,rows=2:(nrow(Tjek)+1),rule ="==0", style=lgreen_c,
                      type="expression")
conditionalFormatting(tjek_file,tjek_sheet,cols=1:4,rows=1,rule ="!=0", style=bold_style,
                      type="expression")
setColWidths(tjek_file,tjek_sheet,cols=1:3,widths=25)
addStyle(tjek_file,tjek_sheet,centering,rows=1:(nrow(Tjek)+1),cols=1:4, gridExpand = T)

#Extracting worklist number from filename
run_number<-str_extract(work_select_A,"[0-9]{1,3}_.{1,9}$")
dest_scan_tjek_name<-paste("Tjek_list",sub("_A","",run_number),sep="_")

#Saving workbook with worklist number in the filename
saveWorkbook(tjek_file, file=paste(dest_scan_tjek_name,"xlsx",sep="."), overwrite=TRUE)

print("Tjek Complete")