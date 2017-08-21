install.packages("dplyr")
install.packages("openxlsx")
install.packages("xlsx")
install.packages("rJava")
library("openxlsx")
library("dplyr")
library("rJava")
library("xlsx")
#get Sheetnames of a excelworkbook
Sheetnames<- getSheetNames("F:\\Demand forecasting\\Original data\\DMBI.xlsx")
#get no of sheets in a excelworkbook
count<-length(Sheetnames)
#ITERATE THROUGHT THE SHEET TO READ ONLY SELCETED FEW COLUMNS IN A EXCEL WORK BOOK
wb<-createWorkbook()
#get unique Materials into a vector
#dataframe1<-read.xlsx("F:\\Demand forecasting\\Original data\\DMBI.xlsx",sheet=i+1,cols=c(1,2,3,10,11))
#uniqueMaterials<-unique(dataframe1$Material)
Material<- c("8170ZN5","8444MS7-1000","8180ZN5","8122ZN5","8180ZN5V","8406MS7","8182ZN5","8106ZN5")
finaldata<-setNames(data.frame(matrix(ncol=length(colname),nrow=0)),colname)
Sheetloop<-length(Sheetnames)-2
Materialloop<-length(Material)
for(i in 1:Materialloop){
  tempMaterial <- Material[i]
  tempdata<-data.frame()
  print(paste("current Material in loop ",tempMaterial))
  print(paste("Material number is",i))
  for(j in 1:Sheetloop){
    dataframe1<-read.xlsx("F:\\Demand forecasting\\Original data\\DMBI.xlsx",sheet=j,cols=c(1,2,3,10,11))
    print(paste("Sheet number is value is",j))
    #apply filter to a dataframe
    Filtereddata<-filter(dataframe1,Material==tempMaterial)
    Salesvalue <- sum(Filtereddata$Umsatz)
    colname<-colnames(Filtereddata)
    tempdata<-rbind(tempdata,Filtereddata[1,])
    tempdata[j,5]<-Salesvalue
  }
  finaldata<-rbind(finaldata,tempdata)
  print(finaldata)
}
write.xlsx(x = finaldata, file = "F:\\Demand forecasting\\Original data\\MOdelData.xlsx",col.names = TRUE,
           sheetName = "TestSheet", row.names = FALSE)
