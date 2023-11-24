#this code wouldn't exist without the work of two Lausanne university students who built a first draw: Guillaume Péclard and Andrea Ferrazzo. Thanks to them.

library(readxl)
library(tibble)
library(dplyr)
library(tidyr)
library(stringr) 

#specifies where our raw data is filed
setwd("C:/Users/Chapatte/Documents/R code and data/")
folder <- "./Weekly Pacing -20230324T125929Z-001/Weekly Pacing/"

#gets a list of all the Excel files in the folder
excel_files <- list.files(folder, pattern = ".xlsx$", full.names = TRUE)

url<-excel_files[0] #first tab

transform_sheets<-function(url,sheet){

  modify_names <- function(names) {
  
  droite = function(x,n){
  substring(x,nchar(x)-n+1)}
  
    for (i in 1:length(names)) {
      if (i <= 2) {
        #if the column is one of the first 2, keep the original name
        new_name <- names[i]
      } else {
        #if the column is not one of the first 2, append ".i-3" to the original name
        new_name <- paste0(gsub(" ","",droite(names[i],2)), ".", i - 3)
      }
      names[i] <- new_name
    }
    return(names)
  }

x <- read_excel(url, sheet) %>% as_tibble()
x <- fill(x,...1)


colnames(x)[3:28] <- x[6,3:28] #the columns of the tibble take the name of the weeks...
colnames(x) <- modify_names(colnames(x))
colnames(x)[1:2] <- c("Category", "Year") #... except for the two first columns
x <- x %>% mutate("Year_str"=x$Year,.after="Year")
x <- subset(x, x$Year !="Previous Year" & x$Year !="2 years ago" & x$Year !="2 years Ago" & x$Year !="2019" & x$Year != "Current vs Previous %" & x$Year != "Current vs 2 years ago %" & x$Year != "Current vs 2019 %" & x$Year != "% Chg" & x$Year != "Supply (Nights)" & x$Year != "Demand (Nights)" & x$Year != "Occupancy Rate" & x$Year != "Average Daily Rate" & x$Year != "RevPAR" & x$Year != "Revenue (USD)")
x <- x %>% 
    mutate("Year" = case_when(Year == "Current Year" ~ str_sub(x[1,4],start = -4)))
	#replaces e. g. "Current year" by the year: "2022"
x <- subset(x,!is.na(Category))

x <- pivot_longer(x, cols = -c(Category, Year,Year_str), names_to = "week", values_to = "values_non_num") #pivots all the columns, except Category and Year
x <- separate(x, week, into = c("Week", "in_x_weeks"), sep = "\\.", fill = "right") #Separates the forecasted column and the number of weeks left before it

x <- x %>% 
    mutate(Values = as.numeric(values_non_num),in_x_weeks=as.numeric(in_x_weeks),Week=as.numeric(Week),Year=as.numeric(Year)) %>% 
    select(-values_non_num)
	#converts the non-numeric values into numeric values and renames values_non_num into Values

x<-add_column(x,Week_absolu=0,.after="Category") # Adds the column Week_absolu

for (i in 2:nrow(x)){
	if(x$Week[i]<x$Week[1]){
		x$Year[i]<-x$Year[i]+1}}


#gives values to the column Week_absolu
for (i in 1:nrow(x)){
if (x$Year[i]==2023){
x$Week_absolu[i]<-x$Week[i]+208}
else if (x$Year[i]==2022){
x$Week_absolu[i]<-x$Week[i]+156}
else if (x$Year[i]==2021){
x$Week_absolu[i]<-x$Week[i]+104}
else if (x$Year[i]==2020){
x$Week_absolu[i]<-x$Week[i]+52}
else if (x$Year[i]==2020){
x$Week_absolu[i]<-x$Week[i]}
}

return(x)
}

#declaration of the variable combined_data and adding the column "Region"
url <- excel_files[1]
combined_data <- transform_sheets(url, excel_sheets(url)[1])
combined_data$Region <- "Canton de VAUD"

#loops through each Excel file and apply the transform_sheet() function to each sheet
for (i in 1:length(excel_files)) {
  # specify the URL of the Excel file
  url <- excel_files[i]

  #reads all sheets of the Excel file except the last one into a list of tibbles
  sheets <- lapply(excel_sheets(url), function(sheet) {
    if (sheet != excel_sheets(url)[length(excel_sheets(url))]) {
      # transform the sheet using the `transform_sheet` function
      sheet_data <- transform_sheets(url, sheet)
      # add the sheet Region as a new column
      sheet_data$Region <- sheet
      sheet_data #returns the value sheet_data
	  
    }
  })
  
  # Removes any duplicates rows from the combined_data tibble
  combined_data <- unique(combined_data)

  # loop through each tibble in the list and copy its data to the combined tibble
  for (j in 1:length(sheets)) {
    new_data <- sheets[[j]]
    
    # check if new_data is empty before attempting to join it with combined_data
    if (!is.null(new_data)) {
      # anti_join puts aside the values from new_data already present in combined data
      new_data <- anti_join(new_data, combined_data, by = colnames(new_data))
      
      # joins the values from combined_data with those from new_data (=those from the currently processed Excel sheet)
      combined_data <- bind_rows(combined_data, new_data)
    }
  }
}

subset_data<-combined_data %>% filter (Year_str=="Current Year" & Category=="Demand (Nights)")
subset_data<-subset_data %>% select (Region,Year,Week_absolu,in_x_weeks,Week,Values)
subset_data<-subset_data %>% arrange (Region,desc(Week_absolu),Week,in_x_weeks)

calcul_des_differences<-function (tibble){
evolution<-tibble(difference = numeric())
for (i in 2:nrow(tibble))
	{
	if(tibble$Week[i]==tibble$Week[i-1])
		{
		difference<-tibble$Values[i-1]-tibble$Values[i]
		evolution<-add_row(evolution,difference)
		}
	else {
	difference<-tibble$Values[i-1]
	evolution<-add_row(evolution,difference)
	}}
difference<-0
evolution<-add_row(evolution,difference)
return (evolution)}
#calculates the difference between two values from one week to the next one. evolution is the tibble that stocks those differences

lissage <-function (tibble) {
for (i in 2:(nrow(tibble)-1)){
	if (tibble[i,7]<=0 & tibble$Week[i]==tibble$Week[i+1]){
	tibble[i,6]<-tibble[i+1,6]
	tibble[i,7]<-0
	}}
	tibble[(nrow(tibble)),6]<-0
	return (tibble)}
#this function "flattens" the results so that the reservations don't go down from one week to the next one

for(i in 1:10){
col_evolution<-calcul_des_differences(subset_data) # 1) calculates the evolution of the number of reservations in a tibble
subset_data<- subset_data %>% mutate(col_evolution) %>% lissage() %>% select(-difference)
# 2) joins again the columns from echantillon and the evolution of the reservations 3)flattens the evolution of the number of reservations. 4) Removes the differences calculation before flattening
i<-i+1
#N.B.: To flatten also for the small destinations, you have to define i in 1:10
}

subset_data <- subset_data %>% filter (Week_absolu <= 221 & Week_absolu >= 148)
write.csv(subset_data, "subset_data.csv", row.names = FALSE)


lignes_pour_completer=tibble(Region="Test",Year=2000,Week_absolu=100,in_x_weeks=26,Week=53,Values=1000)
for (i in 2:nrow(subset_data))
	{if(subset_data$Week[i]==subset_data$Week[i-1]){
		ecart<-subset_data$in_x_weeks[i]-subset_data$in_x_weeks[i-1]
		if(ecart>1)
			{for (j in 1:(ecart-1))
				{nouvelle_ligne<-tibble (
				Region=subset_data$Region[i],
				Year=subset_data$Year[i],
				Week_absolu=subset_data$Week_absolu[i],
				in_x_weeks=subset_data$in_x_weeks[i]-j,
				Week=subset_data$Week[i],
				Values=((j/(ecart))*subset_data$Values[i-1]+(1-(j/(ecart)))*subset_data$Values[i]))
				lignes_pour_completer<-add_row(lignes_pour_completer,nouvelle_ligne)
				}
				}
				}}
#Completes the missing values when in_x_weeks>=1

lignes_pour_completer <- lignes_pour_completer %>%  slice(-1) #deletes the test line to define the format and the name of the variables of the tibble
subset_data<-bind_rows(subset_data,lignes_pour_completer)
subset_data<-subset_data %>% arrange (Region,desc(Week_absolu),Week,in_x_weeks)


lignes_pour_completer=tibble(Region="Test",Year=2000,Week_absolu=100,in_x_weeks=26,Week=53,Values=1000)				
for (i in 2:nrow(subset_data)){
	
	if(subset_data$Week_absolu[i]==subset_data$Week_absolu[i-1]-1) {
		if (subset_data$in_x_weeks[i]>0){
			for (j in 1:subset_data$in_x_weeks[i])
				{nouvelle_ligne<-tibble (
				Region=subset_data$Region[i],
				Year=subset_data$Year[i],
				Week_absolu=subset_data$Week_absolu[i],
				in_x_weeks=subset_data$in_x_weeks[i]-j,
				Week=subset_data$Week[i],
				Values=subset_data$Values[i]+(subset_data$Values[i]-subset_data$Values[i+1])*j)
				lignes_pour_completer<-add_row(lignes_pour_completer,nouvelle_ligne)}
				}}}

#completes the missing values when in_x_weeks=0 and closeby are missing

lignes_pour_completer <- lignes_pour_completer %>%  slice(-1) #deletes the test line to define the format and the name of the variables of the tibble
subset_data<-bind_rows(subset_data,lignes_pour_completer)
subset_data<-subset_data %>% arrange (Region,desc(Week_absolu),Week,in_x_weeks)



lignes_pour_completer=tibble(Region="Test",Year=2000,Week_absolu=100,in_x_weeks=26,Week=53,Values=1000)				
for (i in 2:nrow(subset_data)){
	
	if(subset_data$Week_absolu[i]==subset_data$Week_absolu[i-1]-1) {
		if (subset_data$in_x_weeks[i-1]==24)
				{nouvelle_ligne<-tibble (
				Region=subset_data$Region[i-1],
				Year=subset_data$Year[i-1],
				Week_absolu=subset_data$Week_absolu[i-1],
				in_x_weeks=subset_data$in_x_weeks[i-1],
				Week=subset_data$Week[i-1],
				Values=max(subset_data$Values[i-1]-(subset_data$Values[i-2]-subset_data$Values[i-1]),0))
				lignes_pour_completer<-add_row(lignes_pour_completer,nouvelle_ligne)}
				}}

#Completes the values when in_x_weeks=24 is missing

lignes_pour_completer <- lignes_pour_completer %>%  slice(-1) #deletes the test line to define the format and the name of the variables of the tibble
subset_data<-bind_rows(subset_data,lignes_pour_completer)
subset_data<-subset_data %>% arrange (Region,desc(Week_absolu),Week,in_x_weeks)

write.csv(subset_data, "subset_data.csv", row.names = FALSE)