#this code wouldn't exist without the work of two Lausanne university students who built a first draw: Guillaume Péclard and Andrea Ferrazzo. Thanks to them.

library(readxl)
library(tibble)
library(dplyr)
library(tidyr)
library(stringr)
library(forecast)
library(padr)
library(imputeTS)
library(readr)
library(lubridate)

#loads the data set
setwd("C:/Users/Chapatte/Documents/R code and data")
combined_data <- read_csv("combined_data.csv")

#filters the data
filtered_data <- combined_data %>%
  filter(forecasted_week == 0) %>% 
  select(-c(Week,Year,forecasted_week))

#keeps only the maximum value for each combination of Category, title, and Date
filtered_data <- filtered_data %>%
  group_by(Category, title, Date) %>%
  slice_max(Values) %>%
  ungroup()

#fills in missing weeks with NA values
long_data_pad <- filtered_data %>%
  group_by(Category, title) %>%
  complete(fill = list(Values = NA)) %>%
  #pad fills temporal series with missing values, finding out which are the missing values. Here the interval is the week
  pad(interval = "week", start_val = min(filtered_data$Date), end_val = max(filtered_data$Date)) %>%
  ungroup()

grouped_data <- long_data_pad %>%
  group_by(Category, title)

#imputes missing values for each group using Kalman filter
imputed_data <- grouped_data %>%
  mutate(Values = na_kalman(Values, model = "auto.arima")) %>%
  ungroup()

#replaces original NA values with imputed values
long_data_pad$Values[is.na(long_data_pad$Values)] <- imputed_data$Values[is.na(long_data_pad$Values)]
#long_data_pad$Values[is.na(long_data_pad$Values)] is a list of aroung 1300 valeurs, all NAs. These are the values which we replace with the calculated values

write.csv(long_data_pad, "finaldata.csv", row.names = FALSE)

#add 17 weeks to the df as it the the number of week we want to predict and we want to have a df the same length as the forecasted one for merging
long_data_pad2 <- long_data_pad %>%
  group_by(Category, title) %>%
  complete(fill = list(Values = NA)) %>%
  pad(interval = "week", start_val = min(long_data$Date), end_val = max(long_data$Date)+weeks(17)) %>%
  ungroup()

#creates a function to iterate over all combination of Category and title and perform auto.arima at each iteration
create_forecast <- function(data) {
  categories <- unique(data$Category)
  titles <- unique(data$title)
  forecasts <- list()
  
  for (cat in categories) {
    for (title in titles) {
      subset_data <- data[data$Category == cat & data$title == title, ]
      ts_data <- ts(subset_data$Values, frequency = 52)  
      
      #generates forecast using auto.arima
      forecast_model <- forecast_model <- auto.arima(ts_data)
      forecast_values <- forecast(forecast_model, h = 17)
      
      #stores forecasted values, upper bounds, lower bounds, and associated dates
      forecast_data <- data.frame(
        Date = seq(max(subset_data$Date)+ 7, length.out = 17, by = "week"),
        Forecast = forecast_values$mean,
        Upper = forecast_values$upper,
        Lower = forecast_values$lower
      )
      
      forecasts[[paste(cat, title, sep = " - ")]] <- forecast_data
    }
  }
  
  return(forecasts)
}

forecasts <- create_forecast(long_data_pad)

combined_forecast <- list()

for (key in names(forecasts)) {
  forecast_data <- forecasts[[key]]
  category_title <- strsplit(key, " - ")
  category <- category_title[[1]][1]
  title <- category_title[[1]][2]
  
  forecast_data$Category <- category
  forecast_data$Title <- title
  
  combined_forecast[[key]] <- forecast_data
}

combined_forecast_df <- combined_forecast %>%
  bind_rows() %>%
  select(Category, Title, Date, Forecast, Upper.80., Upper.95., Lower.80., Lower.95.)

long_data_pad4 <- full_join(combined_forecast_df, long_data_pad2, by = c("Category" = "Category",
                                                                         "Date" = "Date", 
                                                                         "Title" = "title"))

#adds actual values to forecasted data for better visualisation

long_data_pad4$Forecast <- ifelse(is.na(long_data_pad4$Forecast), long_data_pad4$Values, long_data_pad4$Forecast)
long_data_pad4$Upper.80. <- ifelse(is.na(long_data_pad4$Upper.80.), long_data_pad4$Values, long_data_pad4$Upper.80.)
long_data_pad4$Upper.95. <- ifelse(is.na(long_data_pad4$Upper.95.), long_data_pad4$Values, long_data_pad4$Upper.95.)
long_data_pad4$Lower.80. <- ifelse(is.na(long_data_pad4$Lower.80.), long_data_pad4$Values, long_data_pad4$Lower.80.)
long_data_pad4$Lower.95. <- ifelse(is.na(long_data_pad4$Lower.95.), long_data_pad4$Values, long_data_pad4$Lower.95.)


forecasted_df <- long_data_pad4


#downloads the data


write.csv(forecasted_df, "forecasted_df.csv", row.names = FALSE)


