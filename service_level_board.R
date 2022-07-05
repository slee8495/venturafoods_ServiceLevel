library(tidyverse)
library(readxl)
library(writexl)
library(reshape2)
library(openxlsx)
library(readxlsb)
library(data.table)


# Read raw data

service_level_data <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Report/Root Cause Report/2022/Feb/2.10.2022/service_level_data.xlsx", 
                                 col_types = c("text", "text", "text", 
                                               "text", "text", "text", "text", "text", 
                                               "text", "text", "text", "numeric", "text", 
                                               "text", "text", "text", "text", "text", 
                                               "numeric", "numeric", "numeric", 
                                               "numeric", "numeric", "text", "text", 
                                               "text", "text", "text", "text", "text", 
                                               "text", "text", "text", "text", "text", 
                                               "text", "text", "text", "text", "text", 
                                               "text", "text", "text", "text", "text", 
                                               "text", "text", "numeric", "numeric"))

service_level_data[, -48:-49] -> service_level_data

colnames(service_level_data)[1] <- "Super_Customer_Number"
colnames(service_level_data)[2] <- "Super_Customer_Name"
colnames(service_level_data)[3] <- "Mfg_Location"
colnames(service_level_data)[4] <- "Mfg_Location_Name"
colnames(service_level_data)[5] <- "Ship_Location"
colnames(service_level_data)[6] <- "Ship_Location_Name"
colnames(service_level_data)[7] <- "Product_Label"
colnames(service_level_data)[8] <- "Product_Label_Name"
colnames(service_level_data)[9] <- "Product_Platform_Desc"
colnames(service_level_data)[10] <- "Line_Number"
colnames(service_level_data)[11] <- "Line_Description"
colnames(service_level_data)[12] <- "AS_Variable_Margine_Per_Pound"
colnames(service_level_data)[13] <- "Pack_Size_Name"
colnames(service_level_data)[14] <- "Product_Category_Name_105"
colnames(service_level_data)[15] <- "Base_Product"
colnames(service_level_data)[16] <- "Base_Product_Name_10206"
colnames(service_level_data)[17] <- "Product_SKU"
colnames(service_level_data)[18] <- "Product_SKU_Name_10206_GNS"
colnames(service_level_data)[19] <- "AS_Cases"
colnames(service_level_data)[20] <- "AS_Order_Quantity"
colnames(service_level_data)[21] <- "AS_Original_Order_Quantity"
colnames(service_level_data)[22] <- "AS_Net_Pounds"
colnames(service_level_data)[23] <- "AS_Variable_Margin"
colnames(service_level_data)[24] <- "Pack_Size"
colnames(service_level_data)[25] <- "Shipped_Date"
colnames(service_level_data)[26] <- "Shipped_Year"
colnames(service_level_data)[27] <- "Shipped_Month"
colnames(service_level_data)[28] <- "Short_Ship_Reason"
colnames(service_level_data)[29] <- "Geo_Location"
colnames(service_level_data)[30] <- "Channel"
colnames(service_level_data)[31] <- "Customer"
colnames(service_level_data)[32] <- "MFG_Site"
colnames(service_level_data)[33] <- "Shipping_Main"
colnames(service_level_data)[34] <- "Shipping_Location"
colnames(service_level_data)[35] <- "Site_Line"
colnames(service_level_data)[36] <- "Item"
colnames(service_level_data)[37] <- "Ship_Month"
colnames(service_level_data)[38] <- "Ship_Year_Month"
colnames(service_level_data)[39] <- "Sysco_BAPA_SKU"
colnames(service_level_data)[40] <- "Label_w_Desc"
colnames(service_level_data)[41] <- "Product_SKU_w_Description"
colnames(service_level_data)[42] <- "COVID_PRIORITY_SKU"
colnames(service_level_data)[43] <- "PRIORITY_SKU"
colnames(service_level_data)[44] <- "MTO_SOP"
colnames(service_level_data)[45] <- "Macro_Platform"
colnames(service_level_data)[46] <- "Starch_Supply_Impacted"
colnames(service_level_data)[47] <- "Starch_Impacted"



service_level_data %>% 
  dplyr::mutate(Case_Cuts         = AS_Cases - AS_Order_Quantity) %>% 
  dplyr::mutate(Percent_of_total  = AS_Net_Pounds) %>%
  dplyr::mutate(Case_Fill_percent = AS_Cases / AS_Order_Quantity) %>%  
  dplyr::mutate(Case_Cut_Original_Order_Qty = AS_Original_Order_Quantity - AS_Cases) -> service_level_data

service_level_data %>% 
  dplyr::mutate(LBS_Short = (AS_Net_Pounds / AS_Cases) * Case_Cuts) %>% 
  dplyr::mutate(Total_Removed = AS_Original_Order_Quantity - AS_Order_Quantity) %>% 
  dplyr::mutate(Lost_Sales_due_to_Allocation = AS_Original_Order_Quantity - AS_Order_Quantity) %>% 
  dplyr::mutate(Original_Case_Cut = AS_Original_Order_Quantity - AS_Cases) %>% 
  dplyr::mutate(Original_Case_Fill = AS_Cases / AS_Original_Order_Quantity) -> service_level_data


replace(service_level_data$Case_Fill_percent, is.nan(service_level_data$Case_Fill_percent), 0) -> Case_Fill_Percent
format(Case_Fill_Percent, digits = 3)

cbind(service_level_data, Case_Fill_Percent) -> a

save(service_level_data, file = "service_level_data")




