library(tidyverse)
library(readxl)
library(writexl)
library(reshape2)
library(openxlsx)
library(readxlsb)
library(data.table)

# Read data



thru_2_20 <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Report/Root Cause Report/2022/Feb/2.20.2022/thru 2.20.xlsx", 
                        col_names = FALSE, col_types = c("text", "text", "text", 
                                                         "text", "text", "text", "text", "text", 
                                                         "text", "text", "text", "numeric", "text", 
                                                         "text", "text", "text", "text", "text", 
                                                         "numeric", "numeric", "numeric", 
                                                         "numeric", "numeric", "text", "text", 
                                                         "text", "text", "text", "text", "text", 
                                                         "text", "text", "text", "text", "text", 
                                                         "text", "text", "text", "text", "text", 
                                                         "text", "text", "text", "text", "text", 
                                                         "text", "text"))


colnames(thru_2_20)[1] <- "Super_Customer_Number"
colnames(thru_2_20)[2] <- "Super_Customer_Name"
colnames(thru_2_20)[3] <- "Mfg_Location"
colnames(thru_2_20)[4] <- "Mfg_Location_Name"
colnames(thru_2_20)[5] <- "Ship_Location"
colnames(thru_2_20)[6] <- "Ship_Location_Name"
colnames(thru_2_20)[7] <- "Product_Label"
colnames(thru_2_20)[8] <- "Product_Label_Name"
colnames(thru_2_20)[9] <- "Product_Platform_Desc"
colnames(thru_2_20)[10] <- "Line_Number"
colnames(thru_2_20)[11] <- "Line_Description"
colnames(thru_2_20)[12] <- "AS_Variable_Margine_Per_Pound"
colnames(thru_2_20)[13] <- "Pack_Size_Name"
colnames(thru_2_20)[14] <- "Product_Category_Name_105"
colnames(thru_2_20)[15] <- "Base_Product"
colnames(thru_2_20)[16] <- "Base_Product_Name_10206"
colnames(thru_2_20)[17] <- "Product_SKU"
colnames(thru_2_20)[18] <- "Product_SKU_Name_10206_GNS"
colnames(thru_2_20)[19] <- "AS_Cases"
colnames(thru_2_20)[20] <- "AS_Order_Quantity"
colnames(thru_2_20)[21] <- "AS_Original_Order_Quantity"
colnames(thru_2_20)[22] <- "AS_Net_Pounds"
colnames(thru_2_20)[23] <- "AS_Variable_Margin"
colnames(thru_2_20)[24] <- "Pack_Size"
colnames(thru_2_20)[25] <- "Shipped_Date"
colnames(thru_2_20)[26] <- "Shipped_Year"
colnames(thru_2_20)[27] <- "Shipped_Month"
colnames(thru_2_20)[28] <- "Short_Ship_Reason"
colnames(thru_2_20)[29] <- "Geo_Location"
colnames(thru_2_20)[30] <- "Channel"
colnames(thru_2_20)[31] <- "Customer"
colnames(thru_2_20)[32] <- "MFG_Site"
colnames(thru_2_20)[33] <- "Shipping_Main"
colnames(thru_2_20)[34] <- "Shipping_Location"
colnames(thru_2_20)[35] <- "Site_Line"
colnames(thru_2_20)[36] <- "Item"
colnames(thru_2_20)[37] <- "Ship_Month"
colnames(thru_2_20)[38] <- "Ship_Year_Month"
colnames(thru_2_20)[39] <- "Sysco_BAPA_SKU"
colnames(thru_2_20)[40] <- "Label_w_Desc"
colnames(thru_2_20)[41] <- "Product_SKU_w_Description"
colnames(thru_2_20)[42] <- "COVID_PRIORITY_SKU"
colnames(thru_2_20)[43] <- "PRIORITY_SKU"
colnames(thru_2_20)[44] <- "MTO_SOP"
colnames(thru_2_20)[45] <- "Macro_Platform"
colnames(thru_2_20)[46] <- "Starch_Supply_Impacted"
colnames(thru_2_20)[47] <- "Starch_Impacted"

thru_2_20 %>% 
  dplyr::mutate(Case_Cuts         = AS_Cases - AS_Order_Quantity) %>% 
  dplyr::mutate(Percent_of_total  = AS_Net_Pounds) %>%
  dplyr::mutate(Case_Fill_percent = AS_Cases / AS_Order_Quantity) %>%  
  dplyr::mutate(Case_Cut_Original_Order_Qty = AS_Original_Order_Quantity - AS_Cases) -> thru_2_20

thru_2_20 %>% 
  dplyr::mutate(LBS_Short = (AS_Net_Pounds / AS_Cases) * Case_Cuts) %>% 
  dplyr::mutate(Total_Removed = AS_Original_Order_Quantity - AS_Order_Quantity) %>% 
  dplyr::mutate(Lost_Sales_due_to_Allocation = AS_Original_Order_Quantity - AS_Order_Quantity) %>% 
  dplyr::mutate(Original_Case_Cut = AS_Original_Order_Quantity - AS_Cases) %>% 
  dplyr::mutate(Original_Case_Fill = AS_Cases / AS_Original_Order_Quantity) -> thru_2_20



rbind(thru_2_13, thru_2_20) -> thru_2_20




###### remove is.na ######

thru_2_20$AS_Variable_Margine_Per_Pound[is.na(thru_2_20$AS_Variable_Margine_Per_Pound)] <- 0
thru_2_20$AS_Cases[is.na(thru_2_20$AS_Cases)] <- 0
thru_2_20$AS_Order_Quantity[is.na(thru_2_20$AS_Order_Quantity)] <- 0
thru_2_20$AS_Original_Order_Quantity[is.na(thru_2_20$AS_Original_Order_Quantity)] <- 0
thru_2_20$AS_Net_Pounds[is.na(thru_2_20$AS_Net_Pounds)] <- 0
thru_2_20$AS_Variable_Margin[is.na(thru_2_20$AS_Variable_Margin)] <- 0


save(thru_2_20, file = "thru_2_20.RData")
