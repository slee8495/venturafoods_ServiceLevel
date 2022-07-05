library(tidyverse)
library(readxl)
library(writexl)
library(reshape2)
library(openxlsx)
library(readxlsb)
library(data.table)
library(janitor)
library(tidyquant)
library(broom)
library(gt)
library(fontawesome)
library(htmltools)



863285


# delete current month (if 1 week range includes at least one day, delete whole month. and pull them again.)
# Deleting current month ----

# when there are 2 months to delete. use below code (Please double check on the total col)


# thru_3_6 %>% 
#  dplyr::filter(Ship_Year_Month != "2022 - (02) Feb" | Ship_Year_Month != "2022 - (03) Mar" ) -> thru_3_6


thru_4_3 %>% 
  dplyr::filter(Ship_Year_Month != "2022 - (04) Apr") -> thru_4_3


# Read data

thru_4_10 <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Report/Root Cause Report/2022/Apr/4.10.2022/thru_4_10.xlsx", 
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



colnames(thru_4_10)[1] <- "Super_Customer_Number"
colnames(thru_4_10)[2] <- "Super_Customer_Name"
colnames(thru_4_10)[3] <- "Mfg_Location"
colnames(thru_4_10)[4] <- "Mfg_Location_Name"
colnames(thru_4_10)[5] <- "Ship_Location"
colnames(thru_4_10)[6] <- "Ship_Location_Name"
colnames(thru_4_10)[7] <- "Product_Label"
colnames(thru_4_10)[8] <- "Product_Label_Name"
colnames(thru_4_10)[9] <- "ProductPlatformDesc"
colnames(thru_4_10)[10] <- "Line_Number"
colnames(thru_4_10)[11] <- "Line_Description"
colnames(thru_4_10)[12] <- "AS_Variable_Margine_Per_Pound"
colnames(thru_4_10)[13] <- "Pack_Size_Name"
colnames(thru_4_10)[14] <- "Product_Category_Name_105"
colnames(thru_4_10)[15] <- "Base_Product"
colnames(thru_4_10)[16] <- "Base_Product_Name_10206"
colnames(thru_4_10)[17] <- "Product_SKU"
colnames(thru_4_10)[18] <- "Product_SKU_Name_10206_GNS"
colnames(thru_4_10)[19] <- "AS_Cases"
colnames(thru_4_10)[20] <- "AS_Order_Quantity"
colnames(thru_4_10)[21] <- "AS_Original_Order_Quantity"
colnames(thru_4_10)[22] <- "AS_Net_Pounds"
colnames(thru_4_10)[23] <- "AS_Variable_Margin"
colnames(thru_4_10)[24] <- "Pack_Size"
colnames(thru_4_10)[25] <- "Shipped_Date"
colnames(thru_4_10)[26] <- "Shipped_Year"
colnames(thru_4_10)[27] <- "Shipped_Month"
colnames(thru_4_10)[28] <- "Short_Ship_Reason"
colnames(thru_4_10)[29] <- "Geo_Location"
colnames(thru_4_10)[30] <- "Channel"
colnames(thru_4_10)[31] <- "Customer"
colnames(thru_4_10)[32] <- "MFG_Site"
colnames(thru_4_10)[33] <- "Shipping_Main"
colnames(thru_4_10)[34] <- "Shipping_Location"
colnames(thru_4_10)[35] <- "Site_Line"
colnames(thru_4_10)[36] <- "Item"
colnames(thru_4_10)[37] <- "Ship_Month"
colnames(thru_4_10)[38] <- "Ship_Year_Month"
colnames(thru_4_10)[39] <- "Sysco_BAPA_SKU"
colnames(thru_4_10)[40] <- "Label_w_Desc"
colnames(thru_4_10)[41] <- "Product_SKU_w_Description"
colnames(thru_4_10)[42] <- "COVID_PRIORITY_SKU"
colnames(thru_4_10)[43] <- "PRIORITY_SKU"
colnames(thru_4_10)[44] <- "MTO_SOP"
colnames(thru_4_10)[45] <- "Macro_Platform"
colnames(thru_4_10)[46] <- "Starch_Supply_Impacted"
colnames(thru_4_10)[47] <- "Starch_Impacted"




thru_4_10 %>% 
  dplyr::mutate(Case_Cuts         = AS_Cases - AS_Order_Quantity) %>% 
  dplyr::mutate(Percent_of_total  = AS_Net_Pounds) %>%
  dplyr::mutate(Case_Fill_percent = AS_Cases / AS_Order_Quantity) %>%  
  dplyr::mutate(Case_Cut_Original_Order_Qty = AS_Original_Order_Quantity - AS_Cases) -> thru_4_10

thru_4_10 %>% 
  dplyr::mutate(LBS_Short = (AS_Net_Pounds / AS_Cases) * Case_Cuts) %>% 
  dplyr::mutate(Total_Removed = AS_Original_Order_Quantity - AS_Order_Quantity) %>% 
  dplyr::mutate(Lost_Sales_due_to_Allocation = AS_Original_Order_Quantity - AS_Order_Quantity) %>% 
  dplyr::mutate(Original_Case_Cut = AS_Original_Order_Quantity - AS_Cases) %>% 
  dplyr::mutate(Original_Case_Fill = AS_Cases / AS_Original_Order_Quantity) -> thru_4_10


# Important index !!! ----

rbind(thru_4_3, thru_4_10) -> thru_4_10





###### remove is.na ######

thru_4_10$AS_Variable_Margine_Per_Pound[is.na(thru_4_10$AS_Variable_Margine_Per_Pound)] <- 0
thru_4_10$AS_Cases[is.na(thru_4_10$AS_Cases)] <- 0
thru_4_10$AS_Order_Quantity[is.na(thru_4_10$AS_Order_Quantity)] <- 0
thru_4_10$AS_Original_Order_Quantity[is.na(thru_4_10$AS_Original_Order_Quantity)] <- 0
thru_4_10$AS_Net_Pounds[is.na(thru_4_10$AS_Net_Pounds)] <- 0
thru_4_10$AS_Variable_Margin[is.na(thru_4_10$AS_Variable_Margin)] <- 0

thru_4_10[-1,] -> thru_4_10
save(thru_4_10, file = "thru_4_10.RData")





#####################################################################################################################################################
########################################################## Platforms on Allocation Report ###########################################################
#####################################################################################################################################################



## Getting weekly report ----
thru_4_10

thru_4_10 %>% 
  dplyr::filter(Short_Ship_Reason == "Y - Order Allocation Change" | Short_Ship_Reason == "Y - Order Allocation Change ACT") -> weekly_report

weekly_report %>% 
  dplyr::filter(ProductPlatformDesc == "PC CUP - 60MM" | ProductPlatformDesc == "PC CUP - 75MM"| 
                  ProductPlatformDesc == "PC POUCH <= 6 OZ, 4 SIDE SEAL") -> weekly_report


weekly_report$Case_Fill_percent -> case_fill_nan

replace(case_fill_nan, is.infinite(case_fill_nan) | is.nan(case_fill_nan) | is.na(case_fill_nan), 0) -> case_fill_nan
sprintf('%.2f', case_fill_nan) -> case_fill_nan
case_fill_nan <- as.double(case_fill_nan)
data.frame(case_fill_nan) -> case_fill_nan

weekly_report[, -ncol(weekly_report)] -> weekly_report
cbind(weekly_report, case_fill_nan) -> weekly_report
colnames(weekly_report)[ncol(weekly_report)] <- "Case_Fill_percent"


# case_fill_percent_report
reshape2::dcast(weekly_report, ProductPlatformDesc ~ Shipped_Year + Ship_Month, value.var = "AS_Cases", sum) -> case_fill_as_cases
reshape2::dcast(weekly_report, ProductPlatformDesc ~ Shipped_Year + Ship_Month, value.var = "AS_Order_Quantity", sum) -> case_fill_as_order_qty

names(case_fill_as_cases) <- str_replace_all(names(case_fill_as_cases), c("_" = "."))



cbind(case_fill_as_cases, case_fill_as_order_qty) -> a

data.frame(a[, 1]) -> ProductPlatformDesc
colnames(ProductPlatformDesc)[1] <- "ProductPlatformDesc"


a %>% 
  dplyr::select(contains(".")) -> cases_set

a %>% 
  dplyr::select(contains("_")) -> order_qty_set



cases_set / order_qty_set -> case_fill_rate
cases_set - order_qty_set -> case_cut

cbind(ProductPlatformDesc, case_fill_rate, case_cut) -> b

# Grand Total - case_cut
cbind(ProductPlatformDesc, case_cut) -> case_cut

case_cut %>% 
  janitor::adorn_totals("row") -> case_cut

# Grand Total - case_fill_rate
cases_set
order_qty_set


cbind(ProductPlatformDesc, cases_set) -> cases_set_1
cbind(ProductPlatformDesc, order_qty_set) -> order_qty_set_1

cases_set_1 %>% 
  janitor::adorn_totals("row") -> cases_set_1

cases_set_1 %>% 
  janitor::adorn_totals("col") -> cases_set_1


order_qty_set_1 %>% 
  janitor::adorn_totals("row") -> order_qty_set_1

order_qty_set_1 %>% 
  janitor::adorn_totals("col") -> order_qty_set_1

cases_set_1[, -1] -> cases_set_1
order_qty_set_1[, -1] -> order_qty_set_1

cases_set_1 / order_qty_set_1 -> case_fill_total


ProductPlatformDesc[nrow(ProductPlatformDesc) + 1, ] <- "Total"
cbind(ProductPlatformDesc, case_fill_total) -> case_fill_rate



case_cut
case_fill_rate

cbind(case_fill_rate, case_cut) -> Platforms_on_allocation_report





# Add total column (case_cut)

case_cut %>% 
  janitor::adorn_totals("col") -> case_cut

# Add total column (case_fill_rate)

case_fill_rate




############################################################# Visualization ####################################################################

case_cut %>%
  gt::gt() %>%
  gt::tab_header(title = gt::md("__Sum of Case Cuts__")) %>%
  gt::tab_style(
    style = gt::cell_text(size = px(15)),
    locations = gt::cells_body()
  ) %>% 
  cols_label(
    ProductPlatformDesc = gt::md("__Product Platform Desc__")
  ) %>% 
  tab_style(
    style = cell_text(size = px(16)),
    locations = cells_body(
      rows = nrow(case_cut)
    )
  ) %>% 
  tab_style(
    style = list(
      cell_text(color = "blue"),
      cell_text(weight = "bold")),
    locations = cells_body(
      rows = nrow(case_cut)
    )
  ) -> case_cut_viz



case_fill_rate %>% 
  gt::gt() %>% 
  gt::tab_header(title = gt::md("__Case Fill Rate__")) %>% 
  gt::tab_style(
    style = gt::cell_text(size = px(15)),
    locations = gt::cells_body()
  ) %>% 
  cols_label(
    ProductPlatformDesc = gt::md("__Product Platform Desc__")
  ) %>% 
  gt::fmt_percent(
    2:ncol(case_fill_rate)
  ) %>% 
  tab_style(
    style = cell_text(size = px(16)),
    locations = cells_body(
      rows = nrow(case_cut)
    )
  ) %>% 
  tab_style(
    style = list(
      cell_text(color = "blue"),
      cell_text(weight = "bold")),
    locations = cells_body(
      rows = nrow(case_cut)
    )
  ) %>% 
  tab_style(
    style = list(
      cell_fill(color = "#FFCCCB"),
      cell_text(weight = "bold"),
      cell_text(color = "#420C09")),
    locations = cells_body(
      columns = Total,
      rows = Total < 0.98
    )
  ) -> case_fill_rate_viz


case_cut_viz
case_fill_rate_viz

