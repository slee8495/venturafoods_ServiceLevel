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

# Read data



thru_2_27 <- read_excel("C:/Users/SLee/OneDrive - Ventura Foods/Ventura Work/SCE/Report/Root Cause Report/2022/Feb/2.27.2022/thru 2.27.xlsx", 
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


colnames(thru_2_27)[1] <- "Super_Customer_Number"
colnames(thru_2_27)[2] <- "Super_Customer_Name"
colnames(thru_2_27)[3] <- "Mfg_Location"
colnames(thru_2_27)[4] <- "Mfg_Location_Name"
colnames(thru_2_27)[5] <- "Ship_Location"
colnames(thru_2_27)[6] <- "Ship_Location_Name"
colnames(thru_2_27)[7] <- "Product_Label"
colnames(thru_2_27)[8] <- "Product_Label_Name"
colnames(thru_2_27)[9] <- "Product_Platform_Desc"
colnames(thru_2_27)[10] <- "Line_Number"
colnames(thru_2_27)[11] <- "Line_Description"
colnames(thru_2_27)[12] <- "AS_Variable_Margine_Per_Pound"
colnames(thru_2_27)[13] <- "Pack_Size_Name"
colnames(thru_2_27)[14] <- "Product_Category_Name_105"
colnames(thru_2_27)[15] <- "Base_Product"
colnames(thru_2_27)[16] <- "Base_Product_Name_10206"
colnames(thru_2_27)[17] <- "Product_SKU"
colnames(thru_2_27)[18] <- "Product_SKU_Name_10206_GNS"
colnames(thru_2_27)[19] <- "AS_Cases"
colnames(thru_2_27)[20] <- "AS_Order_Quantity"
colnames(thru_2_27)[21] <- "AS_Original_Order_Quantity"
colnames(thru_2_27)[22] <- "AS_Net_Pounds"
colnames(thru_2_27)[23] <- "AS_Variable_Margin"
colnames(thru_2_27)[24] <- "Pack_Size"
colnames(thru_2_27)[25] <- "Shipped_Date"
colnames(thru_2_27)[26] <- "Shipped_Year"
colnames(thru_2_27)[27] <- "Shipped_Month"
colnames(thru_2_27)[28] <- "Short_Ship_Reason"
colnames(thru_2_27)[29] <- "Geo_Location"
colnames(thru_2_27)[30] <- "Channel"
colnames(thru_2_27)[31] <- "Customer"
colnames(thru_2_27)[32] <- "MFG_Site"
colnames(thru_2_27)[33] <- "Shipping_Main"
colnames(thru_2_27)[34] <- "Shipping_Location"
colnames(thru_2_27)[35] <- "Site_Line"
colnames(thru_2_27)[36] <- "Item"
colnames(thru_2_27)[37] <- "Ship_Month"
colnames(thru_2_27)[38] <- "Ship_Year_Month"
colnames(thru_2_27)[39] <- "Sysco_BAPA_SKU"
colnames(thru_2_27)[40] <- "Label_w_Desc"
colnames(thru_2_27)[41] <- "Product_SKU_w_Description"
colnames(thru_2_27)[42] <- "COVID_PRIORITY_SKU"
colnames(thru_2_27)[43] <- "PRIORITY_SKU"
colnames(thru_2_27)[44] <- "MTO_SOP"
colnames(thru_2_27)[45] <- "Macro_Platform"
colnames(thru_2_27)[46] <- "Starch_Supply_Impacted"
colnames(thru_2_27)[47] <- "Starch_Impacted"

thru_2_27 %>% 
  dplyr::mutate(Case_Cuts         = AS_Cases - AS_Order_Quantity) %>% 
  dplyr::mutate(Percent_of_total  = AS_Net_Pounds) %>%
  dplyr::mutate(Case_Fill_percent = AS_Cases / AS_Order_Quantity) %>%  
  dplyr::mutate(Case_Cut_Original_Order_Qty = AS_Original_Order_Quantity - AS_Cases) -> thru_2_27

thru_2_27 %>% 
  dplyr::mutate(LBS_Short = (AS_Net_Pounds / AS_Cases) * Case_Cuts) %>% 
  dplyr::mutate(Total_Removed = AS_Original_Order_Quantity - AS_Order_Quantity) %>% 
  dplyr::mutate(Lost_Sales_due_to_Allocation = AS_Original_Order_Quantity - AS_Order_Quantity) %>% 
  dplyr::mutate(Original_Case_Cut = AS_Original_Order_Quantity - AS_Cases) %>% 
  dplyr::mutate(Original_Case_Fill = AS_Cases / AS_Original_Order_Quantity) -> thru_2_27


# Important index !!! ----

rbind(thru_2_20, thru_2_27) -> thru_2_27





###### remove is.na ######

thru_2_27$AS_Variable_Margine_Per_Pound[is.na(thru_2_27$AS_Variable_Margine_Per_Pound)] <- 0
thru_2_27$AS_Cases[is.na(thru_2_27$AS_Cases)] <- 0
thru_2_27$AS_Order_Quantity[is.na(thru_2_27$AS_Order_Quantity)] <- 0
thru_2_27$AS_Original_Order_Quantity[is.na(thru_2_27$AS_Original_Order_Quantity)] <- 0
thru_2_27$AS_Net_Pounds[is.na(thru_2_27$AS_Net_Pounds)] <- 0
thru_2_27$AS_Variable_Margin[is.na(thru_2_27$AS_Variable_Margin)] <- 0


save(thru_2_27, file = "thru_2_27.RData")



#####################################################################################################################################################
########################################################## Platforms on Allocation Report ###########################################################
#####################################################################################################################################################



## Getting weekly report ----
thru_2_27

thru_2_27 %>% 
  dplyr::filter(Short_Ship_Reason == "Y - Order Allocation Change" | Short_Ship_Reason == "Y - Order Allocation Change ACT") -> weekly_report

weekly_report %>% 
  dplyr::filter(Product_Platform_Desc == "PC CUP - 60MM" | Product_Platform_Desc == "PC CUP - 75MM"| 
                  Product_Platform_Desc == "PC POUCH <= 6 OZ, 4 SIDE SEAL") -> weekly_report


weekly_report$Case_Fill_percent -> case_fill_nan

replace(case_fill_nan, is.infinite(case_fill_nan) | is.nan(case_fill_nan) | is.na(case_fill_nan), 0) -> case_fill_nan
sprintf('%.2f', case_fill_nan) -> case_fill_nan
case_fill_nan <- as.double(case_fill_nan)
data.frame(case_fill_nan) -> case_fill_nan

weekly_report[, -ncol(weekly_report)] -> weekly_report
cbind(weekly_report, case_fill_nan) -> weekly_report
colnames(weekly_report)[ncol(weekly_report)] <- "Case_Fill_percent"


# case_fill_percent_report
reshape2::dcast(weekly_report, Product_Platform_Desc ~ Shipped_Year + Ship_Month, value.var = "AS_Cases", sum) -> case_fill_as_cases
reshape2::dcast(weekly_report, Product_Platform_Desc ~ Shipped_Year + Ship_Month, value.var = "AS_Order_Quantity", sum) -> case_fill_as_order_qty

cbind(case_fill_as_cases, case_fill_as_order_qty) -> a

data.frame(a[, 1]) -> Product_Platform_Desc
colnames(Product_Platform_Desc)[1] <- "Product_Platform_Desc"
a[, 2:9] -> cases_set
a[, 11:18] -> order_qty_set

cases_set / order_qty_set -> case_fill_rate
cases_set - order_qty_set -> case_cut

cbind(Product_Platform_Desc, case_fill_rate, case_cut) -> b

# Grand Total - case_cut
cbind(Product_Platform_Desc, case_cut) -> case_cut

case_cut %>% 
  janitor::adorn_totals("row") -> case_cut

# Grand Total - case_fill_rate
cases_set
order_qty_set


cbind(Product_Platform_Desc, cases_set) -> cases_set_1
cbind(Product_Platform_Desc, order_qty_set) -> order_qty_set_1

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


Product_Platform_Desc[nrow(Product_Platform_Desc) + 1, ] <- "Total"
cbind(Product_Platform_Desc, case_fill_total) -> case_fill_rate



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
    Product_Platform_Desc = gt::md("__Product Platform Desc__")
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
    Product_Platform_Desc = gt::md("__Product Platform Desc__")
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


