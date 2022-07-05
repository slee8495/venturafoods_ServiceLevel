## Data wrangling ##

thru_2_13

thru_2_13 %>% 
  dplyr::filter(Short_Ship_Reason == "Y - Order Allocation Change" | 
                  Short_Ship_Reason == "Y - Order Allocation ChangeACT") -> thru_2_13_report
  
thru_2_13_report %>% 
  dplyr::filter(Product_Platform_Desc == "PC POUCH <= 6 OZ, 4 SIDE SEAL" |
                  Product_Platform_Desc == "PC CUP - 60MM" |
                  Product_Platform_Desc == "PC CUP - 75MM") -> thru_2_13_report


reshape2::dcast(thru_2_13_report, Product_Platform_Desc ~ Shipped_Year + Shipped_Month , value.var = "Case_Cuts", sum)

reshape2::dcast(thru_2_13_report, Product_Platform_Desc ~ Shipped_Year + Shipped_Month , value.var = "Case_Fill_percent", sum)


# You Need to solve the problem for InF & NaN problem...to report. 


