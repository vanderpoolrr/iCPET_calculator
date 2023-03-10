# iCPET_calculator: a web-based application to standardize the calculation of alpha distensibility
This is an R Shiny app to calculate alpha distensibility (1) from multi-beat pressure flow curves. Mean pulmonary artery pressure (mPAP), Pulmonary artery wedge pressure (PAWP) and cardiac output (CO) are typcially measured during a right heart catheterization.  

A web-based version of the iCPET calculator is available at [https://vanderpoolrr.shinyapps.io/iCPET_calculator/](https://vanderpoolrr.shinyapps.io/iCPET_calculator/). This link may not always be active. Alternatively, download the App.R file and run in RStudio locally. 

## Overview
The RShiny app for multi-point pressure flow analysis operates in a top down fashion. 
![](RShiny%20iCPET%20calculator.png)
### Step 1
The user will enter information about the study, the date of the measurements, when the analysis was completed and who complted the analysis. 
### Step 2
The user will then select the number of stages that were measured. This will resize the table based on the number of stages selected in by the slider. The user will then enter the Workload (W), Right Atrial Pressure (RAP), mean PA pressure (mPAP), PA wedge pressure (PW) and Cardiac Output (CO) for each stage. The checkboxes on the right are then used to select which exercise stages should be included in the calculations. 
### Step 3
The user then has the ability to verify the data that was entered into the table and see the pressure-flow results for mPAP/Q, PW/Q and RA/Q relationships in the generated figure. 

### Step 4 
In the final step, the user then has the option to download the data and analysis results as a Plot, stand-alone table or an Excel File. The 'Downloaded Excel' file includes the analysis data and the generated figure including alpha distensibility, slope, intercepts and R squared values for the fitted relationships. 
![](Example%20Excel%20Output.PNG)

## Citation
1. Linehan JH, Haworth ST, Nelin LD, Krenz GS, Dawson CA. A simple distensible vessel model for interpreting pulmonary vascular pressure-flow curves. J Appl Physiol. 1992 Sep;73(3):987–94. 

## iCPET R Shiny App
This app requires the following R packages: `shiny, tidyverse, ggpubr, cowplot, viridis, openxlsx`. 

The app has been tested on: 
```
R version 4.2.2 (2022-10-31 ucrt) 
Platform: x86_64-w64-mingw32/x64 (64-bit) 
Running under: Windows 10 x64 (build 19044) 
 
Matrix products: default 
 
locale: 
[1] LC_COLLATE=English_United States.utf8  LC_CTYPE=English_United States.utf8    LC_MONETARY=English_United States.utf8 
[4] LC_NUMERIC=C                           LC_TIME=English_United States.utf8     
 
attached base packages: 
[1] stats     graphics  grDevices utils     datasets  methods   base      
 
loaded via a namespace (and not attached): 
 [1] Rcpp_1.0.10        rstudioapi_0.14    magrittr_2.0.3     hms_1.1.2          xtable_1.8-4       R6_2.5.1           
 [7] rlang_1.0.6        fastmap_1.1.1      fansi_1.0.4        tools_4.2.2        utf8_1.2.3         miniUI_0.1.1.1     
[13] cli_3.6.0          htmltools_0.5.4    ellipsis_0.3.2     digest_0.6.31      tibble_3.1.8       lifecycle_1.0.3    
[19] shiny_1.7.4        tzdb_0.3.0         readr_2.1.4        colourpicker_1.2.0 later_1.3.0        htmlwidgets_1.6.1  
[25] vctrs_0.5.2        promises_1.2.0.1   glue_1.6.2         mime_0.12          compiler_4.2.2     pillar_1.8.1       
[31] httpuv_1.6.9       pkgconfig_2.0.3
```
