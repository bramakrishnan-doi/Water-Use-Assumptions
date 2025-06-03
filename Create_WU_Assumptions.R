## This code will generate the Water Use Assumptions Document in your current working directory.
## change it using setwd() accordingly

setwd("C:/Temp/Water-Use-Assumptions")

library(quarto)

param_scenario <- "Min"
param_mon_year <- "April 2025"

desired_output_filename <- paste0('24-MS LB Water Use Projections - ',
                                  param_mon_year, ' ',
                                  ifelse(param_scenario == 'Most', 
                                         paste0(param_scenario, ' Probable.docx'), 
                                         paste0('Probable ', param_scenario, '.docx')))


# --- Render the Quarto document ---
quarto_render(
  input = "Water Use Assumptions.qmd",  
  output_file = desired_output_filename,
  execute_params = list(
    scenario = param_scenario,
    mon_year = param_mon_year
    
  )
)
