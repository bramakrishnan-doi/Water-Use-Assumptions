## This code will generate the Water Use Assumptions Document in your current working directory.
## change it using setwd() accordingly

setwd("C:/Temp/Water-Use-Assumptions")

library(quarto)

## Set these parameters ###
param_scenario <- "Most"
param_mon_year <- "July 2025"
proj_state_use_file <- "data/Projected State Use -JUL25.xlsx"
sct_data_file <- "data/JUL25-Most.xlsx"

## This code creates the output file name
desired_output_filename <- paste0('24-MS LB Water Use Projections - ',
                                  param_mon_year, ' ',
                                  ifelse(param_scenario == 'Most', 
                                         paste0(param_scenario, ' Probable.docx'), 
                                         paste0('Probable ', param_scenario, '.docx')))


# --- Render the Quarto document ---
quarto_render(
  input = "Water Use Assumptions_v5.qmd",  
  output_file = desired_output_filename,
  execute_params = list(
    scenario = param_scenario,
    mon_year = param_mon_year,
    proj_state_use = proj_state_use_file,
    sct_data = sct_data_file
    
  )
)
