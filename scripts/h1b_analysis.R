# Packages ----

# Set the packages to read in
packages <- c("tidyverse", "tidycensus", "sf", "openxlsx", "arcgisbinding", "conflicted", "zoo")

# Function to check and install missing packages
install_if_missing <- function(package) {
  if (!requireNamespace(package, quietly = TRUE)) {
    install.packages(package, dependencies = TRUE)
  }
}

# Apply the function to each package
invisible(sapply(packages, install_if_missing))

# Load the packages
library(tidyverse)
library(tidycensus)
library(sf)
library(openxlsx)
library(arcgisbinding)
library(conflicted)
library(zoo)

# Prefer certain packages for certain functions
conflicts_prefer(dplyr::filter, dplyr::lag, lubridate::year, base::`||`, base::is.character, base::`&&`, stats::cor, base::as.numeric)

rm(install_if_missing, packages)

# Setting file paths ----

input_data_file_path <- "inputs/h1b_data_2024.xlsx"

metro_shp_file_path <- "C:/Users/ianwe/Downloads/shapefiles/2024/CBSAs/cb_2024_us_cbsa_500K.shp" # Input the file path for the shape file that you would like to read in. 

zip_metro_crosswalk_file_path <- "C:/Users/ianwe/Downloads/ZIP_CBSA_062025.xlsx"

output_file_path_for_tabular_data <- "outputs/h1b_approvals_and_denials_by_metro_2024.xlsx"

output_file_path_for_shape_file <- "outputs/h1b_approvals_and_denials_by_metro_2024.shp"

# Reading in the empty shape files ----

metro_shp <- st_read(metro_shp_file_path)

metro_shp_geo <- metro_shp %>%
  select(GEOID, geometry)

metro_shp_info <- metro_shp %>%
  st_drop_geometry() %>%
  select(-c(NAMELSAD, GEOIDFQ, CSAFP, CBSAFP, LSAD, ALAND, AWATER))

# Reading in data ----

data <- read.xlsx(input_data_file_path)

data <- data %>%
  select(`Employer.(Petitioner).Name`, Petitioner.Zip.Code, Petitioner.City, Petitioner.State, Tax.ID, `Industry.(NAICS).Code`, 
         New.Employment.Approval, New.Employment.Denial, Continuation.Approval, Continuation.Denial) %>%
  rename(city = Petitioner.City, state = Petitioner.State, zip = Petitioner.Zip.Code) %>%
  janitor::clean_names()

data_condensed <- data %>%
  group_by(zip, state) %>%
  summarize(across(new_employment_approval:continuation_denial, ~sum(.))) %>%
  ungroup() %>%
  filter(!state %in% c('PR', 'VI', "GU", "MP", "XX", ""))

data_condensed <- data_condensed %>%
  mutate(zip = case_when(
           is.na(zip) ~ '-99',
           T ~ zip),
         state = case_when(
           is.na(state) ~ '-99',
           T ~ state)
         )


# Reading in crosswalk ----

zip_metro_crosswalk <- read.xlsx(zip_metro_crosswalk_file_path)

zip_metro_crosswalk <- zip_metro_crosswalk %>%
  select(ZIP, CBSA, RES_RATIO) %>%
  janitor::clean_names()

# Data analysis ----

data_condensed <- data_condensed %>%
  filter(zip != '-99') %>%
  left_join(zip_metro_crosswalk, by = 'zip')

data_condensed <- data_condensed %>%
  mutate(across(new_employment_approval:continuation_denial, ~.*res_ratio)) %>%
  select(-res_ratio)

data_cbsa <- data_condensed %>%
  group_by(cbsa) %>%
  summarize(across(new_employment_approval:continuation_denial, ~sum(.))) %>%
  ungroup()

data_cbsa <- data_cbsa %>%
  left_join(metro_shp_info, by = c('cbsa' = 'GEOID')) %>%
  select(NAME, cbsa, everything()) %>%
  mutate(across(new_employment_approval:continuation_denial, ~round(., 0))) 

data_cbsa <- data_cbsa %>%
  rename(new_app = new_employment_approval, 
         cont_app = continuation_approval, 
         new_den = new_employment_denial, 
         cont_den = continuation_denial) %>%
  mutate(tot_app = new_app + cont_app,
         tot_den = new_den + cont_den,
         tot_init = new_app + new_den,
         tot_cont = cont_app + cont_den)

# Output tabular data ----

write.xlsx(data_cbsa, output_file_path_for_tabular_data)

# Output spatial data ----

data_cbsa_spatial <- data_cbsa %>%
  left_join(metro_shp_geo, by = c('cbsa' = 'GEOID')) %>%
  st_as_sf()

arc.check_product()  

arc.write(path = output_file_path_for_shape_file, data = data_cbsa_spatial, overwrite = T, validate = T)
