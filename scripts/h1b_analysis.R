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

input_data_file_path <- "inputs/h1b_data_2024.xlsx" # Download the data here: https://www.uscis.gov/tools/reports-and-studies/h-1b-employer-data-hub 

metro_shp_file_path <- "C:/Users/ianwe/Downloads/shapefiles/2024/CBSAs/cb_2024_us_cbsa_500K.shp" # Download this shape file here: https://www2.census.gov/geo/tiger/GENZ2024/shp/cb_2024_us_cbsa_500k.zip
zip_metro_crosswalk_file_path <- "C:/Users/ianwe/Downloads/ZIP_CBSA_062025.xlsx" # Download the zip-to-metro crosswalk here: https://www.huduser.gov/apps/public/uspscrosswalk/home

output_file_path_for_tabular_data <- "outputs/h1b_approvals_and_denials_by_metro_2024.xlsx"

output_file_path_for_shape_file <- "outputs/h1b_approvals_and_denials_by_metro_2024.shp"

# Reading in the empty shape files ----

metro_shp <- st_read(metro_shp_file_path)

metro_shp_geo <- metro_shp %>%
  select(GEOID, geometry)

metro_shp_info <- metro_shp %>%
  st_drop_geometry() %>%
  select(-c(NAMELSAD, GEOIDFQ, CSAFP, CBSAFP, LSAD, ALAND, AWATER))

# Reading in crosswalk ----

zip_metro_crosswalk <- read.xlsx(zip_metro_crosswalk_file_path)

zip_metro_crosswalk <- zip_metro_crosswalk %>%
  select(ZIP, CBSA, BUS_RATIO) %>%
  janitor::clean_names()

# Reading in labor force data ----

census_api_key <- '6dd2c4143fc5f308c1120021fb663c15409f3757' # Provide the Census API Key, if others are running this you will need to get a Census API key here: https://api.census.gov/data/key_signup.html

acs_year <- 2024
acs_data_type <- 'acs1' # Define the survey to pull data from, 'acs5' for 5-year estimates, 'acs1' for 1 year estimates
geo_level_for_data_pull <- "cbsa" # Define the geography for the ACS data download. Other options include 'state', 'county', 'tract', 'block group', etc.
# See https://walker-data.com/tidycensus/articles/basic-usage.html#geography-in-tidycensus for a comprehensive list of geography options.
read_in_geometry <- FALSE # Change this to TRUE to pull in spatial data along with the data download 
show_api_call = TRUE # Show the call made to the Census API in the console, this will help if an error is thrown

# Load the variables for the year / dataset selected above
# acs_variables <- load_variables(year = 2024, dataset = acs_data_type)

# Read in the preferred variable spreadsheet 
variables <- read.xlsx("inputs/acs_variables_2024_acs1.xlsx", sheet = 'Labor Force')

# Select 'name' and 'amended_label' (and rename 'name' to code')
variables <- variables %>%
  select(name, amended_label) %>%
  rename(code = name)

# Create Codes, containing all of the preferred variable codes
variable_codes <- variables$code
# Create Labels, containing all of the amended labels
variable_labels <- variables$amended_label

labor_force_data <- get_acs(
  geography = geo_level_for_data_pull,
  variables = variable_codes,
  year = acs_year,
  geometry = read_in_geometry,
  key = census_api_key,
  survey = acs_data_type,
  show_call = show_api_call
)

labor_force_data <- labor_force_data %>%
  rename(code = variable) %>%
  left_join(variables, by = 'code') %>%
  rename(variable = amended_label) %>%
  select(-code) %>%
  pivot_wider(names_from = 'variable', values_from = 'estimate', id_cols = c('GEOID', 'NAME'))

labor_force_data <- labor_force_data %>%
  # Drop PR metros
  filter(!str_detect(NAME, pattern = ', PR Metro Area')) %>%
  mutate(clv_25_64 = civ_labor_force_25_64_less_than_hs+civ_labor_force_25_64_hs+civ_labor_force_25_64_some_college+civ_labor_force_25_64_bachelors) %>%
  select(GEOID, NAME, pop, clv_25_64)

# Reading in H1B data ----

data <- read.xlsx(input_data_file_path)

# Data analysis ----

data <- data %>%
  select(`Employer.(Petitioner).Name`, Petitioner.Zip.Code, Petitioner.City, Petitioner.State, Tax.ID, `Industry.(NAICS).Code`, 
         ends_with('Approval'), ends_with('Denial')) %>%
  rename(city = Petitioner.City, state = Petitioner.State, zip = Petitioner.Zip.Code) %>%
  janitor::clean_names()

data_condensed <- data %>%
  filter(!industry_naics_code %in% c('61 - Educational Services', '92 - Public Administration')) %>%
  group_by(zip, state, industry_naics_code) %>%
  summarize(across(new_employment_approval:amended_denial, ~sum(.))) %>%
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

data_condensed <- data_condensed %>%
  filter(zip != '-99') %>%
  left_join(zip_metro_crosswalk, by = 'zip')

data_condensed <- data_condensed %>%
  mutate(across(new_employment_approval:amended_denial, ~.*bus_ratio)) %>%
  select(-bus_ratio)

data_cbsa <- data_condensed %>%
  group_by(cbsa) %>%
  summarize(across(new_employment_approval:amended_denial, ~sum(.))) %>%
  ungroup()

data_cbsa <- data_cbsa %>%
  left_join(metro_shp_info, by = c('cbsa' = 'GEOID')) %>%
  select(NAME, cbsa, everything()) %>%
  mutate(across(new_employment_approval:amended_denial, ~round(., 0))) %>%
  select(-NAME)

data_cbsa <- labor_force_data %>%
  left_join(data_cbsa, by = c('GEOID' = 'cbsa')) 

data_cbsa <- data_cbsa %>%
  mutate(across(new_employment_approval:amended_denial, ~if_else(is.na(.), 0, .))) %>%
  mutate(NAME = str_remove(NAME, pattern = ' Metro Area'),
         NAME = str_remove(NAME, pattern = ' Micro Area'))

# Output tabular data ----

write.xlsx(data_cbsa, output_file_path_for_tabular_data)

# Output spatial data ----

names(data_cbsa) <- str_replace(names(data_cbsa), 
                                pattern = "^civ_labor_force_", 
                                replacement = "clv_")

data_cbsa <- data_cbsa %>%
  rename(new_app = new_employment_approval, 
         cont_app = continuation_approval, 
         chgs_app = change_with_same_employer_approval,
         conc_app = new_concurrent_approval,
         chg_app = change_of_employer_approval,
         am_app = amended_approval,
         new_den = new_employment_denial, 
         cont_den = continuation_denial,
         chgs_den = change_with_same_employer_denial,
         conc_den = new_concurrent_denial,
         chg_den = change_of_employer_denial,
         am_den = amended_denial
  ) %>%
  mutate(imp_app = new_app + conc_app,
         imp_shr = (imp_app / clv_25_64)*100)

data_cbsa_spatial <- data_cbsa %>%
  left_join(metro_shp_geo, by = 'GEOID') %>%
  st_as_sf()

arc.check_product()  

arc.write(path = output_file_path_for_shape_file, data = data_cbsa_spatial, overwrite = T, validate = T)