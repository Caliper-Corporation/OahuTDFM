# This function is not generalized, but will get the pertinent Census 
# data for Oahu.

load_census_data <- function(state = "HI", acs_year = 2018){
  
  library(tidycensus)
  library(tidyverse)
  # library(sf)
  
  county <- "Honolulu"

  bg_file <- "data/input/census_data/downloaded_files/acs_blockgroups.csv"
  if (!file.exists(bg_file)){
    # Read table of census variables and their names/geographies
    # https://api.census.gov/data/2018/acs/acs5/variables.html
    acs_vars <- read_csv("data/input/census_data/acs_bg_variables.csv")
    # Get census data from API
    acs_raw <- get_acs(
      geography = "block group", state = state, year = acs_year,
      county = county, variables = acs_vars$variable, geometry = FALSE
    )
    
    # Join variable names, sum any repeats, and then spread. Vehicle variable
    # names are repeated because the household numbers come from a table that
    # lists them by owner/renter.
    acs_bg <- acs_raw %>%
      as.data.frame() %>%
      left_join(acs_vars, by = "variable") %>%
      group_by(GEOID, name) %>%
      summarize(estimate = sum(estimate)) %>%
      spread(key = name, value = estimate) %>%
      # After review, the total vehicle estimates from table B25046 are missing
      # in some zones. Where it is missing, use a simple equation to calculate.
      mutate(
        veh_tot_temp = 
          veh0 * 0 + 
          veh1 * 1 + 
          veh2 * 2 + 
          veh3 * 3 + 
          veh4 * 4 + 
          veh5 * 5,
        veh_tot = ifelse(is.na(veh_tot), veh_tot_temp, veh_tot)
      ) %>%
      select(-veh_tot_temp)
    
    write_csv(acs_bg, bg_file)
  } else {
    acs_bg <- read_csv(bg_file)
  }
  
  # ACS tracts
  acs_file <- "data/input/census_data/downloaded_files/acs_tracts.shp"
  if (!file.exists(acs_file)){
    # Read table of census variables and their names/geographies
    acs_vars <- read_csv("data/input/census_data/acs_tract_variables.csv")
    # Get census data from API
    acs_raw <- get_acs(
      geography = "tract", state = state, year = acs_year,
      county = county, variables = acs_vars$variable, geometry = FALSE
    )
    
    # Join variable names, sum any repeats, and then spread
    acs_tract <- acs_raw %>%
      # as.data.frame() %>%
      left_join(acs_vars, by = "variable") %>%
      group_by(GEOID, name) %>%
      summarize(estimate = sum(estimate)) %>%
      spread(key = name, value = estimate)
    
    write_csv(acs_tract, acs_file)
  } else {
    acs_tract <- read_csv(acs_file)
  }
  
  # same thing for decennial block group variables
  dec_file <- "data/input/census_data/downloaded_files/dec_blockgroups.shp"
  if (!file.exists(dec_file)){
    dec_vars <- read_csv("data/input/census_data/dec_bg_variables.csv")
    decennial_raw <- get_decennial(
      year = 2010, geography = "block group", state = state,
      county = county, variables = dec_vars$variable, geometry = FALSE
    )
    # model_boundary <- st_transform(model_boundary, st_crs(decennial_raw))
    # decennial_shp <- decennial_raw[
    #   st_intersects(decennial_raw, model_boundary, sparse = FALSE), 
    #   ]
    dec_bg <- decennial_raw %>%
      # as.data.frame() %>%
      left_join(dec_vars, by = "variable") %>%
      select(-variable, -NAME) %>%
      spread(key = name, value = value)
    
    write_csv(dec_bg, dec_file)
  } else {
    dec_bg <- read_csv(dec_file)
  }
  
  # The decennial block file below is not needed.
  
  # dec_block_file <- "data/input/census_data/shapes/dec_blocks.shp"
  # if (!file.exists(dec_block_file)){
  #   block_vars <- data_frame(
  #     variable = c("H003002", "P016001", "P042001"),
  #     name = c("hh", "hh_pop", "gq_pop")
  #   )
  #   dec_raw <- get_decennial(
  #     year = 2010, geography = "block", state = state,
  #     county = counties, variables = block_vars$variable, geometry = TRUE
  #   )
  #   model_boundary <- st_transform(model_boundary, st_crs(dec_raw))
  #   dec_blocks <- dec_raw[st_intersects(dec_raw, model_boundary, sparse = FALSE), ]
  #   block_tbl <- dec_blocks %>%
  #     as.data.frame() %>%
  #     left_join(block_vars, by = "variable") %>%
  #     select(-variable, -NAME) %>%
  #     spread(key = name, value = value)
  #   dec_blocks <- st_as_sf(block_tbl) %>%
  #     st_make_valid()
  #   
  #   st_write(dec_blocks, dec_block_file)
  # } else {
  #   dec_blocks <- st_read(dec_block_file, quiet = TRUE)
  # }
  
  result <- list()
  result$acs_bg <- acs_bg
  result$acs_tract <- acs_tract
  result$dec_bg <- dec_bg
  # result$dec_block <- dec_blocks
  return(result)
}