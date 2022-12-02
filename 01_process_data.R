# Copyright 2022 Province of British Columbia
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
# http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and limitations under the License.

library("here")
library("tidytable")
library("janitor")
library("wrapR") #  devtools::install_github("bcgov/wrapR")
library("readxl")

#read in dataframes--------------
jo_long <- vroom::vroom(here("raw_data",
                            list.files(here("raw_data"), pattern = "Raw JO", ignore.case = TRUE)),
                       locale = readr::locale(encoding = "latin1"),
                       skip=3,
                       col_select = -1)%>%
#  filter(Variable=="Job Openings")%>%
  pivot_longer(cols=-c("NOC","Description","Industry","Variable","Geographic Area"), names_to = "year", values_to = "value")

employment_long <- vroom::vroom(here("raw_data",
                                    list.files(here("raw_data"), pattern = "Raw Emp")),
                               locale = readr::locale(encoding = "latin1"),
                               skip=3,
                               col_select = -1)%>%
  pivot_longer(cols=-c("NOC","Description","Industry","Variable","Geographic Area"), names_to = "year", values_to = "value")

noc_mapping <- vroom::vroom(here::here("raw_data","noc_mapping.csv"))

construction_noc <- vroom::vroom(here::here("raw_data","Construction Trades.csv"), delim = ",")%>%
  pull(noc4)

industry_mapping <- vroom::vroom(here::here("raw_data","industry_to_agg_mapping.csv"), delim=",")%>%
  distinct()%>%
  clean_tbbl()

long <- bind_rows(jo_long, employment_long)%>%
  clean_tbbl()%>%
  full_join(noc_mapping)%>%
  full_join(industry_mapping)

for_heatmap <- long%>%
  filter(aggregate_industry=="construction",
         noc %in% construction_noc,
         geographic_area=="british_columbia")%>%
  select(noc4, industry, variable, year, value)%>%
  pivot_wider(names_from = variable, values_from = value)%>%
  mutate(jo_per_emp=round(100*job_openings/employment,1)
         )%>%
  select(-job_openings)%>%
  camel_to_title()

top_ten <- for_heatmap%>%
  group_by(industry, noc4)%>%
  summarize(employment=sum(employment))%>%
  arrange(industry, desc(employment))%>%
  top_n(employment, n=10)%>%
  select(industry, noc4)

for_heatmap <- semi_join(for_heatmap, top_ten) #top 10 occupations for each sub-industry

write_rds(for_heatmap, here("processed_data", "for_heatmap.rds"))
#load the stokes inputs---------------

path <- here("raw_data", "BritishColumbiaTables.xlsx")

indicators <- read_excel(path, sheet="Key Indicators", skip=2, na = "NA")%>%
  rename(name=`...1`)%>%
  pivot_longer(cols=-name, names_to = "year", values_to = "value")%>%
  na.omit(name)%>%
  filter(name!="% Change")%>%
  write_rds(here("processed_data", "indicators.rds"))

investment <- read_excel(path, sheet="Investment", skip=2, na = "NA")%>%
  rename(thing=`...1`)%>%
  pivot_longer(cols=-thing, names_to = "year", values_to = "value")%>%
  na.omit(thing)%>%
  filter(thing!="% Change")%>%
  write_rds(here("processed_data", "investment.rds"))

#load the historic employment-----------------

historic_employment <- read_excel(here("raw_data","Employment by 64 LMO industries for BC and regions, 1997-2021.xlsx"), skip=3, na = "NA")%>%
  pivot_longer(cols=-c(`Industry Code`, `Industry`), names_to = "year", values_to = "value")%>%
  clean_tbbl() %>%
  filter(str_detect(industry, "construction")| str_detect(industry,"contractors"))%>%
  select(-industry_code)%>%
  rename(name=industry)

#%>%
#  pivot_wider(id_cols=c(year), names_from = industry, values_from = value)%>%
#  rowwise() %>%
#  mutate(Construction = sum(across(!year)))%>%
#  pivot_longer(cols=-year, names_to = "name", values_to = "value")

#get forecast employment----------------

forecast_employment <- long%>%
  filter(aggregate_industry=="construction",
         noc=="#t",
         geographic_area=="british_columbia",
         variable=="employment")%>%
  select(name=industry, year, value)

#%>%
#  pivot_wider(names_from = industry, values_from = value)%>%
#  rowwise() %>%
#  mutate(Construction = sum(across(!year)))%>%
#  pivot_longer(cols=-year, names_to = "name", values_to = "value")

bind_rows(historic_employment, forecast_employment)%>%
  mutate(year=as.numeric(as.character(year)))%>%
  camel_to_title()%>%
  write_rds(here("processed_data", "construction_employment.rds"))

long%>%
  filter(aggregate_industry=="construction",
         noc=="#t",
         geographic_area=="british_columbia")%>%
  select(industry, name=variable, year, value)%>%
  pivot_wider(names_from = name, values_from = value)%>%
  mutate(job_openings_per_employment=round(100*job_openings/employment, 1))%>%
  pivot_longer(cols=-c(industry, year), names_to = "name", values_to = "value")%>%
  camel_to_title()%>%
  write_rds(here("processed_data", "construction.rds"))



