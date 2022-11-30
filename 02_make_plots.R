#libraries------------------
library(RColorBrewer)
library(readr)
library(tidyverse)
library(here)
library(ggplot2)
library(lubridate)
library(patchwork)
library(RColorBrewer)
library(flextable)
#constants-----------------
to_stokes= "Source: J:\\Labour Economics\\BC Labour Market Outlook\\2022 Edition (2022-2032)\\Model Output (Stokes)\\FINAL Macro forecast (May 10, 2022)\\BritishColumbiaTables.xlsx"
#functions---------------------
has_completed <- function(tbbl){
  "Completed" %in% tbbl$project_status
}

has_started <- function(tbbl){
  "Construction started" %in% tbbl$project_status
}

get_cost <- function(tbbl){
  max(tbbl$estimated_cost, na.rm=TRUE)
}

get_net_jobs <- function(tbbl){
  mean(tbbl$construction_jobs, na.rm=TRUE)-mean(tbbl$operating_jobs, na.rm=TRUE)
}



construction_start_date <- function(tbbl){
  temp <- tbbl%>%
    filter(project_status=="Construction started")
  if(nrow(temp)>0){
    as.character(min(temp$last_update))
  }else{
    NA_character_
  }
}

completion_date <- function(tbbl){
  max(tbbl$standardized_completion_date, na.rm=TRUE)
}

make_plt <- function(tbbl, type, ttl, sub_ttl, cap){
  plt <- ggplot()+
   # geom_rect(aes(xmin=2022, xmax=2032, ymin=0, ymax=Inf), alpha=.1)+
    scale_y_continuous(labels=scales::comma)+
    scale_colour_brewer(palette = "Dark2")+
    theme_minimal()+
    expand_limits(y = 0)+
    labs(title=ttl,
         subtitle=sub_ttl,
         x="",
         y="",
         fill="",
         colour="",
         caption= cap)+
    theme(legend.position="bottom")
  if(type=="line"){
    plt <- plt +
      geom_line(data=tbbl,  mapping=aes(year, value, colour=name))
  }else if(type=="area"){
    plt <- plt +
      geom_area(data=tbbl,  mapping=aes(year, value, fill=fct_reorder(name, value)), alpha=.75)+
      scale_fill_brewer(palette="Dark2")
  }else{
    stop('type needs to be "line" or "area"')
  }
  plt+
    geom_rect(aes(xmin=2022, xmax=2032, ymin=0, ymax=Inf), alpha=.1)
}

#heatmaps--------------------------
tbbl<- read_rds(here("processed_data","for_heatmap.rds"))%>%
  mutate(discrete=cut(jo_per_emp, breaks=c(-100,seq(-8.5,-1, 1.5), seq(1,17,4),100),
             include.lowest=TRUE,
             right=FALSE))

plt <- ggplot(tbbl, aes(year,
                        noc4,
                        fill=discrete,
                        label=jo_per_emp
                         ))+
  geom_tile()+
  geom_text(aes(size=employment, alpha=employment))+
  scale_alpha(range = c(0.5, 1))+
  scale_size(range=c(.75,2.25))+
  scale_fill_brewer(palette = "RdYlBu", type = 'qual', direction = -1)+
  facet_wrap(~industry)+
  theme_minimal()+
  labs(title="Job Openings as a percent of Employment",
       x="",
       y="",
       alpha="Employment",
       size="Employment",
       fill="JO/Emp%",
       caption="Source: 2022 Labour Market Outlook")+
  theme(text = element_text(size = 9))

ggsave(file=here("out","static_heatmap.png"), width=16, height=9, dpi="retina", bg="white")


#plots based on data supplied to stokes-----------------------

indicators <- read_rds(here("processed_data","indicators.rds"))
investment <- read_rds(here("processed_data","investment.rds"))

housing_starts <- indicators%>%
  filter(name %in% c("Housing Starts (000s)","Household Formation (000s)"))%>%
  mutate(year=as.numeric(year))

residential_investment <- investment%>%
  filter(thing %in% c("Total Residential Investment","New Housing"))%>%
  mutate(year=as.numeric(year))%>%
  pivot_wider(names_from = thing, values_from = value)%>%
  mutate(`Difference between Total and New Housing`=`Total Residential Investment`-`New Housing`)%>%
  pivot_longer(cols=-year)%>%
  filter(name !="Total Residential Investment")

non_residential <- investment%>%
  filter(thing %in% c("Engineering Construction",
                      "Building Construction",
                      "Industrial Construction",
                      "Inst. & Gov't Construction",
                      "Commercial Construction"
                      ))%>%
  mutate(year=as.numeric(year))%>%
  pivot_wider(names_from = thing, values_from = value)%>%
  mutate(`ICI buildings`=`Industrial Construction`+`Inst. & Gov't Construction`+`Commercial Construction`)%>%
  pivot_longer(cols = -year, names_to = "name", values_to = "value")

housing_starts_plt <- make_plt(housing_starts,
         "line",
         "Residential Construction",
         "Housing Starts and Household Formations (units in 000s)",
         "")


residential_investment_plt <- make_plt(residential_investment,
         "area",
         "Residential Construction",
         "Total Residential Investment (2012 Millions $)",
         "")


engineering_vs_ici_plt <- non_residential%>%
  filter(name %in% c("Engineering Construction","ICI buildings"))%>%
  make_plt(
         "line",
         "Non-residential Construction",
         "Investment (2012 Millions $)",
         "")



ici_plt <- non_residential%>%
  filter(name %in% c("Industrial Construction",
                     "Inst. & Gov't Construction",
                     "Commercial Construction"))%>%
  make_plt(
    "line",
    "Non-residential Construction",
    "Investment: industrial, commercial and institutional (2012 Millions $)",
    "")

to_stokes_plts <-(housing_starts_plt+residential_investment_plt)/(engineering_vs_ici_plt+ici_plt)+
  plot_annotation(caption = to_stokes)
ggsave(file=here("out","stokes_inputs.png"), width=16, height=9, dpi="retina", bg="white")

#employment--------------------------

employment_plt <- read_rds(here("processed_data", "construction_employment.rds"))%>%
  make_plt("area",
           "Construction Employment",
           "Historic Data from LFS, Forecasts from 2022 LMO",
           "Source: LFS and LMO")

#construction forecast-----------------

construction_forecast_plt <- read_rds(here("processed_data", "construction.rds"))%>%
  filter(name %in% c("Job Openings","Employment","Expansion Demand", "Replacement Demand"))%>%
  ggplot(aes(year, value, fill=fct_reorder(industry, value)))+
    geom_col(position="dodge", alpha=.75)+
  scale_fill_brewer(palette="Dark2")+
  scale_y_continuous(labels=scales::comma)+
    facet_wrap(~name, scales = "free_y")+
  theme_minimal()+
  theme(legend.position="bottom")+
  labs(x="",y="",fill="",caption="Source: 2022 LMO")

lmo_plots <- employment_plt+construction_forecast_plt+ plot_layout(guides = "collect") & theme(legend.position = "bottom")


ggsave(file=here("out","lmo_plots.png"), width=16, height=9, dpi="retina", bg="white")

#mpi stuff-------------------------------

mpi <- read_csv(here("processed_data","mpi_clean.csv"))%>%
  janitor::clean_names()%>%
  group_by(project_id, project_name)%>%
  nest()%>%
  mutate(completed=map_lgl(data, has_completed),
         construction_started=map_lgl(data, has_started),
         construction_sd=map_chr(data, construction_start_date),
         construction_sd=lubridate::ymd(construction_sd),
         expected_completion=map_chr(data, completion_date),
         expected_completion=lubridate::yq(expected_completion),
         estimated_cost=map_dbl(data, get_cost),
         net_jobs=map_dbl(data, get_net_jobs),
         first_entry_date=map_chr(data, function(tbbl) as.character(tbbl$first_entry_date[1])),
         first_entry_date=lubridate::ymd(first_entry_date),
         entry_to_start_duration=as.numeric(construction_sd-first_entry_date)/365.25
  )

all_waits <- ggplot(mpi, aes(entry_to_start_duration))+
  geom_histogram(aes(y = after_stat(count / sum(count))), bins=30)+
  theme_minimal()+
  scale_y_continuous(labels=scales::percent)+
  labs(x="Years",
       y="",
       title="Histogram of time spent in MPI pre-construction.",
       subtitle="Many projects only enter MPI after construction commences: time in MPI pre construction = 0",
       caption="Source:  BC Major Project Inventory")

positive_waits <- ggplot(filter(mpi, entry_to_start_duration>0), aes(entry_to_start_duration))+
  geom_histogram(aes(y = after_stat(count / sum(count))), bins=29)+
  theme_minimal()+
  scale_y_continuous(labels=scales::percent)+
  labs(x="Years",
       y="",
       title="Histogram of time spent in MPI pre-construction.",
       subtitle="Positive value only: Median 1.5 years, mean = 2.5 years",
       caption="Source: BC Major Project Inventory")

waiting_time_MPI <- all_waits+positive_waits
ggsave(file=here("out","waiting_time_MPI.png"), width=16, height=9, dpi="retina", bg="white")

mpi_for_lm <- mpi%>%
  filter(net_jobs>0)

mod <- lm(I(log(net_jobs))~I(log(estimated_cost)), data=mpi_for_lm)
mpi$predict_jobs <- exp(predict(mod, newdata = mpi))

(plt0 <- ggplot(mpi, aes(net_jobs, predict_jobs))+
  geom_point()+
  geom_abline(slope=1,intercept = 0)+
  scale_x_continuous(trans="log", labels=scales::comma)+
  scale_y_continuous(trans="log", labels=scales::comma)+
  labs(title="Problem: net jobs only available for 7% of MPI projects.",
      subtitle=expression("Solution: predict with log(net jobs)="~alpha~"+"~beta~"*log(estimated cost) +"~epsilon),
       x="Observed net jobs",
       y="Predicted net jobs",
      caption="Source: all MPI projects that report net jobs")+
  theme_minimal())


not_completed <- mpi%>%
  filter(completed==FALSE,
         construction_started==TRUE,
         is.finite(estimated_cost)
         )%>%
  ungroup()%>%
  arrange(desc(estimated_cost))


value_by_expected_completion <- not_completed%>%
  mutate(year_of_completion=lubridate::year(expected_completion))%>%
  group_by(year_of_completion)%>%
  summarize(estimated_cost=sum(estimated_cost, na.rm=TRUE),
            net_jobs=sum(net_jobs, na.rm=TRUE),
            predict_jobs=sum(predict_jobs, na.rm = TRUE)
            )

plt1 <- ggplot(value_by_expected_completion, aes(year_of_completion, estimated_cost))+
  geom_line()+
  geom_point()+
  scale_y_continuous(labels=function(x) scales::dollar(x, suffix = 'M'))+
  theme_minimal()+
  labs(title="Total value of completed projects",
       x="Expected Completion Year",
       y="",
       caption="Source: All incomplete MPI projects where construction has started.")

plt2 <- ggplot(value_by_expected_completion, aes(year_of_completion, predict_jobs))+
  geom_line()+
  geom_point()+
  scale_y_continuous(labels=scales::comma)+
  theme_minimal()+
  labs(title="Net job losses due to completion of projects",
       subtitle=expression("Predicted by log(net jobs)="~alpha~"+"~beta~"*log(estimated cost) +"~epsilon),
       x="Expected Completion Year",
       y="",
       caption="Source: All incomplete MPI projects where construction has started.")

mpi_plot <- plt0+plt1+plt2
ggsave(file=here("out","mpi_job_loss.png"), width=16, height=9, dpi="retina", bg="white")

#what proportion of projects that have not started construction get completed within 10 years?

mpi_long <- readxl::read_excel(here::here("raw_data","MPIlongraw.xlsx"),
                               col_types=c('numeric',
                               'text',
                               'numeric',
                               'text',
                               'text',
                               'text',
                               'text',
                               'text',
                               'date',
                               'numeric',
                               'text',
                               'date'))%>%
  janitor::clean_names()


#projects prior to 2011 where construction had not yet started.

not_started_prior_2011 <- mpi_long%>%
  filter(last_update<lubridate::ymd("2021-12-01")-years(10))%>%
  group_by(project_id)%>%
  nest()%>%
  mutate(construction_started=map_lgl(data, has_started))%>%
  filter(construction_started==FALSE)%>%
  select(project_id)

status_post_2011 <- mpi_long%>%
  filter(last_update>=lubridate::ymd("2021-12-01")-years(10))%>%
  group_by(project_id)%>%
  nest()%>%
  semi_join(not_started_prior_2011)%>%
  mutate(construction_started=map_lgl(data, has_started),
         category=map_chr(data, function(tbbl) tbbl$project_category_name[1]))%>%
  group_by(category)%>%
  summarize(eventually_started=sum(construction_started),
            total=n(),
            percentage=eventually_started/total)

ft <- status_post_2011%>%
  mutate(percentage=scales::percent(percentage, accuracy = 1))
colnames(ft) <- wrapR::make_title(colnames(ft))

ft <- flextable(ft)
ft <- add_footer_lines(ft, "Source: BC Major Project Inventory")
ft <- set_caption(ft, caption = "Unstarted Major Projects (pre 2011)")
ft
save_as_docx(ft, path=here::here("out","MPI_construction_historic_start_rates.docx"))

#Unstarted projects currently in MPI------------------


