library(readxl)
library(base)
library(tidyverse)
library(ggrepel)
library(kableExtra)
library(rio)
library(gridExtra)
library(magick)
library(htmlTable)

rm(list=ls())
webshot::install_phantomjs()


path = "/Users/alex.rottinghaus/DataGripProjects/Fantasy Football/FantasyFootball.xlsx"

sheets_to_read <- excel_sheets(path) ##Obtain the sheets in the workbook

list_with_sheets <- lapply(sheets_to_read,
                           function(i)read_excel(path, sheet = i))

names(list_with_sheets) <- sheets_to_read ##Add names

list2env(list_with_sheets, .GlobalEnv)

alldat <- import_list(path, rbind = T)

Weeks = list()

weekFilter <- function(dat, week){
  newDat <- dat %>% 
    filter(Week == week) %>% 
    arrange(desc(W), desc(CumulativeScore)) %>% 
    select(Name, CumulativeScore, W, L) %>% 
    mutate(Rank = seq(1,10,1))
  
  return(newDat)
}

View(alldat)
for(i in 1:14){
  Weeks[i] <- weekFilter(alldat, i)
}

rankDat <- data.frame(rank = seq(1,10,1),
                      Week1 = Weeks[1],
                      Week2 = Weeks[2],
                      Week3 = Weeks[3],
                      Week4 = Weeks[4],
                      Week5 = Weeks[5],
                      Week6 = Weeks[6],
                      Week7 = Weeks[7],
                      Week8 = Weeks[8],
                      Week9 = Weeks[9],
                      Week10 = Weeks[10],
                      Week11 = Weeks[11],
                      Week12 = Weeks[12],
                      Week13 = Weeks[13],
                      Week14 = Weeks[14])

colnames(rankDat) <- c('rank','Week1','Week2','Week3','Week4','Week5','Week6','Week7','Week8','Week9','Week10','Week11','Week12','Week13','Week14')

long.dat <- rankDat %>%
  gather(key = "Week", value = "Team", -rank) %>% 
  mutate(week = as.numeric(str_extract(Week, "[0-9]+")))

View(long.dat)

data_ends <- long.dat %>% filter(week ==14) %>% pull(rank, Team)
data_begin <- long.dat %>% filter(week ==1) %>% pull(rank, Team)

rankflow <- ggplot(data = long.dat, aes(x = week, y = rank)) +
  labs(title = "The Professionals 2022 Weekly Ranking Chart", x = "Week", y = "Rank") +
  geom_line(aes(color = Team)) +
  scale_y_reverse(breaks = data_begin, sec.axis = sec_axis(~ .,breaks = data_ends))+
  scale_x_continuous(breaks = seq(1,14,1)) +
  theme(legend.position = "none",
        plot.title = element_text(hjust = 0.5, colour = "#6BA2B8", face = "bold"),
        axis.title.x = element_text(colour = "#6BA2B8", face = "bold"),
        axis.title.y = element_text(colour = "#6BA2B8", face = "bold"),
        axis.ticks = element_line(colour = "#6BA2B8"),
        axis.text = element_text(colour = "#6BA2B8"),
        panel.background = element_rect(fill = "white"),
        plot.background = element_rect(fill = "white"),
        panel.grid.major.x = element_line(colour = "#6BA2B8",linetype = 'dotted'))

rankflow
ggsave("/Users/alex.rottinghaus/DataGripProjects/Fantasy Football/Charts/RankLineGraph.png", rankflow)

scoreflow <- ggplot(data = alldat) +
  labs(title = "The Professionals 2022 Weekly Score Chart", x = "Week", y = "Score") +
  geom_point(aes(x = Week, y = Score, colour = Name))+
  geom_line(aes(x = Week, y = Score, colour = Name))+
  #stat_smooth(aes(x = Week, y = Score, colour = Name), method = "lm", formula = y ~ poly(x, 10), se = FALSE)+
  scale_y_continuous()+
  scale_x_continuous(breaks = seq(1,14,1)) +
  theme(plot.title = element_text(hjust = 0.5, colour = "#6BA2B8", face = "bold"),
        axis.title.x = element_text(colour = "#6BA2B8", face = "bold"),
        axis.title.y = element_text(colour = "#6BA2B8", face = "bold"),
        axis.ticks = element_line(colour = "#6BA2B8"),
        axis.text = element_text(colour = "#6BA2B8"),
        panel.background = element_rect(fill = "white"),
        plot.background = element_rect(fill = "white"),
        panel.grid.major.x = element_line(colour = "#6BA2B8",linetype = 'dotted'))
scoreflow
ggsave("/Users/alex.rottinghaus/DataGripProjects/Fantasy Football/Charts/ScoreFlow.png", scoreflow)

scoreflowFacet <- ggplot(data = alldat) +
  labs(title = "Turnover Machines 2022 Weekly Score Chart", x = "Week", y = "Score") +
  geom_point(aes(x = Week, y = Score), color = "#6BA2B8")+
  geom_line(aes(x = Week, y = Score), color = "#6BA2B8")+
  facet_wrap(vars(Name), ncol = 3)+
  #stat_smooth(aes(x = Week, y = Score, colour = Name), method = "lm", formula = y ~ poly(x, 10), se = FALSE)+
  scale_y_continuous(breaks = seq(100, 250, 10), limits = c(100, 250))+
  scale_x_continuous(breaks = seq(1,14,1)) +
  theme(plot.title = element_text(hjust = 0.5, colour = "#6BA2B8", face = "bold"),
        axis.title.x = element_text(colour = "#6BA2B8", face = "bold"),
        axis.title.y = element_text(colour = "#6BA2B8", face = "bold"),
        axis.ticks = element_line(colour = "#6BA2B8"),
        axis.text = element_text(colour = "#6BA2B8"),
        panel.background = element_rect(fill = "white"),
        plot.background = element_rect(fill = "white"),
        panel.grid.major.x = element_line(colour = "#6BA2B8",linetype = 'dotted'))

scoreflowFacet
ggsave("/Users/alex.rottinghaus/DataGripProjects/Fantasy Football/Charts/scoreflowFacet.png", scoreflowFacet)

flowCharts <- function(dat, teamNum){
  flow <- ggplot(data = dat %>% filter(Name == sheets_to_read[[teamNum]])) +
    labs(title = paste0(sheets_to_read[[teamNum]], " 2022 Weekly Score Chart"), x = "Week", y = "Score") +
    geom_point(aes(x = Week, y = Score), color = "#6BA2B8")+
    geom_line(aes(x = Week, y = Score), color = "#6BA2B8")+
    scale_y_continuous(breaks = seq(100, 250, 10), limits = c(100, 250))+
    scale_x_continuous(breaks = seq(1,14,1)) +
    theme(plot.title = element_text(hjust = 0.5, colour = "#6BA2B8", face = "bold"),
          axis.title.x = element_text(colour = "#6BA2B8", face = "bold"),
          axis.title.y = element_text(colour = "#6BA2B8", face = "bold"),
          axis.ticks = element_line(colour = "#6BA2B8"),
          axis.text = element_text(colour = "#6BA2B8"),
          panel.background = element_rect(fill = "white"),
          plot.background = element_rect(fill = "white"),
          panel.grid.major.x = element_line(colour = "#6BA2B8",linetype = 'dotted'))
  
  filename <- paste0(sheets_to_read[[teamNum]], "_scoreflow.png")
  
  ggsave(paste0("/Users/alex.rottinghaus/DataGripProjects/Fantasy Football/Charts/", filename), flow)
  
  return(flow)
}

for(i in 1:length(sheets_to_read)){
  flowCharts(alldat, i)
}

scoredistro <- alldat %>% ggplot(aes(x = Score)) +
  geom_histogram(binwidth = 5,boundary = 0, fill = "#6BA2B8")+
  labs(title = "Score Distribution 2022") +
  scale_y_continuous(breaks = seq(0,12,2)) +
  scale_x_continuous(breaks = seq(80,280,20), limits = c(100,250)) +
  theme(legend.position = "none",
        plot.title = element_text(hjust = 0.5, colour = "#6BA2B8", face = "bold"),
        axis.title.x = element_text(colour = "#6BA2B8", face = "bold"),
        axis.title.y = element_text(colour = "#6BA2B8", face = "bold"),
        axis.ticks = element_line(colour = "#6BA2B8"),
        axis.text = element_text(colour = "#6BA2B8"),
        panel.background = element_rect(fill = "white"),
        plot.background = element_rect(fill = "white"),
        panel.grid.major.x = element_line(colour = "#6BA2B8",linetype = 'dotted'))

ggsave("/Users/alex.rottinghaus/DataGripProjects/Fantasy Football/Charts/scoredistro.png", scoredistro)

sumTab <- alldat %>% group_by(Name) %>% 
  summarise(Mean = mean(Score), 
            Median = median(Score), 
            Min = min(Score), 
            Max= max(Score),
            'StdDev' = sd(Score)) %>%
  arrange(desc(Mean)) %>%
  as.data.frame(.) %>%
  mutate_if(is.numeric, round, 2)

sumTab

htmlTable(sumTab %>% kable(align = 'l', "html") %>% kable_styling(full_width = F) %>% save_kable(file = "/Users/alex.rottinghaus/DataGripProjects/Fantasy Football/Charts/summary.png"))
