#packages
libraries <-  c("readxl","readxlsb","dplyr","ggplot2","ggpubr")
lapply(libraries,require,character.only=TRUE)

#data import
#formula/process
setwd("C:/Users/deboodt.j/Procter and Gamble/Theia_BIC - Documents/Master files/1_Formulation and Process")
Process <-suppressWarnings(read_excel(
  "MER_Masterfile_MSUC.xlsx",sheet="New format",guess_max = 1048576))
#characterization
setwd("C:/Users/deboodt.j/Procter and Gamble/Theia_BIC - Documents/Master files/2_Characterization/PSD Accusizer")
PSD_Accusizer <- suppressWarnings(read_xlsb(
  "Masterfile PSD AccuSizer.xlsb",sheet="Summary",skip=1,guess_max = 1048576))
setwd("C:/Users/deboodt.j/Procter and Gamble/Theia_BIC - Documents/Master files/2_Characterization/PSD Occhio")
PSD_Occhio <- suppressWarnings(read_excel(
  "Masterfile_Occhio.xlsm",sheet="New format",guess_max = 1048576))
setwd("C:/Users/deboodt.j/Procter and Gamble/Theia_BIC - Documents/Master files/2_Characterization/Mechanical properties")
FS <- suppressWarnings(read_excel(
  "Master file mechanical properties iNano.xlsm",sheet="Summary",guess_max = 1048576))
#performance
setwd("C:/Users/deboodt.j/Procter and Gamble/Theia_BIC - Documents/Master files/3_Performance")
fabrics <-suppressWarnings(read_excel(
  "Master file Fabrics WM Results.xlsx",sheet="Summary",guess_max = 1048576))
leakage <- read_excel("Master file Leakage Results.xlsm", sheet= "Summary")
QFO <- suppressWarnings(read_excel(
  "Master file Non encapsulated oil.xlsx",sheet="Summary",skip=1,guess_max = 1048576))

#Data cleaning/filtering
PSD_Occhio_number <- filter(PSD_Occhio,`Diameter based on area or perimeter`=="Area-based" & PSD_Occhio$`Number/Area/Volume`=="number")
PSD_Occhio_volume <- filter(PSD_Occhio,`Diameter based on area or perimeter`=="Area-based" & PSD_Occhio$`Number/Area/Volume`=="volume")

#data aggregation
fabrics_aggregated_RelHS <- fabrics %>% group_by(`Touch point`,`Slurry ID`,`Perfume level`, Matrix) %>% summarize(RelHS=mean(`Reference comparison`),RelHSSD=mean(`Reference comparison`))
fabrics_aggregated_AbsHS <- fabrics %>% group_by(`Touch point`,`Slurry ID`,`Perfume level`, `PRM name`, Matrix) %>% summarize(RelHS=mean(`PRM value`),RelHSSD=mean(`PRM value`))
PSD_Occhio_volume_aggregated <- PSD_Occhio_volume %>% group_by(`Slurry ID`,`Number/Area/Volume`) %>% summarize(D50=mean(`P50`))
PSD_Accusizer_Aggregated <- PSD_Accusizer %>% group_by(`Slurry.ID`,`Column2`) %>% summarize(D50=mean(`d50`))

intercept <- function(x, y) {
  ux <- mean(x)
  uy <- mean(y)
  slope <- sum((x-ux) *(y-uy)) / sum((x-ux) ^ 2)
  intercept <- uy-slope*ux
  return(intercept)
}

FS_Aggregated_LM <- FS %>%
  group_by(`Sample ID`,`Instrument`, add=FALSE) %>%
  summarize(RF_slope=slope(`Sphere Diameter`,`Ruptureforce`),RF_intercept=intercept(`Sphere Diameter`,`Ruptureforce`),RD_slope=slope(`Sphere Diameter`,`Rupture displacement`),RD_intercept=intercept(`Sphere Diameter`,`Rupture displacement`))

#Join FS and PSD tables, calculate RF @ Dv50
FS_Aggregated_LM_Dv50 <- left_join(FS_Aggregated_LM,PSD_Occhio_volume_aggregated,by = c("Sample ID"="Slurry ID"))
FS_Aggregated_LM_Dv50$RF_Dv50= FS_Aggregated_LM_Dv50$RF_slope*FS_Aggregated_LM_Dv50$D50+FS_Aggregated_LM_Dv50$RF_intercept

#Join RF @ Dv50 with fabrics HS, calculate correlation RF with WFO/DFO
fabrics_FS_correlation <- left_join(fabrics_aggregated_RelHS, FS_Aggregated_LM_Dv50, by = c("Slurry ID"="Sample ID"))
fabrics_FS_correlation <- fabrics_FS_correlation[!is.na(fabrics_FS_correlation$RelHS),]
fabrics_FS_correlation <- fabrics_FS_correlation[!is.na(fabrics_FS_correlation$RF_Dv50),]
fabrics_FS_correlation_WFO <- filter(fabrics_FS_correlation,`Touch point` == "WFO")
fabrics_FS_correlation_DFO <- filter(fabrics_FS_correlation,`Touch point` == "DFO")

slope <- function(x, y) {
  ux <- mean(x)
  uy <- mean(y)
  slope <- sum((x-ux) *(y-uy)) / sum((x-ux) ^ 2)
  return(slope)
}

WFO <- ggplot(fabrics_FS_correlation_WFO) +
  geom_point(aes(x=`RF_Dv50`,y=`RelHS`, color=`Slurry ID`)) +
  ggtitle("Relative wet fabrics headspace vs rupture force at Dv50")
DFO <- ggplot(fabrics_FS_correlation_DFO) +
  geom_point(aes(x=`RF_Dv50`,y=`RelHS`, color=`Slurry ID`))+
  ggtitle("Relative dry fabrics headspace vs rupture force at Dv50")
ggarrange(WFO,DFO, ncol = 2,nrow=1,common.legend=TRUE)

#Clear environment
#rm(list = ls(all.names = TRUE))
