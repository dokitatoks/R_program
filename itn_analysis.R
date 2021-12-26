library(readxl)
library(purrr)
library(tidyverse)
library(data.table)
library(meta)
library(scales)
library(gridExtra)
library(patchwork)
library(ggthemr)


ggthemr("fresh")


path = "./Data/main_data.xlsx"
sheets <- readxl::excel_sheets(path)
all <- purrr::map_df(sheets, ~dplyr::mutate(readxl::read_excel(path, sheet = .x), 
                                     outcome= .x)) %>%
  mutate(perc = (r/ n)*100,
         studyID = paste0(author, ", ", country))

path = "./Data/int_data.xlsx"
sheets <- readxl::excel_sheets(path)
int <- purrr::map_df(sheets, ~dplyr::mutate(readxl::read_excel(path, sheet = .x), 
                                            sheetname = .x))

# Use_General

pdf(file = "./Figures/Figure1_Use_General.pdf", width = 9, height = 13)
all %>% filter(outcome == "Use_General") %>%
  metaprop(r, n, author, data=., byvar = net_type,
           comb.fixed = F, sm  = "PFT", pscale = 100) %>%
  forest(., hetstat = FALSE, bylab = "", print.byvar = T, pooled.totals = F,
       leftcols = c("studlab","country", "event", "n"),
       sortvar=TE, xlim = c(0, 100), addrow.subgroups = T,
       leftlabs = c("Study", "Country",  "Use", "Total"),
       rightlabs = c("Prevalence",  "(95% CI)"),
       col.inside = "black", col.square="grey", col.square.lines="black", col.by="black", 
       overall = F, fontsize=12, just.addcols = "left")
dev.off()



  


# "Own_General"  
pdf(file = "./Figures/Figure2_Own_General.pdf", width = 9, height = 14)
all %>% filter(outcome == "Own_General") %>%
  metaprop(r, n, author, data=., byvar = net_type,
           comb.fixed = F, sm  = "PFT", pscale = 100) %>%
  forest(., hetstat = FALSE, bylab = "", print.byvar = T, pooled.totals = F,
         leftcols = c("studlab","country", "event", "n"),
         sortvar=TE, xlim = c(0, 100), addrow.subgroups = T,
         leftlabs = c("Study", "Country",  "Own", "Total"),
         rightlabs = c("Prevalence",  "(95% CI)"),
         col.inside = "black", col.square="grey", col.square.lines="black", col.by="black", 
         overall = F, fontsize=12, just.addcols = "left")
dev.off()



# "Use_children" 
pdf(file = "./Figures/Figure3_Use_children.pdf", width = 9, height = 11)
all %>% filter(outcome == "Use_children") %>%
  metaprop(r, n, author, data=., byvar = net_type,
           comb.fixed = F, sm  = "PFT", pscale = 100) %>%
  forest(., hetstat = FALSE, bylab = "", print.byvar = T, pooled.totals = F,
         leftcols = c("studlab","country", "event", "n"),
         sortvar=TE, xlim = c(0, 100), addrow.subgroups = T,
         leftlabs = c("Study", "Country",  "Use", "Total"),
         rightlabs = c("Prevalence",  "(95% CI)"),
         col.inside = "black", col.square="grey", col.square.lines="black", col.by="black", 
         overall = F, fontsize=12, just.addcols = "left")
dev.off()


# "Own_children"
pdf(file = "./Figures/Figure4_Own_children.pdf", width = 9, height = 9)
all %>% filter(outcome == "Own_children") %>%
  metaprop(r, n, author, data=., byvar = net_type,
           comb.fixed = F, sm  = "PFT", pscale = 100) %>%
  forest(., hetstat = FALSE, bylab = "", print.byvar = T, pooled.totals = F,
         leftcols = c("studlab","country", "event", "n"),
         sortvar=TE, xlim = c(0, 100), addrow.subgroups = T,
         leftlabs = c("Study", "Country",  "Own", "Total"),
         rightlabs = c("Prevalence",  "(95% CI)"),
         col.inside = "black", col.square="grey", col.square.lines="black", col.by="black", 
         overall = F, fontsize=12, just.addcols = "left")
dev.off()

# "Use_Women" 
pdf(file = "./Figures/Figure5_Use_Women.pdf", width = 9, height = 8)
all %>% filter(outcome == "Use_Women") %>%
  metaprop(r, n, author, data=., byvar = net_type,
           comb.fixed = F, sm  = "PFT", pscale = 100) %>%
  forest(., hetstat = FALSE, bylab = "", print.byvar = T, pooled.totals = F,
         leftcols = c("studlab","country", "event", "n"),
         sortvar=TE, xlim = c(0, 100), addrow.subgroups = T,
         leftlabs = c("Study", "Country",  "Use", "Total"),
         rightlabs = c("Prevalence",  "(95% CI)"),
         col.inside = "black", col.square="grey", col.square.lines="black", col.by="black", 
         overall = F, fontsize=12, just.addcols = "left")
dev.off()

# "Own_women"   
pdf(file = "./Figures/Figure6_Own_women.pdf", width = 9, height = 22)
all %>% filter(outcome == "Own_women") %>%
  metaprop(r, n, author, data=., byvar = net_type,
           comb.fixed = F, sm  = "PFT", pscale = 100) %>%
  forest(., hetstat = FALSE, bylab = "", print.byvar = T, pooled.totals = F,
         leftcols = c("studlab","country", "event", "n"),
         sortvar=TE, xlim = c(0, 100), addrow.subgroups = T,
         leftlabs = c("Study", "Country",  "Own", "Total"),
         rightlabs = c("Prevalence",  "(95% CI)"),
         col.inside = "black", col.square="grey", col.square.lines="black", col.by="black", 
         overall = F, fontsize=12, just.addcols = "left")
dev.off()

# Interventions
# https://stackoverflow.com/questions/10589693/convert-data-from-long-format-to-wide-format-with-multiple-measure-columns
# own
pdf(file = "./Figures/Figure7_Own_Intervention.pdf", width = 11, height = 7)

read_excel("./Data/int_data.xlsx", 1) %>%
  pivot_wider(names_from = c(interven),
              values_from = c(r, n)) %>%
  rename(r2 = `r_Post-intervention`,
         n2 = `n_Post-intervention`,
         r1 = `r_Pre-intervention`,
         n1 = `n_Pre-intervention`) %>%
  mutate(
    per1 = (r1/n1)*100,
    per2 = (r2/n2)*100,
    per1 = sprintf("%0.1f", round(per1, digits =1)),
    per2 = sprintf("%0.1f", round(per2, digits =1))) %>% 
  metabin(r2, n2, r1, n1, author, data = ., sm = "OR", method="I", comb.fixed = F, byvar = net_type) %>%
  forest(., overall.hetstat = TRUE,bylab = "", print.byvar = T, 
       leftcols = c("studlab", "event.e", "n.e", "per2","event.c", "n.c", "per1"),
       leftlabs=c("Study", "Own", "Total", "%Own", "Own", "Total", "%Own"),
       lab.e = "                      Post-Intervention   ", 
       lab.c = "                      Pre-Intervention",
       sortvar=-TE, comb.fixed = F,
       col.inside = "black", col.square="grey", col.square.lines="black", 
       overall = T, just.addcols = "left")


dev.off()

# use
pdf(file = "./Figures/Figure8_Use_Intervention.pdf", width = 11, height = 7)

read_excel("./Data/int_data.xlsx", 2) %>%
  pivot_wider(names_from = c(interven),
              values_from = c(r, n)) %>%
  rename(r2 = `r_Post-intervention`,
         n2 = `n_Post-intervention`,
         r1 = `r_Pre-intervention`,
         n1 = `n_Pre-intervention`) %>%
  mutate(
    per1 = (r1/n1)*100,
    per2 = (r2/n2)*100,
    per1 = sprintf("%0.1f", round(per1, digits =1)),
    per2 = sprintf("%0.1f", round(per2, digits =1))) %>% 
  metabin(r2, n2, r1, n1, author, data = ., sm = "OR", method="I", comb.fixed = F, byvar = net_type) %>%
  forest(., overall.hetstat = TRUE,bylab = "", print.byvar = T, 
         leftcols = c("studlab", "event.e", "n.e", "per2","event.c", "n.c", "per1"),
         leftlabs=c("Study", "Use", "Total", "%Use", "Use", "Total", "%Use"),
         lab.e = "                      Post-Intervention   ", 
         lab.c = "                      Pre-Intervention",
         sortvar=-TE, comb.fixed = F,
         col.inside = "black", col.square="grey", col.square.lines="black", 
         overall = T, just.addcols = "left", just	= "center")


dev.off()


