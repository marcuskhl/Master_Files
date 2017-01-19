if(is.na(match(c("devtools"),installed.packages()[,"Package"]))) install.packages(new.packages) else library(devtools)
suppressMessages(devtools::install_github("marcuskhl/BasicSettings"));suppressMessages(library(BasicSettings))

x <- 60 * 60 * 24 * 1.5 
Sys.sleep(x) # delay (in seconds)

today <- Sys.Date()+1
today <- gsub("-", "", as.character(today), fixed = T)


#~~~TRAX Start~~~#
TRAX <- read.xlsx("M:/Television and Broadband/INTELLIGENCE/TRAX/TRAX_TV_Media_Intelligence.xlsx", sheet = 1, colNames = T )

TRAX <- TRAX[is.na(TRAX$Pub), ]

TRAX <- TRAX[,-c(2,3,7,17)]

# TRAX <- TRAX[complete.cases(TRAX[15:length(names(TRAX))]),] # there are cells that is intentionally blank in a row
TRAX_error <- TRAX[rowSums(is.na(TRAX[15:length(names(TRAX))])) == 1 | rowSums(is.na(TRAX[15:length(names(TRAX))])) == 2,]# if there are only 1 or 2 NAs, it is probably error

TRAX <- TRAX[rowSums(is.na(TRAX[15:length(names(TRAX))])) != 1, ]
TRAX <- TRAX[rowSums(is.na(TRAX[15:length(names(TRAX))])) != 2, ] 
# rounding to 4 digits
is.num <- sapply(TRAX, is.numeric)
TRAX[is.num] <- lapply(TRAX[is.num], round, 4)

dir.create(file.path("M:/Television and Broadband/INTELLIGENCE/TRAX/Valued/", paste0(today)), showWarnings = FALSE)
# fwrite(TRAX, paste0("M:/Television and Broadband/INTELLIGENCE/TRAX/Valued/", today, "TV_Media_TRAX_", today, ".csv"), row.names = F)
# fwrite(TRAX_error, paste0("M:/Television and Broadband/INTELLIGENCE/TRAX/Valued/", today, "Error.TV_Media_TRAX_", today, ".csv"), row.names = F)
save.xlsx(paste0("M:/Television and Broadband/INTELLIGENCE/TRAX/Valued/", today, "TV_Media_TRAX_", today, ".xlsx"), TRAX, TRAX_error)
#~~~TRAX End~~~#




#~~~Master Files Start~~~#
name_list <- c("TV_CORE_NEW_MEA", "TV_CORE_NEW_US", "TV_CORE_NEW_AP", "TV_CORE_NEW_EE1", "TV_CORE_NEW_EE2", "TV_CORE_NEW_WE1", "TV_CORE_NEW_WE2", "TV_CORE_NEW_3DIM")



dir.create(file.path("M:/Television and Broadband/INTELLIGENCE/master files/NEW/Valued/", paste0(today)), showWarnings = FALSE)



for (i in 1 : length(name_list)){
  p1 <- proc.time()
  wb <- loadWorkbook(paste0("M:/Television and Broadband/INTELLIGENCE/master files/NEW/", name_list[i], ".xlsx"))
  for (j in 1:length(excel_sheets(paste0("M:/Television and Broadband/INTELLIGENCE/master files/NEW/", name_list[i], ".xlsx")))) {
    writeComment(wb, j, col = 1, row = 2, comment = c1)
  }
  saveWorkbook(wb, paste0("M:/Television and Broadband/INTELLIGENCE/master files/NEW/Valued/", today, name_list[i], ".xlsx"), overwrite = T)
  print(past0(name_list[i], ".xlsx took ", proc.time()-p1, " seconds."))
  rm(wb)
}

#~~~Experimental~~~#
p1 <- proc.time()
wb <- loadWorkbook("M:/Television and Broadband/INTELLIGENCE/master files/NEW/TV_CORE_NEW_MEA.xlsx")
c1 <- createComment(comment = "", visible = F)
for (j in 1:length(excel_sheets("M:/Television and Broadband/INTELLIGENCE/master files/NEW/TV_CORE_NEW_MEA.xlsx"))) {
  writeComment(wb, j, col = 1, row = 2, comment = c1)
}

saveWorkbook(wb, "C:/testing_MEA.xlsx", overwrite = T)
proc.time()-p1
#~~~Experimental~~~#
#~~~Master Files End~~~#






