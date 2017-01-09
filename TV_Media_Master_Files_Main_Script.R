if(is.na(match(c("devtools"),installed.packages()[,"Package"]))) install.packages(new.packages) else library(devtools)
suppressMessages(devtools::install_github("marcuskhl/BasicSettings"));suppressMessages(library(BasicSettings))

TRAX <- read.xlsx("M:/Television and Broadband/INTELLIGENCE/TRAX/TRAX_TV_Media_Intelligence.xlsx", sheet = 1, colNames = T )

TRAX <- TRAX[is.na(TRAX$Pub), ]

TRAX <- TRAX[,-c(2,3,7,17)]

# TRAX <- TRAX[complete.cases(TRAX[15:length(names(TRAX))]),] # there are cells that is intentionally blank in a row
TRAX_error <- TRAX[rowSums(is.na(TRAX[15:length(names(TRAX))])) == 1 | rowSums(is.na(TRAX[15:length(names(TRAX))])) == 2,]# if there are 1 or 2 NAs, it is probably error

TRAX <- TRAX[rowSums(is.na(TRAX[15:length(names(TRAX))])) != 1, ]
TRAX <- TRAX[rowSums(is.na(TRAX[15:length(names(TRAX))])) != 2, ] 
# rounding to 4 digits
is.num <- sapply(TRAX, is.numeric)
TRAX[is.num] <- lapply(TRAX[is.num], round, 4)

today <- Sys.Date()
today <- gsub("-", "", as.character(today), fixed = T)

dir.create(file.path("M:/Television and Broadband/INTELLIGENCE/TRAX/Valued/", paste0(today)), showWarnings = FALSE)
fwrite(TRAX, paste0("M:/Television and Broadband/INTELLIGENCE/TRAX/Valued/", today, "TV_Media_TRAX_", today, ".csv"), row.names = F)
fwrite(TRAX_error, paste0("M:/Television and Broadband/INTELLIGENCE/TRAX/Valued/", today, "Error.TV_Media_TRAX_", today, ".csv"), row.names = F)