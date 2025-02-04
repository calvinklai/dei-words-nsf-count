# Author: Calvin K. Lai
# Date: 02/04/2025
# Function: Counts the number of words times words were usedi n the word document. There are some slight discrepancies between how my script counts compared to Microsoft Word, but it should be pretty close. 

# Installing and loading relevant package
list.of.packages <- c("officer", "stringr", "dplyr")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)
library(officer)
library(stringr)
library(dplyr)

# Load the document
# NOTE: Change setwd() and doc_path to your directory and file name, respectively.
setwd("YOURDIRECTORY")
doc_path <- "YOURFILENAME.docx"
doc <- read_docx(doc_path)

# Extract text from paragraphs
text <- paste(unlist(docx_summary(doc)$text), collapse = " ")
text <- tolower(text)  # Convert to lowercase for case-insensitive matching

# List of words/phrases to count
terms <- c("activism", "activists", "advocacy", "advocate", "advocates", "barrier", "barriers",
           "biased", "biased toward", "biases", "biases towards", "bipoc", "black and latinx",
           "community diversity", "community equity", "cultural differences", "cultural heritage",
           "culturally responsive", "disabilities", "disability", "discriminated", "discrimination",
           "discriminatory", "diverse backgrounds", "diverse communities", "diverse community",
           "diverse group", "diverse groups", "diversified", "diversify", "diversifying",
           "diversity and inclusion", "diversity equity", "enhance the diversity", "enhancing diversity",
           "equal opportunity", "equality", "equitable", "equity", "ethnicity", "excluded", "female",
           "females", "fostering inclusivity", "gender", "gender diversity", "genders", "hate speech",
           "racial inequality", "racial justice", "racially", "racism", "sense of belonging",
           "sexual preferences", "social justice", "socio cultural", "socio economic", "sociocultural",
           "socioeconomic", "status", "stereotypes", "systemic", "trauma", "under appreciated",
           "under represented", "under served", "underrepresentation", "underrepresented",
           "underserved", "undervalued", "victim", "women", "women and underrepresented",
           "antiracist", "hispanic minority", "historically", "implicit bias", "implicit biases",
           "inclusion", "inclusive", "inclusiveness", "inclusivity", "increase diversity",
           "increase the diversity", "indigenous community", "inequalities", "inequality",
           "inequitable", "inequities", "institutional", "lgbt", "marginalize", "marginalized",
           "minorities", "minority", "multicultural", "polarization", "political", "prejudice",
           "privileges", "promoting diversity", "race and ethnicity", "racial", "racial diversity")

# Ensure exact word matching by using word boundaries
word_counts <- sapply(terms, function(term) str_count(text, paste0("\\b", term, "\\b")))

# Convert to data frame
word_count_df <- data.frame(Word_Phrase = terms, Count = word_counts)

# Calculate total count
total_count <- sum(word_counts)

# Print table
print(word_count_df)
cat("\nTotal Count:", total_count)
