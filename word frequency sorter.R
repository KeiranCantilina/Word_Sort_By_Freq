## Script for looking at word frequencies in a Word docx document
## I wrote this to help Ren find words she uses too frequently in her writing

library(officer)
library(stringi)
library(data.table)

## Set document path
path <- "C://Users//Keiran//Documents//NaNoWriMo 2018_Draft4.docx"

## Import document contents
imported_doc <- system.file(package = "officer", path)
doc <- read_docx(path)
content <- docx_summary(doc)

## Turn text paragraphs into a single giant character string
text_list<- c(content$text[1:length(content$text)])
merged_text <- paste(text_list[1:length(text_list)], sep="", collapse=" ")

## Coerce encoding to ASCII to avoid problems with grep and quotation marks
ascii_text<- stri_enc_toascii(merged_text)

## Remove puncuation
scrubbed_text<- gsub('\032', " ", ascii_text, fixed=TRUE)
scrubbed_text<- gsub(".", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub("!", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub("?", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub(",", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub("  ", " ", scrubbed_text, fixed=TRUE)

## Words to omit from list
scrubbed_text<- gsub(" or ", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub(" the ", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub(" and ", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub(" but ", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub(" for ", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub(" so ", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub(" if ", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub(" then ", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub(" at ", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub(" t ", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub(" s ", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub(" m ", " ", scrubbed_text, fixed=TRUE)
scrubbed_text<- gsub(" re ", " ", scrubbed_text, fixed=TRUE)

## Finally scrub extra spaces
scrubbed_text<- gsub("  ", " ", scrubbed_text, fixed=TRUE)

## Split giant string into a vector of words
words <- strsplit(scrubbed_text, " ", fixed=TRUE)
words <- c(words[[1]])

## Organize vector of words by frequency and output as table
table <- as.data.frame(sort(table(words), decreasing=TRUE))
fwrite(table, file="C://Users//Keiran//Desktop//words.csv")
