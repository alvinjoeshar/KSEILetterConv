library(pdftools)
library(tabulizer)
library(dplyr)
library(writexl)
library(stringr)

extract_tables_from_pdf <- function(pdf_link, output_file) {
  download.file(pdf_link, destfile = "temp.pdf", mode = "wb")
  
   tables <- tabulizer::extract_tables("temp.pdf")
  
  data_frames <- lapply(tables, function(x) {
    df <- as.data.frame(x, stringsAsFactors = FALSE)
    
    if ("V2" %in% colnames(df)) {
      df$V2 <- stringr::str_replace_all(df$V2, "[^0-9.]", "")
    }
    if ("V3" %in% colnames(df)) {
      df$V3 <- stringr::str_replace_all(df$V3, " % p.a", "")
    }
    
    return(df)
  })
  
  file.remove("temp.pdf")
  
  names(data_frames) <- paste0("Table", seq_along(data_frames))
  writexl::write_xlsx(data_frames, path = output_file)
}

pdf_link <- "https://www.ksei.co.id/Announcement/Files/156611_ksei_1859_dir_0723_202307031943.pdf"
output_file <- "datapublikasi.xlsx"
extract_tables_from_pdf(pdf_link, output_file)
