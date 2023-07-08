# Load necessary libraries
library(pdftools)
library(tabulizer)
library(dplyr)
library(writexl)
library(stringr)

# Function to extract tables from PDF
extract_tables_from_pdf <- function(pdf_link, output_file) {
  # Download the PDF file
  download.file(pdf_link, destfile = "temp.pdf", mode = "wb")
  
  # Extract tables from the PDF
  tables <- tabulizer::extract_tables("temp.pdf")
  
  # Convert tables to data frames
  data_frames <- lapply(tables, function(x) {
    df <- as.data.frame(x, stringsAsFactors = FALSE)
    
    # Clean up specific columns
    if ("V2" %in% colnames(df)) {
      df$V2 <- stringr::str_replace_all(df$V2, "[^0-9.]", "")
    }
    if ("V3" %in% colnames(df)) {
      df$V3 <- stringr::str_replace_all(df$V3, " % p.a", "")
    }
    
    return(df)
  })
  
  # Clean up temporary PDF file
  file.remove("temp.pdf")
  
  # Write data frames to Excel file
  names(data_frames) <- paste0("Table", seq_along(data_frames))
  writexl::write_xlsx(data_frames, path = output_file)
}

# Use the function
pdf_link <- "https://www.ksei.co.id/Announcement/Files/156611_ksei_1859_dir_0723_202307031943.pdf"
output_file <- "datapublikasi.xlsx"
extract_tables_from_pdf(pdf_link, output_file)
