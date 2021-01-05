# Convert pdf files from plannign portal into png image files for georeferencing in QGIS

library(pdftools)

existing <- file.path("00467504.pdf")
pdf_convert(existing)

proposed <- file.path("00467505.pdf")
pdf_convert(proposed)

habmap <- file.path("00467527.pdf")
pdf_convert(habmap, pages = 39)
