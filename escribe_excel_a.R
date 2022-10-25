#-------------------------------------------------------------------------------

Escribe_Excel_a <- function(dir0, num_hoja, tab){
  library(tatoo)
  library(openxlsx)
  A = openxlsx::read.xlsx(dir0, sheet = num_hoja, colNames=F, skipEmptyRows=F)
  ren = which(A$X1%in%"Estados Unidos Mexicanos") 
  wb1 = openxlsx::loadWorkbook(dir0)
  #openxlsx::writeData(wb, sheet = num_hoja, tabs[[1]], startRow = ren, colNames = F)
  
#--------------------------------------------------------------------------
  
  #wb1 <- openxlsx::createWorkbook() # create workbook
  
  #openxlsx::addWorksheet(wb1, sheetName = "data") # add sheet
  
  openxlsx::writeData(wb1, sheet=num_hoja, x=tab, xy=c(1, ren), colNames = F) # write data on workbook
  
  for(i in grep("\\_\\[([A-z0-9\\s]*)\\]", wb1$sharedStrings)){
    # if empty string in superscript notation, then just remove the superscript notation
    if(grepl("\\_\\[\\]", wb1$sharedStrings[[i]])){
      wb1$sharedStrings[[i]] <- gsub("\\_\\[\\]", "", wb1$sharedStrings[[i]])
      next # skip to next iteration
    }
    
    # insert additioanl formating in shared string
    wb1$sharedStrings[[i]] <- gsub("<si>", "<si><r>", gsub("</si>", "</r></si>", wb1$sharedStrings[[i]]))
    
    # find the "_[...]" pattern, remove brackets and udnerline and enclose the text with superscript format
    wb1$sharedStrings[[i]] <- gsub("\\_\\[([A-z0-9\\s]*)\\]", "</t></r><r><rPr><vertAlign val=\"superscript\"/></rPr><t xml:space=\"preserve\">\\1</t></r><r><t xml:space=\"preserve\">", wb1$sharedStrings[[i]])
  }
  
    openxlsx::saveWorkbook(wb1, dir0, overwrite = T)

}
