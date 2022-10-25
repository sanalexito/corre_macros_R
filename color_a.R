
pinta <- function(est, cv, hoja, ruta){
  
  A <- openxlsx::read.xlsx(ruta, sheet = hoja, colNames = F, skipEmptyRows = F)
  ren <- which(A$X1%in%'Estados Unidos Mexicanos')
  wb <- openxlsx::loadWorkbook(ruta)

  estilo_a0 <- openxlsx::createStyle(fontName = 'Arial', fontSize = 8, numFmt = "### ### ##0", textDecoration = "Bold",
                                    halign = 'right', valign = 'center', fgFill = '#FF5400')

  estilo_a1 <- openxlsx::createStyle(fontName = 'Arial', fontSize = 8, numFmt = "### ### ##0",
                                    halign = 'right', valign = 'center', fgFill = '#FFEA00')
  estilo_a2 <- openxlsx::createStyle(fontName = 'Arial', fontSize = 8,  numFmt = "### ### ##0",
                                    halign = 'right', valign = 'center', fgFill = '#FF5400')
  estilo_a3 <- openxlsx::createStyle(fontName = 'Arial', fontSize = 8,  numFmt = "### ### ##0",
                                     halign = 'right', valign = 'center')
  
  estilo_r1 <- openxlsx::createStyle(fontName = 'Arial', fontSize = 8, numFmt = "0.0",
                                    halign = 'right', valign = 'center', fgFill = '#FFEA00')
  estilo_r2 <- openxlsx::createStyle(fontName = 'Arial', fontSize = 8,  numFmt = "0.0",
                                     halign = 'right', valign = 'center', fgFill = '#FF5400')
  estilo_r3 <- openxlsx::createStyle(fontName = 'Arial', fontSize = 8,  numFmt = "0.0",
                                     halign = 'right', valign = 'center', fgFill = '#FFFFFF')
  

    for( i in 1:dim(cv)[1] ){
      if(cv[i,2]=="NA" | is.na(cv[i,2]) | cv[i,2]==0){
      openxlsx::addStyle(wb, sheet = hoja, estilo_a3, rows = (ren - 1 +i), cols = 2)
      }else{
    if(any( 15<=as.numeric(cv[i,2]) & as.numeric(cv[i,2])<30, na.rm = T )==T){
      openxlsx::addStyle(wb, sheet = hoja, estilo_a1, rows = (ren - 1 +i), cols = 2)
    }else if(any(30 <= as.numeric(cv[i,2]), na.rm = T)==T ){
      openxlsx::addStyle(wb, sheet = hoja, estilo_a2, rows = (ren - 1 +i), cols = 2)
    }else{

          }
      }
      }
  
  
for(k in 1:dim(est)[2]){
if(sum(is.na(est[,2])) < dim(est)[1])
    for(j in 1:dim(est)[1]){
      if(sum(!cv[j,k]%in%c("NA", NA, 0) )>0){
      if(any( 15 <= as.numeric(cv[j,k]) & as.numeric(cv[j,k]) < 30, na.rm = T)==T){
        openxlsx::addStyle(wb, sheet = hoja, estilo_a1, rows = ren-1 + j, cols = k)
      }else if(any( 30 <= as.numeric(cv[j,k]), na.rm = T)==T){
        openxlsx::addStyle(wb, sheet = hoja, estilo_a2, rows = ren-1+j, cols = k)
      }else{
      } 
    }
  }
  }
  
  openxlsx::writeData(wb,sheet = hoja,est,startRow = ren,colNames = F, na.string = F )
  
  for(i in grep("\\_\\[([A-z0-9\\s]*)\\]", wb$sharedStrings)){
    # if empty string in superscript notation, then just remove the superscript notation
    if(grepl("\\_\\[\\]", wb$sharedStrings[[i]])){
      wb$sharedStrings[[i]] <- gsub("\\_\\[\\]", "", wb$sharedStrings[[i]])
      next # skip to next iteration
    }
    
    # insert additioanl formating in shared string
    wb$sharedStrings[[i]] <- gsub("<si>", "<si><r>", gsub("</si>", "</r></si>", wb$sharedStrings[[i]]))
    
    # find the "_[...]" pattern, remove brackets and udnerline and enclose the text with superscript format
    wb$sharedStrings[[i]] <- gsub("\\_\\[([A-z0-9\\s]*)\\]", "</t></r><r><rPr><vertAlign val=\"superscript\"/></rPr><t xml:space=\"preserve\">\\1</t></r><r><t xml:space=\"preserve\">", wb$sharedStrings[[i]])
    
  }
  
  openxlsx::saveWorkbook(wb,ruta,overwrite = T)

  }


