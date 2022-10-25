# Se ponen primero las rutas de los archivos que voy a comparar y la ruta de la macro que se va a actualizar
# La función actualiza los datos con los que se correrá el macro.
macro_compara <- function(ruta1, ruta2, renglones = 1000, columnas=50){
dir0 <- "D:/Varios/correr_macro_R/correr_macro_R/macros_R/"
#ruta1 <- "D:/Varios/correr_macro_R/correr_macro_R/libros/III_experiencia_pagos_tramites_COPIA.xlsx"
#ruta2 <- "D:/Varios/correr_macro_R/correr_macro_R/libros/III_experiencia_pagos_tramites_solicitudes_encig2015_est.xlsx"
#renglones <- 800
#columnas <- 50
tabla <- as.data.frame(ruta1)
tabla[2,1] <- ruta2
tabla[3,1] <- renglones
tabla[4,1] <- columnas
informacion <- c("Ruta1:", "Ruta2:", "Renglones:", "Columnas:")
tabla <- cbind(informacion, tabla)

wb <- openxlsx::loadWorkbook(paste0(dir0, "compara_libros_asp.xlsm"))
#openxlsx::addWorksheet(wb, sheetName = "Hoja_compara")
openxlsx::writeData(wb, sheet = 1, x = tabla, startCol = 1, colNames = F)
#openxlsx::removeWorksheet(wb, sheet = 1)
openxlsx::saveWorkbook(wb, paste0(dir0, "compara_libros_asp.xlsm"), overwrite = T )

xlApp <- COMCreate("Excel.Application")
xlWbk <- xlApp$Workbooks()$Open(paste0(dir0,"compara_libros_asp.xlsm"))
xlApp[['Visible']] <- FALSE
##CORRER LA MACRO###########################################
xlApp$Run("Compara_libros")
# Close the workbook and quit the app:
xlWbk$Close(FALSE)
xlApp$Quit()
}
