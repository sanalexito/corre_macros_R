#Pone el superíndice en la primera columna.

library(RDCOMClient)
#############3
dir = "~/libros/"
setwd(dir)
#################################LEER ARCHIVOS EN EL DIRECTORIO
a = list.files(dir)
#############################################################
xlApp <- COMCreate("Excel.Application")
#####################################
xlWbk <- xlApp$Workbooks()$Open(paste0(dir,a[grep("MARCADOR pa Buscar en a",a)]))#Corresponde a un libro en la lista a##
xlApp[['Visible']] <- FALSE
###################################
xlWbk <- xlApp$Workbooks()$Open("~/superindice_asp.xlsm")
xlApp[['Visible']] <- FALSE
##CORRER LA MACRO################
xlApp$Run("superindice")
# Close the workbook and quit the app:
xlWbk$Close(FALSE)
xlApp$Quit()
#######################################
#######################################
