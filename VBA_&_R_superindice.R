library(RDCOMClient)
#############3
dir = "D:/Varios/correr_macro_R/correr_macro_R/libros/"
setwd(dir)
#################################LEER ARCHIVOS EN EL DIRECTORIO
a = list.files(dir)
#############################################################
xlApp <- COMCreate("Excel.Application")
#####################################
xlWbk <- xlApp$Workbooks()$Open(paste0(dir,a[grep("COPIA",a)]))#Corresponde al libro de estimaciones##
xlApp[['Visible']] <- FALSE
###################################
xlWbk <- xlApp$Workbooks()$Open("D:/Varios/correr_macro_R/correr_macro_R/macros_R/superindice_asp.xlsm")
xlApp[['Visible']] <- FALSE
##CORRER LA MACRO################
xlApp$Run("superindice")
# Close the workbook and quit the app:
xlWbk$Close(FALSE)
xlApp$Quit()
#######################################
#######################################