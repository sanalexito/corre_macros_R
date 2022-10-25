#Se requiere la ruta de la macro que pinta que debe estar habilitada para ejecutarse así chido
# como pueden tenerse varios libros en un mismo directorio se pide por separado el directorio
# y el nombre de los est y cv entre comillas.

macro_pinta <- function(ruta_m, lib_est, lib_cv){

#dir = "D:/Varios/correr_macro_R/correr_macro_R/libros/" O sea donde están tus libros
#setwd(dir)
#################################LEER ARCHIVOS EN EL DIRECTORIO
#a = list.files(dir)
#############################################################
xlApp <- COMCreate("Excel.Application")
#####################################
#Corresponde al libro de estimaciones#
xlWbk <- xlApp$Workbooks()$Open(lib_est)
xlApp[['Visible']] <- FALSE
#####################################
#Corresponde al libro de cv#
xlWbk <- xlApp$Workbooks()$Open(lib_cv)
xlApp[['Visible']] <- FALSE
###################################
xlWbk <- xlApp$Workbooks()$Open(ruta_m)
xlApp[['Visible']] <- FALSE
##CORRER LA MACRO################
xlApp$Run("pinta")
# Close the workbook and quit the app:
xlWbk$Close(FALSE)
xlApp$Quit()
#######################################
}
