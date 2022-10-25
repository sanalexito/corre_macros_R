library(RDCOMClient)
#-------------------------------------------------------------------------------
ruta_m <- "Ruta/ de la/macro_pinta_asp.xlsm"
ruta_est <- "Ruta/libro para pintar/estimaciones.xlsx"
ruta_cvs <- "Ruta/libro para pintar/coeficientes_variacion.xlsx"
#setwd(dir)
#-------------------------------------------------------------------------------
macro_pinta(ruta_m = ruta_m, lib_est = ruta_est, lib_cv = ruta_cv )
