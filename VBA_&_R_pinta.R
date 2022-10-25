library(RDCOMClient)
source('D:/ENADIS/Calculadores/macro_pinta.R')
#-------------------------------------------------------------------------------
ruta_m <- "D:/Varios/correr_macro_R/correr_macro_R/macros_R/macro_pinta_asp.xlsm"
ruta_est <- "D:/Varios/correr_macro_R/correr_macro_R/libros/III_experiencia_pagos_tramites_COPIA.xlsx"
ruta_cvs <- "D:/Varios/correr_macro_R/correr_macro_R/libros/III_experiencia_pagos_tramites_solicitudes_encig2015_cv.xlsx"
#setwd(dir)
#-------------------------------------------------------------------------------
macro_pinta(ruta_m = ruta_m, lib_est = ruta_est, lib_cv = ruta_cv )
