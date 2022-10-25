library(RDCOMClient)
source("D:/ENADIS/Calculadores/macro_compara.R")
#Directorio donde está mi macro que compara: "D:/Varios/correr_macro_R/correr_macro_R/macros_R/"
# Se ponen primero las rutas de los archivos que voy a comparar
# Como no guarda en automático tienes que hacerlo de uno por uno para que no se pasme

# Estimaciones
dir_libros <- "D:/Varios/correr_macro_R/correr_macro_R/libros/"
est1 <- "III_experiencia_pagos_tramites_COPIA.xlsx"
est2 <- "III_experiencia_pagos_tramites_solicitudes_encig2015_est.xlsx"
macro_compara(paste0(dir_libros, est1), paste0(dir_libros, est2))

# CV 
est1 <- "enve2022_cv_III_denuncia_delito - copia.xlsx"
est2 <- "enve2022_cv_III_denuncia_delito.xlsx"
macro_compara(paste0(dir_libros, est1), paste0(dir_libros, est2))

# INT
int1 <- "enve2022_cv_III_denuncia_delito - copia.xlsx"
int2 <- "enve2022_cv_III_denuncia_delito.xlsx"
macro_compara(paste0(dir_libros, est1), paste0(dir_libros, est2))

# SE
int1 <- "enve2022_cv_III_denuncia_delito - copia.xlsx"
int2 <- "enve2022_cv_III_denuncia_delito.xlsx"
macro_compara(paste0(dir_libros, est1), paste0(dir_libros, est2))

