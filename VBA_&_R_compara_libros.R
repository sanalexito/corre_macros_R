library(RDCOMClient)
source("Ruta/donde/tienes la macro/macro_compara.R")
#Directorio donde está mi macro que compara: "D:/Varios/correr_macro_R/correr_macro_R/macros_R/"
# Se ponen primero las rutas de los archivos que voy a comparar

# Estimaciones
dir_libros <- "Ruta/donde/guardas los archivos/para_comparar/libros/"
est1 <- "libro_COPIA.xlsx"
est2 <- "libro_original.xlsx"
macro_compara(paste0(dir_libros, est1), paste0(dir_libros, est2))

