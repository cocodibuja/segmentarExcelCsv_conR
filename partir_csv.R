################################################################################
## cómo usarlo
## 1) obtenga el path del archivo csv y coloquelo en la variable PATH
## 2) defina la cantidad de FILAS que tendrá cada sheet del excel de salida
## 3) coloque el nombre de salida con terminacion .xlsx NOMBRE_SALIDA_EXCEL
 ################################################################################

PATH = "datos_ejemplo.csv"
FILAS = 3
NOMBRE_SALIDA_EXCEL = "segmentos.xlsx"

#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#install.packages("openxlsx")
library(openxlsx)
posiciones_inicio_fin = function(filas_originales,intervalo = 100){
  #puedo conocer las posiciones donde corto el excel original
  inicio = seq(1,filas_originales, by = intervalo)
  fin = seq(1,filas_originales, by = intervalo)
  fin  = fin -1
  fin = c(fin,filas_originales)
  fin = fin[-c(1)]
  
  return(rbind(inicio,fin))
}

#----actulizamos las variables---
path = PATH

#file.choose()# para abir y buscar el path del archivo

#datos_original = read_excel(path,sheet = 'Hoja1')
datos_original = read.csv(file = path)

#cantidad de filas
filas_originales = dim(datos_original)[1]



# en cuantos pedazos lo partimos según la cantidad de filas maximas por hoja?
max_filas = FILAS # por hoja

if (max_filas <= filas_originales) {
 
  # creo un libro de trabajo en donde pondremos todas las hojas/sheets
  wb <- createWorkbook()
  
  cantidad_sheets = round(filas_originales/max_filas)
  
  
 
  
  #genero las hojas donde voy a escribir los datos
  for (numero in 1:cantidad_sheets) {
    addWorksheet(wb, sheetName = paste0("hoja",numero))
  }
  
  #obtengo las posiciones donde corto el dataframe segun max_filas
  posiciones_corte = posiciones_inicio_fin(filas_originales,max_filas )
  
  for (numero in 1:cantidad_sheets) {
    inicio = posiciones_corte[,numero]["inicio"]
    fin = posiciones_corte[,numero]["fin"]
    df_segmento = datos_original[inicio:fin,]
    writeData(wb, sheet = numero, x = df_segmento, startCol = 1, startRow = 1)
    
  }
  
  # si existe el archivo lo borro y lo genero de nuevo
  if (file.exists(NOMBRE_SALIDA_EXCEL)) {
    file.remove(NOMBRE_SALIDA_EXCEL)
  }
  
  saveWorkbook(wb, file = NOMBRE_SALIDA_EXCEL)
}else{
  print("la cantidad de filas de salida es mayor que la original")
}





