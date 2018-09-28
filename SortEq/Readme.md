Script para Ordenar la Tx IE05/IH08
-----------------------------------------
-----------------------------------------
Explicació del uso del scrip [ReadWriteExcel](https://github.com/KainaK0/KPIs/blob/master/SortEq/ReadWriteExcel.r) por [Johan Callomamani](https://github.com/KainaK0). l siguiente documento cuenta con la instrucciones y programas necesarios para ejecutar el scrip satisfactoriamente. 

###Requisitos
* Instalar el compilador [R](https://cran.r-project.org/mirrors.html)
* Instalar el IDE [RStudio](https://www.rstudio.com/products/rstudio/)
* Data SAP PM Tx IE05/IH08 con las Columnas **Equipo**,**Equipo Superior**y **Tipo de Equipo**. (Sin importar el Orden)

El script cumple con la función de agregar 3 columnas de tal forma que se ordene la data Estas son: **Equipo Padre**, **Sistema**y **Equipo Hijo**.
