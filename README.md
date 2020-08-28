Para poder usar este código primero tienes que instalar el paquete openpyxl

pip install openpyxl

y luego, situar las tablas en la carpeta tablas_excel. 

Es posible que se quiera cambiar las celdas desde las que se extraeran los datos. Para esto, solo cambiarlas en la parte

    NomAp=Hoja1["A5"].value
    DNI=Hoja1["A7"].value
    Dir=Hoja1["A9"].value
    Dist=Hoja1["A11"].value
    Tel=Hoja1["A13"].value
    Ant=Hoja1["A15"].value
    Sint=Hoja1["A31"].value

    DProd=Hoja1["E5"].value
    Prod=Hoja1["E7"].value
    Numser=Hoja1["E9"].value
    FechaCom=Hoja1["E11"].value
    
y también en
        try:
        HojaLN["A"+str(i1+1)]=NomAp
        HojaLN["B"+str(i1+1)]=DNI
        HojaLN["C"+str(i1+1)]=Dir
        HojaLN["D"+str(i1+1)]=Dist
        HojaLN["E"+str(i1+1)]=Tel
        HojaLN["F"+str(i1+1)]=Ant
        HojaLN["G"+str(i1+1)]=Sint
        HojaLN["H" + str(i1+1)] =DProd
        HojaLN["I" + str(i1+1)] =Prod
        HojaLN["J" + str(i1+1)] = Numser
        HojaLN["K" + str(i1+1)] = FechaCom
  
