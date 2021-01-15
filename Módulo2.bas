Attribute VB_Name = "Módulo2"
Sub iva()


Dim presio, cont, ox, numeroOrden As Integer
numeroOrden = 9
presio = 6
ox = 3
Do While Cells(ox, numeroOrden).Value <> ""
  
cont = cont + 1
Cells(ox, presio).Value = Cells(ox, presio).Value - (Cells(ox, presio).Value * 0.16)

  ox = ox + 1

Loop



End Sub
