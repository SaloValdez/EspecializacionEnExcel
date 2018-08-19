Sub mostrarResultadoFuncion()

    numero1 = Worksheets("Hoja1").Range("A10").Value
    numero2 = Worksheets("Hoja1").Range("B10").Value
    
    'Muestra el resultado  de la multiplicacion
    MsgBox mulNo(numero1, numero2)
    Worksheets("Hoja1").Range("C10") = mulNo(numero1, numero2) 'ejecuta la funcion y muestra en la celda C10
    
    'Muestra el resultado  de la suma
    MsgBox suma(numero1, numero2)
    Worksheets("Hoja1").Range("D10") = suma(numero1, numero2)  'ejecuta la funcion y muestra en la celda D10
End Sub

Function mulNo(a, b)
    mulNo = a + b
End Function

Function suma(a, b)
    suma = a * b
End Function
