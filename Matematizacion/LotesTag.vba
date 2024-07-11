Sub AutoMatematizacion()
    'Coloca cotas en Lineas y Arcos perimetrales orientado al centroide de la Region
    'Coloca numero de lote segun orden de seleccion
    'coloca etiqueta de area correspondiente a lote
    'alinea textos segun linea/arco mas largo
    
    SelectedElements = ThisDrawing.ModelSpace.Count - 1
    
    consecutivo = 1
    For elemento = 0 To SelectedElements
        x = ThisDrawing.ModelSpace.Item(elemento).Explode
        totalSubelementos = UBound(x)
        subConsecutivo = 1
        maximoSubelemento = 0
        For subElemento = 0 To totalSubelementos
            nombreSubelemento = x(subElemento).ObjectName
            If nombreSubelemento = "AcDbLine" Then
                thisLength = x(subElemento).Length
            ElseIf nombreSubelemento = "AcDbArc" Then
                thisLength = x(subElemento).ArcLength
            End If
            x(subElemento).Delete
            
            If thisLength > maximoSubelemento Then
            maximoSubelemento = thisLength
            Else
            maximoSubelemento = maximoSubelemento
            End If
            
            Debug.Print consecutivo & "\" & subConsecutivo & "\" & thisLength
            subConsecutivo = subConsecutivo + 1
        Next subElemento
        MaxElemento = maximoSubelemento
        Debug.Print consecutivo & "\" & MaxElemento & "<<---max"
        consecutivo = consecutivo + 1
    Next elemento
    
End Sub
