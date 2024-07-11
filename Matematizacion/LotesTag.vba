Sub AutoMatematizacion()
    'Coloca cotas en Lineas y Arcos perimetrales orientado al centroide de la Region
    'Coloca numero de lote segun orden de seleccion
    'coloca etiqueta de area correspondiente a lote
    'alinea textos segun linea/arco mas largo
    
    SelectedElements = ThisDrawing.ModelSpace.Count - 1
    
    Consecutive = 1
    For elemento = 0 To SelectedElements
        ExplotedObject = ThisDrawing.ModelSpace.Item(elemento).Explode
        totalSubelementos = UBound(ExplotedObject)
        SubConsecutive = 1
        MaxLength = 0
        For subElemento = 0 To totalSubelementos
            nombreSubelemento = ExplotedObject(subElemento).ObjectName
            If nombreSubelemento = "AcDbLine" Then
                thisLength = ExplotedObject(subElemento).Length
            ElseIf nombreSubelemento = "AcDbArc" Then
                thisLength = ExplotedObject(subElemento).ArcLength
            End If
            DeleteObject = ExplotedObject(subElemento).Delete
            
            If thisLength > MaxLength Then
            MaxLength = thisLength
            Else
            MaxLength = MaxLength
            End If
            
            Debug.Print Consecutive & "\" & SubConsecutive & "\" & thisLength
            SubConsecutive = SubConsecutive + 1
        Next subElemento

        Debug.Print Consecutive & "\" & MaxLength & "<<---max"
        Consecutive = Consecutive + 1
    Next elemento
    
End Sub
