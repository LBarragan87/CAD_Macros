Sub AutoMatematizacion()
    'Coloca cotas en Lineas y Arcos perimetrales orientado al centroide de la Region
    'Coloca numero de lote segun orden de seleccion
    'coloca etiqueta de area correspondiente a lote
    'alinea textos segun linea/arco mas largo
    
    Set SelectedElements = SelectElements("Seleccione Regiones", "ListaRegiones")
        
    Consecutive = 1
    For Each element In SelectedElements
        ExplotedObject = element.Explode
        SubElementsUBound = UBound(ExplotedObject)
        SubConsecutive = 1
        MaxLength = 0
        For SubElementIndex = 0 To SubElementsUBound
            SubElementType = ExplotedObject(SubElementIndex).ObjectName
            If SubElementType = "AcDbLine" Then
                thisLength = ExplotedObject(SubElementIndex).Length
            ElseIf SubElementType = "AcDbArc" Then
                thisLength = ExplotedObject(SubElementIndex).ArcLength
            End If
            DeleteObject = ExplotedObject(SubElementIndex).Delete
            
            If thisLength > MaxLength Then
            MaxLength = thisLength
            Else
            MaxLength = MaxLength
            End If
            
            Debug.Print Consecutive & "\" & SubConsecutive & "\" & thisLength
            SubConsecutive = SubConsecutive + 1
        Next SubElementIndex

        Debug.Print Consecutive & "\" & MaxLength & "<<---max"
        Consecutive = Consecutive + 1
    Next
    
End Sub
