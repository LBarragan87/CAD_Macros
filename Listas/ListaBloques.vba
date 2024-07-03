Sub ListaBloques2025()
    Dim ObjCoord    As Variant
    Set IN0 = SelectElements("Seleccione Bloques", "ListaBloques")
    
    Counter = 1
    For Each Element In IN0
        ElemenType = Element.ObjectName
        If ElemenType = "AcDbBlockReference" Then
            ObjName = Element.EffectiveName
            ObjLayer = Element.Layer
            ObjCoord = GetInsPointCoordenates(Element)
            Debug.Print Counter & "\" & ObjLayer & "\" & ObjName & "\" & ObjCoord(0) & "\" & ObjCoord(1)
            Counter = Counter + 1
        End If
    Next
    
End Sub
