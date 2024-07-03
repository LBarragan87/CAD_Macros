Sub BlockList()
    'Require [Functions/GetInsPointCoordinates,Utilities/SelectElements]
    Dim ObjCoord    As Variant
    Set Elements = SelectElements("Seleccione Bloques", "ListaBloques")
    
    Counter = 1
    For Each Element In Elements
        ElemenType = Element.ObjectName
        If ElemenType = "AcDbBlockReference" Then
            ElemenType = Element.ObjectName
            ObjName = Element.EffectiveName
            ObjLayer = Element.Layer
            ObjCoord = GetInsPointCoordenates(Element)
            Debug.Print Counter & "\" & ElemenType & "\" & ObjLayer & "\" & ObjName & "\" & ObjCoord(0) & "\" & ObjCoord(1)
            Counter = Counter + 1
        End If
    Next
    
End Sub
