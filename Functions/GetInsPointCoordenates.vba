Function GetInsPointCoordenates(Element) As Variant

    Dim thisCoordinates(0 To 2) As Variant
    Dim insertionPoint(0 To 2) As Double
    Dim thisInsertionPoint As Variant
    
    thisInsertionPoint = Element.insertionPoint
    thisCoordinates(0) = thisInsertionPoint(0)
    thisCoordinates(1) = thisInsertionPoint(1)
    thisCoordinates(2) = thisInsertionPoint(2)
    
    GetInsPointCoordenates = thisCoordinates
    
End Function
