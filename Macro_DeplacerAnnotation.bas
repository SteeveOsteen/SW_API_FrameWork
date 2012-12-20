Attribute VB_Name = "Macro_DeplacerAnnotation"
'=====================================================================================
'|
'| Transfert cToutes les dimensions dans la vue d'annotation "Objets non affectés"
'|
'=====================================================================================

Public Sub Main()
    Dim Sw                  As SldWorks.SldWorks
    Dim Modele              As ModelDoc2
    Dim GestPiece           As New ExtPiece
    
    Dim vVueAnnotation      As Variant
    Dim VueAnnotation       As AnnotationView
    
    Dim vDimension          As Variant
    Dim Dimension           As DisplayDimension
    Dim colDimensions       As Collection
    Dim tabAnnotations()    As Annotation
    Dim Annotation          As Annotation
    
    Dim i As Integer
    
    Set Sw = Application.SldWorks
    Set Modele = Sw.ActiveDoc
    
    For Each vVueAnnotation In Modele.Extension.AnnotationViews
        
        Set VueAnnotation = vVueAnnotation
        
        If VueAnnotation.Name = "Objets non affectés" Then
            
            Set colDimensions = ListeDesDimensions(Modele)
            ReDim tabAnnotations(0 To colDimensions.Count - 1)
            
            i = 0
            
            For Each vDimension In colDimensions
                
                Set Dimension = vDimension
                Set Annotation = Dimension.GetAnnotation
                
                Set tabAnnotations(i) = Annotation
                
                i = i + 1
                
            Next vDimension
            
            'VueAnnotation.Activate
            VueAnnotation.MoveAnnotations tabAnnotations
            
            Exit For
            
        End If
        
        
    Next vVueAnnotation
    
End Sub


Private Function ListeDesDimensions(Modele As ModelDoc2) As Collection
    
    On Error GoTo GestErreur
    
    Dim Fonction    As Feature
    Dim Dimension   As DisplayDimension
    
    Set ListeDesDimensions = New Collection
    
    Set Fonction = Modele.FirstFeature
    
    Do Until Fonction Is Nothing
        Set Dimension = Fonction.GetFirstDisplayDimension
        Do Until Dimension Is Nothing
            ListeDesDimensions.Add Dimension
            Set Dimension = Fonction.GetNextDisplayDimension(Dimension)
        Loop
        Set Fonction = Fonction.GetNextFeature
    Loop
    
    Exit Function
    
GestErreur:
    Debug.Print "Erreur [ListeDesDimensions] : " & Err.Number & " ->  " & Err.Description
    Resume Next
    
End Function
