Module Data
    Public Const tailleMap As Integer = 400
    Public tailleCarre As Integer = 11
    Public tabIndiv(tailleMap * tailleMap) As Perso
    'on crée le tableau qui va stocker nos personnages crées avec le bouton générer


    Public Structure Perso
        Dim Position As Position
        Dim NCre As Double
        'niveau de crédulité
        Dim FP As Double
        'force de persuasion
        Dim NCro As Double
        'niveau de croyance 1
        Dim NCro2 As Double
        'niveau de croyance 2
        Dim NCro3 As Double
        'niveau de croyance 3
    End Structure


    Structure Position
        Dim x As Integer
        Dim y As Integer
    End Structure

End Module