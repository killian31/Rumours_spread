Module Traitements

    Public Sub Deplacement(ByVal nbIndiv As Integer,
                           ByVal size As Integer)
        Dim tabDir(8) As String
        tabDir(0) = "Nord"
        tabDir(1) = "Sud"
        tabDir(2) = "Est"
        tabDir(3) = "Ouest"
        tabDir(4) = "NordOuest"
        tabDir(5) = "NordEst"
        tabDir(6) = "SudOuest"
        tabDir(7) = "SudEst"
        Dim rnd As New Random
        For i = 0 To nbIndiv - 1
            Dim direction As String = tabDir(rnd.Next(8))
            Select Case direction
                Case "Nord"
                    If tabIndiv(i).Position.y - size < 0 Then
                        tabIndiv(i).Position.y = tailleMap
                        'la position x ne change pas
                    Else
                        tabIndiv(i).Position.y -= size
                        'la position x ne change pas
                    End If

                Case "Sud"
                    If tabIndiv(i).Position.y + size > tailleMap Then
                        tabIndiv(i).Position.y = 0
                        'la position x ne change pas
                    Else
                        tabIndiv(i).Position.y += size
                        'la position x ne change pas
                    End If

                Case "Est"
                    If tabIndiv(i).Position.x + size > tailleMap Then
                        tabIndiv(i).Position.x = 0
                        'y ne change pas
                    Else
                        tabIndiv(i).Position.x += size
                        'la position y ne change pas
                    End If

                Case "Ouest"
                    If tabIndiv(i).Position.x - size < 0 Then
                        tabIndiv(i).Position.x = tailleMap
                        'y ne change pas
                    Else
                        tabIndiv(i).Position.x -= size
                        'la position y ne change pas
                    End If

                Case "NordOuest"
                    If tabIndiv(i).Position.x - size < 0 And tabIndiv(i).Position.y - size < 0 Then
                        tabIndiv(i).Position.x = tailleMap
                        tabIndiv(i).Position.y = tailleMap
                    ElseIf tabIndiv(i).Position.x - size < 0 Then
                        tabIndiv(i).Position.x = tailleMap
                        tabIndiv(i).Position.y -= size
                    ElseIf tabIndiv(i).Position.y - size < 0 Then
                        tabIndiv(i).Position.y = tailleMap
                        tabIndiv(i).Position.x -= size
                    Else
                        tabIndiv(i).Position.x -= size
                        tabIndiv(i).Position.y -= size
                    End If

                Case "NordEst"
                    If tabIndiv(i).Position.x + size > tailleMap And tabIndiv(i).Position.y - size < 0 Then
                        tabIndiv(i).Position.x = 0
                        tabIndiv(i).Position.y = tailleMap
                    ElseIf tabIndiv(i).Position.x + size > tailleMap Then
                        tabIndiv(i).Position.x = 0
                        tabIndiv(i).Position.y -= size
                    ElseIf tabIndiv(i).Position.y - size < 0 Then
                        tabIndiv(i).Position.y = tailleMap
                        tabIndiv(i).Position.x += size
                    Else
                        tabIndiv(i).Position.x += size
                        tabIndiv(i).Position.y -= size
                    End If

                Case "SudOuest"
                    If tabIndiv(i).Position.x - size < 0 And tabIndiv(i).Position.y + size > tailleMap Then
                        tabIndiv(i).Position.x = tailleMap
                        tabIndiv(i).Position.y = 0
                    ElseIf tabIndiv(i).Position.x - size < 0 Then
                        tabIndiv(i).Position.x = tailleMap
                        tabIndiv(i).Position.y += size
                    ElseIf tabIndiv(i).Position.y + size > tailleMap Then
                        tabIndiv(i).Position.y = 0
                        tabIndiv(i).Position.x -= size
                    Else
                        tabIndiv(i).Position.x -= size
                        tabIndiv(i).Position.y += size
                    End If

                Case "SudEst"
                    If tabIndiv(i).Position.x + size > tailleMap And tabIndiv(i).Position.y + size > tailleMap Then
                        tabIndiv(i).Position.x = 0
                        tabIndiv(i).Position.y = 0
                    ElseIf tabIndiv(i).Position.x + size > tailleMap Then
                        tabIndiv(i).Position.x = 0
                        tabIndiv(i).Position.y += size
                    ElseIf tabIndiv(i).Position.y + size > tailleMap Then
                        tabIndiv(i).Position.y = 0
                        tabIndiv(i).Position.x += size
                    Else
                        tabIndiv(i).Position.x += size
                        tabIndiv(i).Position.y += size
                    End If
            End Select
        Next
    End Sub

    Sub Transmission(ByRef PA As Perso, ByVal PB As Perso)

        If PA.NCro < 1 And PA.NCro2 < 1 And PA.NCro3 < 1 Then

            If PB.NCro > 0.25 Then
                If PB.FP >= 0.8 Then
                    If PA.NCro + (1 / 2) * PA.NCre * (PB.NCro + 1.1 * PB.FP) < 1 Then
                        PA.NCro += (1 / 2) * PA.NCre * (PB.NCro + 1.1 * PB.FP)
                    Else
                        PA.NCro = 1
                    End If
                Else
                    If PA.NCro + (1 / 2) * PA.NCre * (PB.NCro + PB.FP) < 1 Then
                        PA.NCro += (1 / 2) * PA.NCre * (PB.NCro + PB.FP)
                    Else
                        PA.NCro = 1
                    End If
                End If
            Else
                If PB.FP > 0.4 Then
                    If PB.NCro <= 0.2 Then
                        If PA.NCro - (1 / 2) * PA.NCre * (PB.NCro + 1.1 * PB.FP) > 0 Then
                            PA.NCro -= (1 / 2) * PA.NCre * (PB.NCro + 1.1 * PB.FP)
                        Else
                            PA.NCro = 0
                        End If
                    Else
                        If PA.NCro - (1 / 2) * PA.NCre * (PB.NCro + PB.FP) > 0 Then
                            PA.NCro -= (1 / 2) * PA.NCre * (PB.NCro + PB.FP)
                        Else
                            PA.NCro = 0
                        End If
                    End If
                End If
            End If


            If PB.NCro2 > 0.25 Then
                If PB.FP >= 0.8 Then
                    If PA.NCro2 + (1 / 2) * PA.NCre * (PB.NCro2 + 1.1 * PB.FP) < 1 Then
                        PA.NCro2 += (1 / 2) * PA.NCre * (PB.NCro2 + 1.1 * PB.FP)
                    Else
                        PA.NCro2 = 1
                    End If
                Else
                    If PA.NCro2 + (1 / 2) * PA.NCre * (PB.NCro2 + PB.FP) < 1 Then
                        PA.NCro2 += (1 / 2) * PA.NCre * (PB.NCro2 + PB.FP)
                    Else
                        PA.NCro2 = 1
                    End If
                End If
            Else
                If PB.FP > 0.4 Then
                    If PB.NCro2 <= 0.2 Then
                        If PA.NCro2 - (1 / 2) * PA.NCre * (PB.NCro2 + 1.1 * PB.FP) > 0 Then
                            PA.NCro2 -= (1 / 2) * PA.NCre * (PB.NCro2 + 1.1 * PB.FP)
                        Else
                            PA.NCro2 = 0
                        End If
                    Else
                        If PA.NCro2 - (1 / 2) * PA.NCre * (PB.NCro2 + PB.FP) > 0 Then
                            PA.NCro2 -= (1 / 2) * PA.NCre * (PB.NCro2 + PB.FP)
                        Else
                            PA.NCro2 = 0
                        End If
                    End If
                End If
            End If

            If PB.NCro3 > 0.25 Then
                If PB.FP >= 0.8 Then
                    If PA.NCro3 + (1 / 2) * PA.NCre * (PB.NCro3 + 1.1 * PB.FP) < 1 Then
                        PA.NCro3 += (1 / 2) * PA.NCre * (PB.NCro3 + 1.1 * PB.FP)
                    Else
                        PA.NCro3 = 1
                    End If
                Else
                    If PA.NCro3 + (1 / 2) * PA.NCre * (PB.NCro3 + PB.FP) < 1 Then
                        PA.NCro3 += (1 / 2) * PA.NCre * (PB.NCro3 + PB.FP)
                    Else
                        PA.NCro3 = 1
                    End If
                End If
            Else
                If PB.FP > 0.4 Then
                    If PB.NCro3 <= 0.2 Then
                        If PA.NCro3 - (1 / 2) * PA.NCre * (PB.NCro3 + 1.1 * PB.FP) > 0 Then
                            PA.NCro3 -= (1 / 2) * PA.NCre * (PB.NCro3 + 1.1 * PB.FP)
                        Else
                            PA.NCro3 = 0
                        End If
                    Else
                        If PA.NCro3 - (1 / 2) * PA.NCre * (PB.NCro3 + PB.FP) > 0 Then
                            PA.NCro3 -= (1 / 2) * PA.NCre * (PB.NCro3 + PB.FP)
                        Else
                            PA.NCro3 = 0
                        End If
                    End If
                End If
            End If

        End If

    End Sub

    Sub VerifAndTransmission(ByRef P As Perso, ByVal nbIndiv As Integer, ByVal size As Integer)
        For j = 0 To nbIndiv - 1
            'On parcourt le tableau qui contient nos personnages, 
            'On regarde si on trouve des individus à la même position ou à une position voisine de l'individu P :
            For i = 1 To Math.Round(size / 2, 0)
                If tabIndiv(j).Position.x = P.Position.x And tabIndiv(j).Position.y = P.Position.y Then
                    Transmission(P, tabIndiv(j))
                End If
                If tabIndiv(j).Position.x = P.Position.x + i And tabIndiv(j).Position.y = P.Position.y - i Then
                    Transmission(P, tabIndiv(j))
                End If
                If tabIndiv(j).Position.x = P.Position.x + i And tabIndiv(j).Position.y = P.Position.y Then
                    Transmission(P, tabIndiv(j))
                End If
                If tabIndiv(j).Position.x = P.Position.x + i And tabIndiv(j).Position.y = P.Position.y + i Then
                    Transmission(P, tabIndiv(j))
                End If
                If tabIndiv(j).Position.x = P.Position.x And tabIndiv(j).Position.y = P.Position.y + i Then
                    Transmission(P, tabIndiv(j))
                End If
                If tabIndiv(j).Position.x = P.Position.x - i And tabIndiv(j).Position.y = P.Position.y + i Then
                    Transmission(P, tabIndiv(j))
                End If
                If tabIndiv(j).Position.x = P.Position.x - i And tabIndiv(j).Position.y = P.Position.y Then
                    Transmission(P, tabIndiv(j))
                End If
                If tabIndiv(j).Position.x = P.Position.x - i And tabIndiv(j).Position.y = P.Position.y - i Then
                    Transmission(P, tabIndiv(j))
                End If
                If tabIndiv(j).Position.x = P.Position.x And tabIndiv(j).Position.y = P.Position.y - i Then
                    Transmission(P, tabIndiv(j))
                End If
            Next
        Next
    End Sub

    Public Function maxValue(ByVal a As Double, ByVal b As Double) As Double
        'simple fonction qui détermine le maximum entre 2 nombres en paramètre
        If a >= b Then
            Return a
        Else
            Return b
        End If
    End Function

    Public Sub generation(ByRef P As Perso)
        'Cette procédure permet la création d'un seul individu, non croyant
        Dim rnd As New Random
        'On donne des coordonnées aléatoires à l'individu P en paramètre
        P.Position.x = rnd.Next(0, tailleMap)
        P.Position.y = rnd.Next(0, tailleMap)
        'on affecte des valeurs aléatoires aux niveaux de crédulité, de croyance et force de persuasion de l'individu
        Dim rndCre As Double
        Dim rndCro As Double
        Dim rndCro2 As Double
        Dim rndCro3 As Double
        Dim rndFp As Double
        'chiffre des décimales aléatoire auquel on ajoute le chiffre des centièmes aléatoire
        rndCre = rnd.Next(1, 10) / 100 + rnd.Next(0, 10) / 10 
        rndFp = rnd.Next(1, 10) / 100 + rnd.Next(0, 10) / 10
        rndCro = 0
        rndCro2 = 0
        rndCro3 = 0
        P.NCre = rndCre
        P.NCro = rndCro
        P.NCro2 = rndCro2
        P.NCro3 = rndCro3
        P.FP = rndFp
    End Sub

    Public Sub generation_croy1(ByRef P As Perso)
        'Cette procédure permet la création d'un seul individu, croyant à la croyance 1
        Dim rnd As New Random
        'On donne des coordonnées aléatoires à l'individu P en paramètre
        P.Position.x = rnd.Next(0, tailleMap)
        P.Position.y = rnd.Next(0, tailleMap)
        'on affecte des valeurs aléatoires aux niveaux de crédulité, de croyance et force de persuasion de l'individu
        Dim rndCre As Double
        Dim rndCro As Double
        Dim rndCro2 As Double
        Dim rndCro3 As Double
        Dim rndFp As Double
        'chiffre des décimales aléatoire auquel on ajoute le chiffre des centièmes aléatoire
        rndCre = rnd.Next(1, 10) / 100 + rnd.Next(0, 10) / 10
        rndFp = rnd.Next(1, 10) / 100 + rnd.Next(0, 10) / 10
        rndCro = rnd.Next(0, 10) / 100 + rnd.Next(5, 10) / 10
        rndCro2 = 0
        rndCro3 = 0
        P.NCre = rndCre
        P.NCro = rndCro
        P.NCro2 = rndCro2
        P.NCro3 = rndCro3
        P.FP = rndFp
    End Sub

    Public Sub generation_croy2(ByRef P As Perso)
        'Cette procédure permet la création d'un seul individu, croyant à la croyance 2
        Dim rnd As New Random
        'On donne des coordonnées aléatoires à l'individu P en paramètre
        P.Position.x = rnd.Next(0, tailleMap)
        P.Position.y = rnd.Next(0, tailleMap)
        'on affecte des valeurs aléatoires aux niveaux de crédulité, de croyance et force de persuasion de l'individu
        Dim rndCre As Double
        Dim rndCro As Double
        Dim rndCro2 As Double
        Dim rndCro3 As Double
        Dim rndFp As Double
        'chiffre des décimales aléatoire auquel on ajoute le chiffre des centièmes aléatoire
        rndCre = rnd.Next(1, 10) / 100 + rnd.Next(0, 10) / 10
        rndFp = rnd.Next(1, 10) / 100 + rnd.Next(0, 10) / 10
        rndCro = 0
        rndCro2 = rnd.Next(0, 10) / 100 + rnd.Next(5, 10) / 10
        rndCro3 = 0
        P.NCre = rndCre
        P.NCro = rndCro
        P.NCro2 = rndCro2
        P.NCro3 = rndCro3
        P.FP = rndFp
    End Sub

    Public Sub generation_croy3(ByRef P As Perso)
        'Cette procédure permet la création d'un seul individu, croyant à la croyance 3
        Dim rnd As New Random
        'On donne des coordonnées aléatoires à l'individu P en paramètre
        P.Position.x = rnd.Next(0, tailleMap)
        P.Position.y = rnd.Next(0, tailleMap)
        'on affecte des valeurs aléatoires aux niveaux de crédulité, de croyance et force de persuasion de l'individu
        Dim rndCre As Double
        Dim rndCro As Double
        Dim rndCro2 As Double
        Dim rndCro3 As Double
        Dim rndFp As Double
        'chiffre des décimales aléatoire auquel on ajoute le chiffre des centièmes aléatoire
        rndCre = rnd.Next(1, 10) / 100 + rnd.Next(0, 10) / 10
        rndFp = rnd.Next(1, 10) / 100 + rnd.Next(0, 10) / 10
        rndCro = 0
        rndCro2 = 0
        rndCro3 = rnd.Next(0, 10) / 100 + rnd.Next(5, 10) / 10
        P.NCre = rndCre
        P.NCro = rndCro
        P.NCro2 = rndCro2
        P.NCro3 = rndCro3
        P.FP = rndFp
    End Sub

End Module
