Public Class Form1

    'Variables utilisées par plusieurs procédures :
    Public Feuille As Bitmap = New Bitmap(tailleMap + 1, tailleMap + 1)
    Public Dessin As Graphics = Graphics.FromImage(Feuille)

    Dim simulation As Boolean

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'initialisation des trackbars
        TrackBar1.SetRange(1, 0.01 * tailleMap * tailleMap)
        TrackBar2.SetRange(0, 0.01 * tailleMap * tailleMap)
        TrackBar3.SetRange(0, 0.01 * tailleMap * tailleMap)
        TrackBar4.SetRange(0, 0.01 * tailleMap * tailleMap)


        'Initialisation des textBox, etc
        txtIndiv.Text = 0
        txtCroyants.Text = 0
        txtCro2.Text = 0
        txtCro3.Text = 0
        TrackBar1.Value = TrackBar1.Minimum
        TrackBar2.Value = TrackBar2.Minimum
        TrackBar3.Value = TrackBar3.Minimum
        TrackBar4.Value = TrackBar4.Minimum
        cbColor1.Text = "Rouge"
        cbColor2.Text = "Vert"
        cbColor3.Text = "Bleu"
        'initialisation de la picture box
        Dessin.Clear(Color.White)
        pbMap.Image = Feuille

        'on n'est pas en train de simuler
        simulation = False 'ce booléen servira à renforcer la robustesse de l'application
    End Sub


    Private Sub btnGenerer_Click(sender As Object, e As EventArgs) Handles btnGenerer.Click
        btnGenerer.Cursor = Cursors.WaitCursor
        'Initialisation des labels
        lblTime.Text = 0
        lblCroyants.Text = 0
        lblCroyants2.Text = 0
        lblCroyants3.Text = 0

        Dim rnd As New Random

        'On récupère les informations saisies par l'utilisateur (les nombres valent 0 si il n'a rien écrit)
        Dim nbCroyants As Integer
        Dim nbCroyants2 As Integer
        Dim nbCroyants3 As Integer
        Dim nbIndiv As Integer = CInt(txtIndiv.Text)

        If txtCroyants.Text <> "" Then
            nbCroyants = CInt(txtCroyants.Text)
        Else
            nbCroyants = 0
        End If
        If txtCro2.Text <> "" Then
            nbCroyants2 = CInt(txtCro2.Text)
        Else
            nbCroyants2 = 0
        End If
        If txtCro3.Text <> "" Then
            nbCroyants3 = CInt(txtCro3.Text)
        Else
            nbCroyants3 = 0
        End If

        'On supprime les potentiels anciens dessins de la picturebox
        Dessin.Clear(Color.White)

        'Couleurs du graph

        Select Case cbColor1.Text
            Case "Rouge"
                Me.Chart2.Series("Croyance 1").Color = Color.Red
            Case "Vert"
                Me.Chart2.Series("Croyance 1").Color = Color.Green
            Case "Bleu"
                Me.Chart2.Series("Croyance 1").Color = Color.Blue
            Case Else
                Me.Chart2.Series("Croyance 1").Color = Color.Red
        End Select

        Select Case cbColor2.Text
            Case "Rouge"
                Me.Chart2.Series("Croyance 2").Color = Color.Red
            Case "Vert"
                Me.Chart2.Series("Croyance 2").Color = Color.Green
            Case "Bleu"
                Me.Chart2.Series("Croyance 2").Color = Color.Blue
            Case Else
                Me.Chart2.Series("Croyance 2").Color = Color.Green
        End Select

        Select Case cbColor3.Text
            Case "Rouge"
                Me.Chart2.Series("Croyance 3").Color = Color.Red
            Case "Vert"
                Me.Chart2.Series("Croyance 3").Color = Color.Green
            Case "Bleu"
                Me.Chart2.Series("Croyance 3").Color = Color.Blue
            Case Else
                Me.Chart2.Series("Croyance 3").Color = Color.Blue
        End Select



        'La génération

        If nbIndiv > 0.1 * (tailleMap * tailleMap) Then
            MsgBox("Le nombre d'individus est trop élevé, la simulation prendrait trop de temps.")
        ElseIf nbCroyants + nbCroyants2 + nbCroyants3 > nbIndiv Then
            MsgBox("Le nombre de croyants ne peut pas être plus grand que le nombre total d'individus !")
        Else
            'création des croyants 1
            For i = 0 To nbCroyants - 1
                If CheckBox1.Checked Then
                    If i = 0 Then 'le premier individu n'a pas besoin de vérifier si il peut se placer 
                        generation_croy1(tabIndiv(i))
                        'dessin :
                        Select Case cbColor1.Text
                            Case "Rouge"
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Vert"
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Bleu"
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case Else
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)

                        End Select

                        ProgressBar1.Value = 100 * i / nbIndiv
                        lblCroyants.Text = 1
                        lblIndiv.Text = 1
                    Else
                        generation_croy1(tabIndiv(i))

                        'Vérification que les individus n'ont pas la même position :
                        Dim temp As Boolean = True
                        While temp = True
                            temp = False
                            For j = 0 To i - 1 'on parcourt le tableau avant la position de l'individu qu'on est en train de créer
                                If tabIndiv(j).Position.x = tabIndiv(i).Position.x And tabIndiv(j).Position.y = tabIndiv(i).Position.y Then
                                    tabIndiv(i).Position.x = rnd.Next(0, tailleMap)
                                    tabIndiv(i).Position.y = rnd.Next(0, tailleMap)
                                    temp = True
                                    Exit For
                                End If
                            Next
                        End While

                        'dessin :
                        Select Case cbColor1.Text
                            Case "Rouge"
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Vert"
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Bleu"
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case Else
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                        End Select

                        ProgressBar1.Value = 100 * i / nbIndiv

                        lblCroyants.Text += 1
                        lblIndiv.Text += 1
                    End If
                Else
                    generation(tabIndiv(i))

                    'on va vérifier en parcourant le tableau que les coordonnées affectées n'existent pas déjà, 
                    'si elles existent on lui en donne de nouvelles et on recommence la vérification :

                    Dim temp As Boolean = True
                    While temp = True
                        temp = False
                        For j = 0 To i - 1 'on parcourt la partie gauche du tableau (tous les éléments situés avant l'élément que l'on teste)
                            If tabIndiv(j).Position.x = tabIndiv(i).Position.x And tabIndiv(j).Position.y = tabIndiv(i).Position.y Then
                                'on donne de nouvelles coordonnées aléatoires si on trouve un élément ayant les mêmes coordonnées
                                tabIndiv(i).Position.x = rnd.Next(0, tailleMap)
                                tabIndiv(i).Position.y = rnd.Next(0, tailleMap)
                                temp = True
                            End If
                        Next
                    End While

                    ProgressBar1.Value = 100 * i / nbIndiv 'une fois que l'individu i a des coordonnées, on fait avancer la progressbar

                    'on dessine le point sur la pb :

                    Dessin.FillRectangle(Brushes.Black, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)

                    lblIndiv.Text += 1 'on met à jour l'affichage du nombre d'individus 

                End If
            Next

            'croyance 2 :

            For i = nbCroyants To nbCroyants + nbCroyants2 - 1
                If CheckBox2.Checked Then
                    If i = 0 Then 'le premier individu n'a pas besoin de vérifier si il peut se placer 
                        generation_croy2(tabIndiv(i))

                        'dessin :
                        Select Case cbColor2.Text
                            Case "Rouge"
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Vert"
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Bleu"
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case Else
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                        End Select

                        ProgressBar1.Value = 100 * i / nbIndiv
                        lblCroyants2.Text = 1
                        lblIndiv.Text = 1
                    Else
                        generation_croy2(tabIndiv(i))

                        'Vérification que les individus n'ont pas la même position :
                        Dim temp As Boolean = True
                        While temp = True
                            temp = False
                            For j = 0 To i - 1 'on parcourt le tableau avant la position de l'individu qu'on est en train de créer
                                If tabIndiv(j).Position.x = tabIndiv(i).Position.x And tabIndiv(j).Position.y = tabIndiv(i).Position.y Then
                                    tabIndiv(i).Position.x = rnd.Next(0, tailleMap)
                                    tabIndiv(i).Position.y = rnd.Next(0, tailleMap)
                                    temp = True
                                End If
                            Next
                        End While

                        'dessin :
                        Select Case cbColor2.Text
                            Case "Rouge"
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Vert"
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Bleu"
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case Else
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                        End Select

                        ProgressBar1.Value = 100 * i / nbIndiv

                        lblCroyants2.Text += 1
                        lblIndiv.Text += 1
                    End If
                Else 'on place des indiv normaux si la croyance n'est pas cochée

                    generation(tabIndiv(i))
                    'on va vérifier en parcourant le tableau que les coordonnées affectées n'existent pas déjà, 
                    'si elles existent on lui en donne de nouvelles et on recommence la vérification :

                    Dim temp As Boolean = True
                    While temp = True
                        temp = False
                        For j = 0 To i - 1 'on parcourt la partie gauche du tableau (tous les éléments situés avant l'élément que l'on teste)
                            If tabIndiv(j).Position.x = tabIndiv(i).Position.x And tabIndiv(j).Position.y = tabIndiv(i).Position.y Then
                                'on donne de nouvelles coordonnées aléatoires si on trouve un élément ayant les mêmes coordonnées
                                tabIndiv(i).Position.x = rnd.Next(0, tailleMap)
                                tabIndiv(i).Position.y = rnd.Next(0, tailleMap)
                                temp = True
                            End If
                        Next
                    End While

                    ProgressBar1.Value = 100 * i / nbIndiv 'une fois que l'individu i a des coordonnées, on fait avancer la progressbar

                    'on dessine le point sur la pb :

                    Dessin.FillRectangle(Brushes.Black, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)


                    lblIndiv.Text += 1 'on met à jour l'affichage du nombre d'individus 

                End If
            Next

            'croyance 3 :

            For i = nbCroyants + nbCroyants2 To nbCroyants + nbCroyants2 + nbCroyants3 - 1
                If CheckBox3.Checked Then
                    If i = 0 Then 'le premier individu n'a pas besoin de vérifier si il peut se placer 
                        generation_croy3(tabIndiv(i))

                        'dessin :
                        Select Case cbColor3.Text
                            Case "Rouge"
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Vert"
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Bleu"
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case Else
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                        End Select

                        ProgressBar1.Value = 100 * i / nbIndiv
                        lblCroyants3.Text = 1
                        lblIndiv.Text = 1
                    Else
                        generation_croy3(tabIndiv(i))

                        'Vérification que les individus n'ont pas la même position :
                        Dim temp As Boolean = True
                        While temp = True
                            temp = False
                            For j = 0 To i - 1 'on parcourt le tableau avant la position de l'individu qu'on est en train de créer
                                If tabIndiv(j).Position.x = tabIndiv(i).Position.x And tabIndiv(j).Position.y = tabIndiv(i).Position.y Then
                                    tabIndiv(i).Position.x = rnd.Next(0, tailleMap)
                                    tabIndiv(i).Position.y = rnd.Next(0, tailleMap)
                                    temp = True
                                End If
                            Next
                        End While

                        'dessin :
                        Select Case cbColor3.Text
                            Case "Rouge"
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Vert"
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Bleu"
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case Else
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                        End Select

                        ProgressBar1.Value = 100 * i / nbIndiv

                        lblCroyants3.Text += 1
                        lblIndiv.Text += 1
                    End If
                Else

                    generation(tabIndiv(i))
                    'on va vérifier en parcourant le tableau que les coordonnées affectées n'existent pas déjà, 
                    'si elles existent on lui en donne de nouvelles et on recommence la vérification :

                    Dim temp As Boolean = True
                    While temp = True
                        temp = False
                        For j = 0 To i - 1 'on parcourt la partie gauche du tableau (tous les éléments situés avant l'élément que l'on teste)
                            If tabIndiv(j).Position.x = tabIndiv(i).Position.x And tabIndiv(j).Position.y = tabIndiv(i).Position.y Then
                                'on donne de nouvelles coordonnées aléatoires si on trouve un élément ayant les mêmes coordonnées
                                tabIndiv(i).Position.x = rnd.Next(0, tailleMap)
                                tabIndiv(i).Position.y = rnd.Next(0, tailleMap)
                                temp = True
                            End If
                        Next
                    End While

                    ProgressBar1.Value = 100 * i / nbIndiv 'une fois que l'individu i a des coordonnées, on fait avancer la progressbar

                    'on dessine le point (carré) sur la pb :

                    Dessin.FillRectangle(Brushes.Black, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)


                    lblIndiv.Text += 1 'on met à jour l'affichage du nombre d'individus 

                End If
            Next

            'on place les autres :

            For i = nbCroyants + nbCroyants2 + nbCroyants3 To nbIndiv - 1
                If i = 0 Then 'cette condition est vérifiée si le nombre de croyants est de 0
                    generation(tabIndiv(i))
                    'on dessine un point sur la picture box :
                    Dessin.FillRectangle(Brushes.Black, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                    'on fait progresser la progressbar
                    ProgressBar1.Value = 100 * i / nbIndiv
                    'on met à jour l'affichage du nombre de croyants et d'individus
                    lblIndiv.Text = 1
                    lblCroyants.Text = 0
                    lblCroyants2.Text = 0
                    lblCroyants3.Text = 0
                Else
                    generation(tabIndiv(i))
                    'on va vérifier en parcourant le tableau que les coordonnées affectées n'existent pas déjà, 
                    'si elles existent on lui en donne de nouvelles et on recommence la vérification :

                    Dim temp As Boolean = True
                    While temp = True
                        temp = False
                        For j = 0 To i - 1 'on parcourt la partie gauche du tableau (tous les éléments situés avant l'élément que l'on teste)
                            If tabIndiv(j).Position.x = tabIndiv(i).Position.x And tabIndiv(j).Position.y = tabIndiv(i).Position.y Then
                                'on donne de nouvelles coordonnées aléatoires si on trouve un élément ayant les mêmes coordonnées
                                tabIndiv(i).Position.x = rnd.Next(0, tailleMap)
                                tabIndiv(i).Position.y = rnd.Next(0, tailleMap)
                                temp = True
                            End If
                        Next
                    End While

                    ProgressBar1.Value = 100 * i / nbIndiv 'une fois que l'individu i a des coordonnées, on fait avancer la progressbar

                    'on dessine le point sur la pb :

                    Dessin.FillRectangle(Brushes.Black, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)

                    lblIndiv.Text += 1 'on met à jour l'affichage du nombre d'individus 

                End If
            Next

            Dim totalCre As Double = 0
            Dim totalCre2 As Double = 0
            Dim totalCre3 As Double = 0

            Dim totalCro As Double = 0
            Dim totalCro2 As Double = 0
            Dim totalCro3 As Double = 0

            Dim totalFP As Double = 0
            Dim totalFP2 As Double = 0
            Dim totalFP3 As Double = 0

            'on calcule les niveaux totaux pour en faire la mmoyenne :

            For j = 0 To nbCroyants - 1
                totalCre += tabIndiv(j).NCre

                totalFP += tabIndiv(j).FP
            Next
            For j = nbCroyants To nbCroyants + nbCroyants2 - 1
                totalCre2 += tabIndiv(j).NCre

                totalFP2 += tabIndiv(j).FP
            Next
            For j = nbCroyants + nbCroyants2 To nbCroyants + nbCroyants2 + nbCroyants3 - 1
                totalCre3 += tabIndiv(j).NCre

                totalFP3 += tabIndiv(j).FP
            Next

            For i = 0 To nbIndiv - 1
                totalCro += tabIndiv(i).NCro
                totalCro2 += tabIndiv(i).NCro2
                totalCro3 += tabIndiv(i).NCro3
            Next

            lblTotCroy.Text = nbCroyants + nbCroyants2 + nbCroyants3

            If nbCroyants = 0 Then
                lblCreMoy.Text = 0
                lblFPMoy.Text = 0
            Else
                lblCreMoy.Text = Math.Round(totalCre / nbCroyants, 4)
                lblFPMoy.Text = Math.Round(totalFP / nbCroyants, 4)
            End If

            If nbCroyants2 = 0 Then
                lblCreMoy2.Text = 0
                lblFPMoy2.Text = 0
            Else
                lblCreMoy2.Text = Math.Round(totalCre2 / nbCroyants2, 4)
                lblFPMoy2.Text = Math.Round(totalFP2 / nbCroyants2, 4)
            End If

            If nbCroyants3 = 0 Then
                lblCreMoy3.Text = 0
                lblFPMoy3.Text = 0
            Else
                lblCreMoy3.Text = Math.Round(totalCre3 / nbCroyants3, 4)
                lblFPMoy3.Text = Math.Round(totalFP3 / nbCroyants3, 4)
            End If

            lblCroMoy.Text = Math.Round(totalCro / nbIndiv, 4)
            lblCroMoy2.Text = Math.Round(totalCro2 / nbIndiv, 4)
            lblCroMoy3.Text = Math.Round(totalCro3 / nbIndiv, 4)

            pbMap.Image = Feuille
            ProgressBar1.Value = 100
            Me.Chart1.Series("Nombre de croyants").Points.Clear()
            Me.Chart1.Series("Nombre de croyants").Points.AddXY("Croyance 1", nbCroyants)
            Me.Chart1.Series("Nombre de croyants").Points.AddXY("Croyance 2", nbCroyants2)
            Me.Chart1.Series("Nombre de croyants").Points.AddXY("Croyance 3", nbCroyants3)

            Me.Chart2.Series("Croyance 1").Points.Clear()
            Me.Chart2.Series("Croyance 2").Points.Clear()
            Me.Chart2.Series("Croyance 3").Points.Clear()
            Me.Chart2.Series("Croyance 1").Points.AddXY(0, nbCroyants)
            Me.Chart2.Series("Croyance 2").Points.AddXY(0, nbCroyants2)
            Me.Chart2.Series("Croyance 3").Points.AddXY(0, nbCroyants3)

        End If
        btnGenerer.Cursor = Cursors.Hand
    End Sub

    Private Sub btnRandom_Click(sender As Object, e As EventArgs) Handles btnRandom.Click
        Dim rnd As New Random
        txtIndiv.Text = rnd.Next(0.01 * (tailleMap * tailleMap)) 'on ne prend que 1% de l'espace total pour que le nb d'individus ne soit pas trop élevé

        Dim nbInd As Integer = CInt(txtIndiv.Text)

        Dim nbC1 As Integer = rnd.Next(nbInd)
        Dim nbC2 As Integer = rnd.Next(nbInd)
        Dim nbC3 As Integer = rnd.Next(nbInd)

        While nbC1 + nbC2 + nbC3 >= nbInd
            nbC1 = rnd.Next(nbInd)
            nbC2 = rnd.Next(nbInd)
            nbC3 = rnd.Next(nbInd)
        End While

        txtCroyants.Text = nbC1
        txtCro2.Text = nbC2
        txtCro3.Text = nbC3
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        Dim nbIndiv As Integer = CInt(lblIndiv.Text)
        Dim nbCroyants As Integer = CInt(lblCroyants.Text)
        Dim nbCroyants2 As Integer = CInt(lblCroyants2.Text)
        Dim nbCroyants3 As Integer = CInt(lblCroyants3.Text)

        'le déplacement et la mise à jour des stats ne se fait que si le nombre total de croyants est inférieur au nombre total d'individus
        If nbIndiv <> nbCroyants Then
            lblLoading.Text = "Propagations des rumeurs  :"
            lblTime.Text += 1

            Dessin.Clear(Color.White)

            'on déplace les individus
            Deplacement(nbIndiv, Int(tailleCarre / 2))

            'on transmet
            For i = 0 To nbIndiv - 1
                VerifAndTransmission(tabIndiv(i), nbIndiv, tailleCarre)
                Dim max1 As Double = maxValue(tabIndiv(i).NCro, tabIndiv(i).NCro2)
                Dim max2 As Double = maxValue(max1, tabIndiv(i).NCro3)
                ' on représente chaque individu de la couleur de sa croyance dominante
                If max2 = tabIndiv(i).NCro Then
                    If tabIndiv(i).NCro >= 0.5 Then
                        Select Case cbColor1.Text
                            Case "Rouge"
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Vert"
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Bleu"
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case Else
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                        End Select
                    Else

                        Dessin.FillRectangle(Brushes.Black, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                    End If
                ElseIf max2 = tabIndiv(i).NCro2 Then
                    If tabIndiv(i).NCro2 >= 0.5 Then
                        Select Case cbColor2.Text
                            Case "Rouge"
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Vert"
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Bleu"
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case Else
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                        End Select
                    Else
                        Dessin.FillRectangle(Brushes.Black, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                    End If
                Else
                    If tabIndiv(i).NCro3 >= 0.5 Then
                        Select Case cbColor3.Text
                            Case "Rouge"
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Vert"
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Bleu"
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case Else
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                        End Select
                    Else
                        Dessin.FillRectangle(Brushes.Black, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                    End If
                End If
            Next

            'on recalcule le total de croyants de chaque croyance
            Dim totCroy As Integer = 0
            Dim totCroy2 As Integer = 0
            Dim totCroy3 As Integer = 0
            For i = 0 To nbIndiv - 1
                Dim max1 As Double = maxValue(tabIndiv(i).NCro, tabIndiv(i).NCro2)
                Dim max2 As Double = maxValue(max1, tabIndiv(i).NCro3)
                If max2 = tabIndiv(i).NCro Then
                    If tabIndiv(i).NCro >= 0.5 Then
                        totCroy += 1
                    End If
                ElseIf max2 = tabIndiv(i).NCro2 Then
                    If tabIndiv(i).NCro2 >= 0.5 Then
                        totCroy2 += 1
                    End If
                Else
                    If tabIndiv(i).NCro3 >= 0.5 Then
                        totCroy3 += 1
                    End If
                End If
            Next

            ProgressBar1.Value = 100 * (totCroy + totCroy2 + totCroy3) / nbIndiv

            'on affiche les moyennes
            
            lblPourcentage.Text = Math.Round(100 * totCroy / nbIndiv, 2)
            lblPourcentage2.Text = Math.Round(100 * totCroy2 / nbIndiv, 2)
            lblPourcentage3.Text = Math.Round(100 * totCroy3 / nbIndiv, 2)

            lblCroyants.Text = totCroy
            lblCroyants2.Text = totCroy2
            lblCroyants3.Text = totCroy3


            lblCroMoy.Text = Math.Round(totCroy / nbIndiv, 3)
            lblCroMoy2.Text = Math.Round(totCroy2 / nbIndiv, 3)
            lblCroMoy3.Text = Math.Round(totCroy3 / nbIndiv, 3)
            lblTotCroy.Text = totCroy + totCroy2 + totCroy3

            pbMap.Image = Feuille

            'mise à jour du graph et de l'histogramme
            Me.Chart1.Series("Nombre de croyants").Points.Clear()
            Me.Chart1.Series("Nombre de croyants").Points.AddXY("Croyance 1", totCroy)
            Me.Chart1.Series("Nombre de croyants").Points.AddXY("Croyance 2", totCroy2)
            Me.Chart1.Series("Nombre de croyants").Points.AddXY("Croyance 3", totCroy3)

            Me.Chart2.Series("Croyance 1").Points.AddXY(lblTime.Text, totCroy)
            Me.Chart2.Series("Croyance 2").Points.AddXY(lblTime.Text, totCroy2)
            Me.Chart2.Series("Croyance 3").Points.AddXY(lblTime.Text, totCroy3)
        Else
            simulation = False
            Timer.Stop()
            'lorsque la simulation est finie, on affiche un message pour donner le nombre de périodes écoulées et la couleur de la croyance ayant le plus de croyants
            If nbCroyants > nbCroyants2 And nbCroyants > nbCroyants3 Then
                MsgBox("Toute la population est croyante, en " & lblTime.Text & " périodes !" & " Couleur Victorieuse : " & cbColor1.Text)
            ElseIf nbCroyants2 > nbCroyants And nbCroyants2 > nbCroyants3 Then
                MsgBox("Toute la population est croyante, en " & lblTime.Text & " périodes !" & " Couleur Victorieuse : " & cbColor2.Text)
            ElseIf nbCroyants3 > nbCroyants And nbCroyants3 > nbCroyants2 Then
                MsgBox("Toute la population est croyante, en " & lblTime.Text & " périodes !" & " Couleur Victorieuse : " & cbColor3.Text)
            ElseIf nbCroyants = nbCroyants2 And nbCroyants > nbCroyants3 Then
                MsgBox("Toute la population est croyante, en " & lblTime.Text & " périodes !" & " Couleurs Victorieuses : " & cbColor1.Text & ", " & cbColor2.Text)
            ElseIf nbCroyants = nbCroyants3 And nbCroyants > nbCroyants2 Then
                MsgBox("Toute la population est croyante, en " & lblTime.Text & " périodes !" & " Couleurs Victorieuses : " & cbColor1.Text & ", " & cbColor3.Text)
            ElseIf nbCroyants2 = nbCroyants3 And nbCroyants2 > nbCroyants Then
                MsgBox("Toute la population est croyante, en " & lblTime.Text & " périodes !" & " Couleurs Victorieuses : " & cbColor2.Text & ", " & cbColor3.Text)
            End If
            btnLecture.BackColor = Color.Red
        End If
    End Sub


    Private Sub btnLecture_Click(sender As Object, e As EventArgs) Handles btnLecture.Click
        Select Case cbSimuTime.Text
        'on récupère la vitesse de simulation saisie par l'utilisateur
            Case "x1"
                Timer.Interval = 1000
            Case "x2"
                Timer.Interval = 500
            Case "x5"
                Timer.Interval = 200
            Case "x10"
                Timer.Interval = 100
            Case "x100"
                Timer.Interval = 10
        End Select

        If btnLecture.BackColor = Color.Red Then
            btnLecture.BackColor = Color.Blue
            simulation = True
            Timer.Start()
        Else
            btnLecture.BackColor = Color.Red
            simulation = False
            Timer.Stop()
        End If

    End Sub

    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles Timer.Tick

        Dim nbIndiv As Integer = CInt(lblIndiv.Text)
        Dim nbCroyants As Integer = CInt(lblCroyants.Text)
        Dim nbCroyants2 As Integer = CInt(lblCroyants2.Text)
        Dim nbCroyants3 As Integer = CInt(lblCroyants3.Text)

        If nbIndiv <> nbCroyants + nbCroyants2 + nbCroyants3 Then
            lblLoading.Text = "Propagation des rumeurs :"
            lblTime.Text += 1

            Dessin.Clear(Color.White)

            Deplacement(nbIndiv, Int(tailleCarre / 2))

            For i = 0 To nbIndiv - 1
                VerifAndTransmission(tabIndiv(i), nbIndiv, tailleCarre)
                Dim max1 As Double = maxValue(tabIndiv(i).NCro, tabIndiv(i).NCro2)
                Dim max2 As Double = maxValue(max1, tabIndiv(i).NCro3)
                If max2 = tabIndiv(i).NCro Then
                    If tabIndiv(i).NCro >= 0.5 Then
                        Select Case cbColor1.Text
                            Case "Rouge"
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Vert"
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Bleu"
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case Else
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                        End Select
                    Else
                        Dessin.FillRectangle(Brushes.Black, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                    End If
                ElseIf max2 = tabIndiv(i).NCro2 Then
                    If tabIndiv(i).NCro2 >= 0.5 Then
                        Select Case cbColor2.Text
                            Case "Rouge"
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Vert"
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Bleu"
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case Else
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                        End Select
                    Else
                        Dessin.FillRectangle(Brushes.Black, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                    End If
                Else
                    If tabIndiv(i).NCro3 >= 0.5 Then
                        Select Case cbColor3.Text
                            Case "Rouge"
                                Dessin.FillRectangle(Brushes.Red, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Vert"
                                Dessin.FillRectangle(Brushes.Green, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case "Bleu"
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                            Case Else
                                Dessin.FillRectangle(Brushes.Blue, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                        End Select
                    Else
                        Dessin.FillRectangle(Brushes.Black, tabIndiv(i).Position.x, tabIndiv(i).Position.y, tailleCarre, tailleCarre)
                    End If
                End If
            Next
            pbMap.Image = Feuille

            Dim totCroy As Integer = 0
            Dim totCroy2 As Integer = 0
            Dim totCroy3 As Integer = 0
            For i = 0 To nbIndiv - 1
                Dim max1 As Double = maxValue(tabIndiv(i).NCro, tabIndiv(i).NCro2)
                Dim max2 As Double = maxValue(max1, tabIndiv(i).NCro3)
                If max2 = tabIndiv(i).NCro Then
                    If tabIndiv(i).NCro >= 0.5 Then
                        totCroy += 1
                    End If
                ElseIf max2 = tabIndiv(i).NCro2 Then
                    If tabIndiv(i).NCro2 >= 0.5 Then
                        totCroy2 += 1
                    End If
                Else
                    If tabIndiv(i).NCro3 >= 0.5 Then
                        totCroy3 += 1
                    End If
                End If
            Next

            ProgressBar1.Value = 100 * (totCroy + totCroy2 + totCroy3) / nbIndiv

            lblPourcentage.Text = Math.Round(100 * totCroy / nbIndiv, 2)
            lblPourcentage2.Text = Math.Round(100 * totCroy2 / nbIndiv, 2)
            lblPourcentage3.Text = Math.Round(100 * totCroy3 / nbIndiv, 2)

            lblCroyants.Text = totCroy
            lblCroyants2.Text = totCroy2
            lblCroyants3.Text = totCroy3

            lblCroMoy.Text = Math.Round(totCroy / nbIndiv, 3)
            lblCroMoy2.Text = Math.Round(totCroy2 / nbIndiv, 3)
            lblCroMoy3.Text = Math.Round(totCroy3 / nbIndiv, 3)
            lblTotCroy.Text = totCroy + totCroy2 + totCroy3

            Me.Chart1.Series("Nombre de croyants").Points.Clear()
            Me.Chart1.Series("Nombre de croyants").Points.AddXY("Croyance 1", totCroy)
            Me.Chart1.Series("Nombre de croyants").Points.AddXY("Croyance 2", totCroy2)
            Me.Chart1.Series("Nombre de croyants").Points.AddXY("Croyance 3", totCroy3)

            Me.Chart2.Series("Croyance 1").Points.AddXY(lblTime.Text, totCroy)
            Me.Chart2.Series("Croyance 2").Points.AddXY(lblTime.Text, totCroy2)
            Me.Chart2.Series("Croyance 3").Points.AddXY(lblTime.Text, totCroy3)
        Else
            simulation = False
            Timer.Stop()

            If nbCroyants > nbCroyants2 And nbCroyants > nbCroyants3 Then
                MsgBox("Toute la population est croyante, en " & lblTime.Text & " périodes !" & " Couleur Victorieuse : " & cbColor1.Text)
            ElseIf nbCroyants2 > nbCroyants And nbCroyants2 > nbCroyants3 Then
                MsgBox("Toute la population est croyante, en " & lblTime.Text & " périodes !" & " Couleur Victorieuse : " & cbColor2.Text)
            ElseIf nbCroyants3 > nbCroyants And nbCroyants3 > nbCroyants2 Then
                MsgBox("Toute la population est croyante, en " & lblTime.Text & " périodes !" & " Couleur Victorieuse : " & cbColor3.Text)
            ElseIf nbCroyants = nbCroyants2 And nbCroyants > nbCroyants3 Then
                MsgBox("Toute la population est croyante, en " & lblTime.Text & " périodes !" & " Couleurs Victorieuses : " & cbColor1.Text & ", " & cbColor2.Text)
            ElseIf nbCroyants = nbCroyants3 And nbCroyants > nbCroyants2 Then
                MsgBox("Toute la population est croyante, en " & lblTime.Text & " périodes !" & " Couleurs Victorieuses : " & cbColor1.Text & ", " & cbColor3.Text)
            ElseIf nbCroyants2 = nbCroyants3 And nbCroyants2 > nbCroyants Then
                MsgBox("Toute la population est croyante, en " & lblTime.Text & " périodes !" & " Couleurs Victorieuses : " & cbColor2.Text & ", " & cbColor3.Text)
            End If
            btnLecture.BackColor = Color.Red
        End If
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        txtCroyants.Text = 0
        txtCro2.Text = 0
        txtCro3.Text = 0

        txtIndiv.Text = 0
        lblIndiv.Text = "en attente..."
        lblTotCroy.Text = "en attente..."

        lblCroyants.Text = "en attente..."
        lblCroMoy.Text = "en attente..."
        lblFPMoy.Text = "en attente.."
        lblCreMoy.Text = "en attente.."
        lblPourcentage.Text = 0.00

        lblCroyants2.Text = "en attente..."
        lblCroMoy2.Text = "en attente..."
        lblFPMoy2.Text = "en attente.."
        lblCreMoy2.Text = "en attente.."
        lblPourcentage2.Text = 0.00

        lblCroyants3.Text = "en attente..."
        lblCroMoy3.Text = "en attente..."
        lblFPMoy3.Text = "en attente.."
        lblCreMoy3.Text = "en attente.."
        lblPourcentage3.Text = 0.00

        cbSimuTime.Text = ""
        lblTime.Text = 0
        lblLoading.Text = "Chargement :"
        ProgressBar1.Value = 0
        Dessin.Clear(Color.White)
        pbMap.Image = Feuille

        Me.Chart1.Series("Nombre de croyants").Points.Clear()
        Me.Chart2.Series("Croyance 1").Points.Clear()
        Me.Chart2.Series("Croyance 2").Points.Clear()
        Me.Chart2.Series("Croyance 3").Points.Clear()
    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        txtIndiv.Text = TrackBar1.Value
        TrackBar2.Maximum = txtIndiv.Text - TrackBar3.Value - TrackBar4.Value
        TrackBar3.Maximum = txtIndiv.Text - TrackBar2.Value - TrackBar4.Value
        TrackBar4.Maximum = txtIndiv.Text - TrackBar3.Value - TrackBar2.Value
    End Sub

    Private Sub TrackBar2_Scroll(sender As Object, e As EventArgs) Handles TrackBar2.Scroll
        TrackBar2.Maximum = txtIndiv.Text - TrackBar3.Value - TrackBar4.Value

        txtCroyants.Text = TrackBar2.Value
        If txtIndiv.Text <> "" Then
            TrackBar3.Maximum = txtIndiv.Text - TrackBar2.Value - TrackBar4.Value
            TrackBar4.Maximum = txtIndiv.Text - TrackBar3.Value - TrackBar2.Value
        Else
            MsgBox("Veuillez rentrer un nombre d'individus.")
        End If
    End Sub

    Private Sub TrackBar3_Scroll(sender As Object, e As EventArgs) Handles TrackBar3.Scroll
        TrackBar3.Maximum = txtIndiv.Text - TrackBar2.Value - TrackBar4.Value

        txtCro2.Text = TrackBar3.Value
        If txtIndiv.Text <> "" Then
            TrackBar2.Maximum = txtIndiv.Text - TrackBar3.Value - TrackBar4.Value
            TrackBar4.Maximum = txtIndiv.Text - TrackBar3.Value - TrackBar2.Value
        Else
            MsgBox("Veuillez rentrer un nombre d'individus.")
        End If

    End Sub

    Private Sub TrackBar4_Scroll(sender As Object, e As EventArgs) Handles TrackBar4.Scroll

        TrackBar4.Maximum = txtIndiv.Text - TrackBar3.Value - TrackBar2.Value

        txtCro3.Text = TrackBar4.Value
        If txtIndiv.Text <> "" Then
            TrackBar2.Maximum = txtIndiv.Text - TrackBar3.Value - TrackBar4.Value
            TrackBar3.Maximum = txtIndiv.Text - TrackBar2.Value - TrackBar4.Value
        Else
            MsgBox("Veuillez rentrer un nombre d'individus.")
        End If
    End Sub

    Private Sub txtIndiv_TextChanged(sender As Object, e As EventArgs) Handles txtIndiv.TextChanged
        TrackBar2.Value = TrackBar2.Minimum
        txtCroyants.Text = 0
        TrackBar3.Value = TrackBar3.Minimum
        txtCro2.Text = 0
        TrackBar4.Value = TrackBar4.Minimum
        txtCro3.Text = 0
    End Sub

    Private Sub cbColor1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbColor1.SelectedIndexChanged
        Select Case cbColor1.Text
            Case "Rouge"
                txtCroyants.ForeColor = Color.Red
                gb1.BackColor = Color.Tomato
            Case "Vert"
                txtCroyants.ForeColor = Color.Green
                gb1.BackColor = Color.YellowGreen
            Case "Bleu"
                txtCroyants.ForeColor = Color.Blue
                gb1.BackColor = Color.LightSkyBlue
            Case Else
                txtCroyants.ForeColor = Color.Red 'couleur rouge par défaut
                gb1.BackColor = Color.Tomato
        End Select

    End Sub

    Private Sub cbColor2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbColor2.SelectedIndexChanged
        Select Case cbColor2.Text
            Case "Rouge"
                txtCro2.ForeColor = Color.Red
                gb2.BackColor = Color.Tomato
            Case "Vert"
                txtCro2.ForeColor = Color.Green
                gb2.BackColor = Color.YellowGreen
            Case "Bleu"
                txtCro2.ForeColor = Color.Blue
                gb2.BackColor = Color.LightSkyBlue
            Case Else
                txtCro2.ForeColor = Color.Green 'couleur verte par défaut
                gb2.BackColor = Color.Tomato
        End Select
    End Sub

    Private Sub cbColor3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbColor3.SelectedIndexChanged
        Select Case cbColor3.Text
            Case "Rouge"
                txtCro3.ForeColor = Color.Red
                gb3.BackColor = Color.Tomato
            Case "Vert"
                txtCro3.ForeColor = Color.Green
                gb3.BackColor = Color.YellowGreen
            Case "Bleu"
                txtCro3.ForeColor = Color.Blue
                gb3.BackColor = Color.LightSkyBlue
            Case Else
                txtCro3.ForeColor = Color.Blue 'couleur bleue par défaut
                gb3.BackColor = Color.LightSkyBlue
        End Select
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            Select Case cbColor1.Text
                Case "Rouge"
                    txtCroyants.ForeColor = Color.Red
                    gb1.BackColor = Color.Tomato
                Case "Vert"
                    txtCroyants.ForeColor = Color.Green
                    gb1.BackColor = Color.YellowGreen
                Case "Bleu"
                    txtCroyants.ForeColor = Color.Blue
                    gb1.BackColor = Color.LightSkyBlue
                Case Else
                    txtCroyants.ForeColor = Color.Red 'couleur rouge par défaut
                    gb1.BackColor = Color.Tomato
            End Select
        Else
            gb1.BackColor = Color.DarkGray
        End If
    End Sub

    Private Sub pbMap_MouseMove(sender As Object, e As MouseEventArgs) Handles pbMap.MouseMove
        Label16.Text = "X = " & e.X & " ; Y = " & e.Y
        Dim nbIndiv As Integer = CInt(txtIndiv.Text)
        Dim tempX As Integer = e.X
        Dim tempY As Integer = e.Y

        Dim perso As Perso
        For Each perso In tabIndiv
            For i = -Math.Round(tailleCarre / 2, 0) To Math.Round(tailleCarre / 2, 0)
                If perso.Position.x = tempX + i Then
                    If perso.Position.y = tempY + i Then
                        Label23.Text = "Force de persuasion = " & perso.FP & " ; " & " Niveau de crédulité = " & perso.NCre
                    End If
                End If
            Next
        Next
    End Sub

    Private Sub pbMap_MouseLeave(sender As Object, e As EventArgs) Handles pbMap.MouseLeave
        Label16.Text = "Coordonnées"
        Label23.Text = "Informations sur l'individu"
    End Sub

    Private Sub btnChooseBackColor_Click(sender As Object, e As EventArgs) Handles btnChooseBackColor.Click
        ColorDialog1.ShowDialog()
        Me.BackColor = ColorDialog1.Color
        Me.Chart1.Series("Nombre de croyants").Color = ColorDialog1.Color
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            Select Case cbColor2.Text
                Case "Rouge"
                    gb2.BackColor = Color.Tomato
                Case "Vert"
                    gb2.BackColor = Color.YellowGreen
                Case "Bleu"
                    gb2.BackColor = Color.LightSkyBlue
                Case Else
                    'couleur rouge par défaut
                    gb2.BackColor = Color.YellowGreen
            End Select
        Else
            gb2.BackColor = Color.DarkGray
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked Then
            Select Case cbColor3.Text
                Case "Rouge"
                    gb3.BackColor = Color.Tomato
                Case "Vert"
                    gb3.BackColor = Color.YellowGreen
                Case "Bleu"
                    gb3.BackColor = Color.LightSkyBlue
                Case Else
                    'couleur Bleu par défaut
                    gb3.BackColor = Color.LightSkyBlue
            End Select
        Else
            gb3.BackColor = Color.DarkGray
        End If
    End Sub

End Class
