VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEmailCommissionBookingHobe 
   Caption         =   "Calcul Commission Booking Hobe"
   ClientHeight    =   1824
   ClientLeft      =   132
   ClientTop       =   504
   ClientWidth     =   2880
   OleObjectBlob   =   "frmEmailCommissionBookingHobe.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEmailCommissionBookingHobe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub calculCommission(EnvoiEmail)
   '+++++++++++++++++++++++++++++++++++++++++++++++++++++
    '+ On calcule la valeur de la commission Booking
    '+ Et on envoie le mail si la demande est faite
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    '---------------------------------------------------------------------
    '1. Initialisation
    '---------------------------------------------------------------------
    Dim moisSelectionne As Integer               'Le mois sélectionné
    Dim anneeSelectionnee As Integer             'L'année sélectionnée
    Dim ListeLogements As Variant                'La liste des logements
    Dim nombreLogements As Integer               'Le nombre de logements
    Dim listeReservations As Variant             'La liste des réservations
    
    moisSelectionne = SelectionMois.ListIndex + 1
    anneeSelectionnee = SelectionAnnee.ListIndex + 2023
    
    '---------------------------------------------------------------------
    '2. On contrôle si les données sont valables
    '---------------------------------------------------------------------
    If moisSelectionne = 0 Or anneeSelectionnee = 2022 Then Exit Sub
    
    '---------------------------------------------------------------------
    '3. On lance le calcul
    '---------------------------------------------------------------------
    ListeLogements = Range("Logements[Logements]").value
    nombreLogements = UBound(ListeLogements)
    listeReservations = Range("ListeRésas")
    
    Dim calculCommission()
    ReDim calculCommission(nombreLogements, 2)
    Dim totalCommission(3)
    
    'On lance la boucle pour mettre à jour la table calculCommission
    Dim i, j As Integer
    
    For i = 1 To UBound(listeReservations)
        If listeReservations(i, 2) = "Booking" Then
            If Month(listeReservations(i, 3)) = moisSelectionne And Year(listeReservations(i, 3)) = anneeSelectionnee Then
                'On trouve le bon appartement
                For j = 1 To nombreLogements
                    If Range("Logements[Logements]")(j).value = listeReservations(i, 1) Then
                        calculCommission(j, 1) = calculCommission(j, 1) + listeReservations(i, 7)
                        calculCommission(j, 2) = calculCommission(j, 2) + listeReservations(i, 9) - listeReservations(i, 7)
                        totalCommission(1) = totalCommission(1) + listeReservations(i, 7)
                        totalCommission(2) = totalCommission(2) + listeReservations(i, 9) - listeReservations(i, 7)
                        Exit For
                    End If
                Next j
            End If
            
        End If
    Next i
    
    
    '---------------------------------------------------------------------
    '4. On affiche la commission
    '---------------------------------------------------------------------
    Commission.Caption = " " + FormatNumber(totalCommission(1) + totalCommission(2), 2, vbTrue, vbTrue, vbTrue) + " €"
    
    '--------------------------------------------------------------------------
    '5. On envoie l'email si la demande est faite
    '--------------------------------------------------------------------------
    
    If EnvoiEmail Then
        frmEmailCommissionBookingHobe.Hide
        
        'On construit le message
        Dim message As String
        
        message = "Cher Louis," + vbCrLf + vbCrLf + _
                  "Comme chaque mois, je te propose de trouver ci-joint le détail de la facturation de tes prestations pour le mois de "
        message = message + SelectionMois.Text + " " + SelectionAnnee.Text + " :" + vbCrLf + vbCrLf
        message = message + "Le total facturé est de " + FormatNumber(totalCommission(1) + totalCommission(2), 2, vbTrue, vbTrue, vbTrue) + " € qui se décompose comme suit :" + vbCrLf + vbCrLf
            
        message = message + "Montant des ménages = " + FormatNumber(totalCommission(1), 2, vbTrue, vbTrue, vbTrue) + " € dont :" + vbCrLf
        For i = 1 To nombreLogements
            If calculCommission(i, 1) > 0 Then
                message = message + "- " + Range("Logements[Logements]")(i).value + " - " + FormatNumber(calculCommission(i, 1), 2, vbTrue, vbTrue, vbTrue) + " €" + vbCrLf
            End If
        Next i
        
        message = message + vbCrLf + "Montant de la commission = " + FormatNumber(totalCommission(2), 2, vbTrue, vbTrue, vbTrue) + " € dont :" + vbCrLf
        For i = 1 To nombreLogements
            If calculCommission(i, 2) > 0 Then
                message = message + "- " + Range("Logements[Logements]")(i).value + " - " + FormatNumber(calculCommission(i, 2), 2, vbTrue, vbTrue, vbTrue) + " €" + vbCrLf
            End If
        Next i
        
        message = message + vbCrLf + "Je reste à ta disposition pour toutes informations complémentaires."
        
        'On envoie le message
        Dim OutApp As Object
        Dim OutMail As Object
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
        With OutMail
            .To = "booking@tricaud.com"
            .Subject = "Relevé des commissions mensuelles Booking de " + SelectionMois.Text + " " + SelectionAnnee.Text
            .body = message
            .send                                    'envoie directement le mail
        End With
        Set OutMail = Nothing
        Set OutApp = Nothing
        'On met un message de confirmation
        
        MsgBox "Le mail a bien été envoyé."
    End If
    
End Sub



Private Sub btAnnuler_Click()
    Me.Hide
End Sub


Private Sub SelectionAnnee_Change()
    calculCommission False
End Sub


Private Sub SelectionMois_Change()
    calculCommission False
End Sub


 Private Sub Valider_Click()
    calculCommission True
End Sub

Private Sub UserForm_Activate()
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++
    '+ On met à jour les champs possible pour sélectionner
    '+ le mois qui nous intéresse
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    SelectionMois.Clear
    SelectionMois.List = Array("Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novemebre", "Décembre")
    
    SelectionMois.ListIndex = Month(Now) - 1

    Dim i As Integer
    For i = 2023 To Year(Now) + 1
        SelectionAnnee.AddItem CStr(i)
    Next

    SelectionAnnee.ListIndex = Year(Now) - 2023

End Sub


