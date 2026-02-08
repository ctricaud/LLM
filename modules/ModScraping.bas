Attribute VB_Name = "ModScraping"

Option Explicit
Sub scrapeExemple()
    '---------------------------------------------------------
    'Connexion sur airbnb.fr
    '---------------------------------------------------------
    '1. On configure la connexion
    '---------------------------------------------------------
    Dim driver As New WebDriver
   
    driver.Start "Edge"
    driver.get "https://www.airbnb.fr"
    driver.Window.Maximize
    'Application.Wait Now + TimeValue("00:00:05")
    WaitReadyState driver
    
    '---------------------------------------------------------
    '2. On met en place le cookie
    '---------------------------------------------------------
    Call driver.Manage.AddCookie("_aat", "0%7CUHTWxsV%2FM%2FM4FAmbtNpbfP9yXFMJ9%2ByOL%2BShfedYq4c%2B%2F9fTYSYUifRRb6U4q3qE", ".airbnb.fr")
    driver.FindElementByXPath("//*[contains(text(), 'Accepter tout')]").Click
    
   '---------------------------------------------------------
    '3. On affiche une réservation
    '---------------------------------------------------------
    driver.get "https://www.airbnb.fr/hosting/reservations/details/HME2CAEKSB"
    WaitReadyState driver
    
    driver.Quit

End Sub



Sub WaitReadyState(driver As WebDriver)
Dim ReadyState
    Do
    ReadyState = driver.ExecuteScript("return document.readyState")
    ' Attendre une courte période avant de vérifier à nouveau
    Application.Wait Now + TimeValue("00:00:01")
Loop While ReadyState <> "complete"
End Sub


