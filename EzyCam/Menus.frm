VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3480
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5805
   LinkTopic       =   "Form2"
   ScaleHeight     =   3480
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu world 
      Caption         =   "World"
      Begin VB.Menu canada 
         Caption         =   "Canada"
         Begin VB.Menu niagara 
            Caption         =   "Niagara Falls"
         End
         Begin VB.Menu white 
            Caption         =   "White Rock Beach"
         End
      End
      Begin VB.Menu samerica 
         Caption         =   "South America"
         Begin VB.Menu brazil 
            Caption         =   "Brazil"
            Begin VB.Menu brasilia 
               Caption         =   "Brasilia"
            End
            Begin VB.Menu sao 
               Caption         =   "Sao Paulo"
            End
         End
         Begin VB.Menu argentina 
            Caption         =   "Argentina"
            Begin VB.Menu buenosaires 
               Caption         =   "Buenos Aires"
            End
            Begin VB.Menu puerto 
               Caption         =   "Puerto Madero"
            End
         End
      End
   End
   Begin VB.Menu animals 
      Caption         =   "Animals"
      Begin VB.Menu ants 
         Caption         =   "Ants Cam"
      End
      Begin VB.Menu bat 
         Caption         =   "Bat Cam"
      End
      Begin VB.Menu puppy 
         Caption         =   "Puppy Cam"
      End
      Begin VB.Menu seal 
         Caption         =   "Seal Cam"
      End
   End
   Begin VB.Menu people 
      Caption         =   "People"
      Begin VB.Menu arts 
         Caption         =   "Arts"
         Begin VB.Menu rushmore 
            Caption         =   "Mount Rushmore"
         End
         Begin VB.Menu corvette 
            Caption         =   "National Corvette Museum"
         End
      End
      Begin VB.Menu education 
         Caption         =   "Education"
         Begin VB.Menu alaskauniv 
            Caption         =   "Alaska Pacific University"
         End
         Begin VB.Menu rhode 
            Caption         =   "Rhode Islands College"
         End
         Begin VB.Menu hess 
            Caption         =   "Hess Hall"
         End
      End
      Begin VB.Menu Sports 
         Caption         =   "Sports"
         Begin VB.Menu body 
            Caption         =   "Body Cam"
         End
         Begin VB.Menu lake 
            Caption         =   "Lake Louise (ski)"
         End
         Begin VB.Menu ohio 
            Caption         =   "Ohio Stadium"
         End
      End
   End
   Begin VB.Menu space 
      Caption         =   "Space"
      Begin VB.Menu earthspace 
         Caption         =   "Earth from Space"
      End
      Begin VB.Menu robo 
         Caption         =   "RoboSky (nasa)"
      End
      Begin VB.Menu suncam 
         Caption         =   "Sun cam Doppler (nasa)"
      End
   End
   Begin VB.Menu transport 
      Caption         =   "Transport"
      Begin VB.Menu airports 
         Caption         =   "Airports"
         Begin VB.Menu geneve 
            Caption         =   "Geneve Airport Cointrin"
         End
         Begin VB.Menu salzburg 
            Caption         =   "Salzburg Airport"
         End
         Begin VB.Menu stockholms 
            Caption         =   "Stockholms Bromma Airport"
         End
      End
      Begin VB.Menu taxi 
         Caption         =   "Taxi cams"
         Begin VB.Menu cabcam 
            Caption         =   "Cab Cam (NY)"
         End
         Begin VB.Menu thetaxicam 
            Caption         =   "The taxi cam"
         End
      End
   End
   Begin VB.Menu weather 
      Caption         =   "Weather"
      Begin VB.Menu moonphase 
         Caption         =   "Current Moon Phase"
      End
      Begin VB.Menu daylightzone 
         Caption         =   "Current Daylight zone"
      End
   End
   Begin VB.Menu financial 
      Caption         =   "Financial"
      Begin VB.Menu down 
         Caption         =   "Down Jones"
      End
      Begin VB.Menu nasdaq 
         Caption         =   "Nasdaq"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub alaskauniv_Click()
Form1.Text1.Text = "http://apu1.alaskapacific.edu/webcam/webcam.jpg"
showcam
End Sub

Private Sub ants_Click()
Form1.Text1.Text = "http://www.discovery.com/cams/ant/images/ant.jpg?ct=3906c830"
showcam
End Sub

Private Sub bat_Click()
Form1.Text1.Text = "http://www.discovery.com/cams/bat/images/bat.jpg?ct=3916cbb7"
showcam
End Sub

Private Sub body_Click()
Form1.Text1.Text = "http://camimg.discovery.com/people/body/body.jpg?ct=3a1bf093"
showcam
End Sub

Private Sub brasilia_Click()
Form1.Text1.Text = "http://www.uol.com.br/aliwebcam2/brasilia.jpg"
showcam
End Sub

Private Sub buenosaires_Click()
Form1.Text1.Text = "http://itaucam.itau.com.ar/fullsize.jpg"
showcam
End Sub

Private Sub cabcam_Click()
Form1.Text1.Text = "http://www.ny-taxi.com/nycc/archive/taxi.jpg?0.7404138"
showcam
End Sub

Private Sub corvette_Click()
Form1.Text1.Text = "http://www2.corvettemuseum.com/webcam/images/tv_image5.JPG"
showcam
End Sub

Private Sub daylightzone_Click()
Form1.Text1.Text = "http://www.bsdi.com/icons/misc/datetime.gif"
showcam
End Sub

Private Sub down_Click()
Form1.Text1.Text = "http://chart.neural.com/scripts/chart/dllchart.dll?MfcISAPICommand=Chart&sym1=$indu&dres=min&size=325x180&cbckg=FFFFFF&csym1=blk&csym2=red&csym3=dgrn&csym4=blu&csym5=mag&cbcku=ffffcc&cbckl=ffffcc&cbckd=ffffcc&ctxtu=blu&ctxtd=blu&ctxtl=blu&cind1a=red&cind2=blk&cind3=red&cind3a=lblu&cind4=7200AA&cind5=AA681A&cind6=mag&cavg1=lred&cavg2=cyn&height=180&width=325&source=WALLST&ignore=1220001642"
showcam
End Sub

Private Sub earthspace_Click()
Form1.Text1.Text = "http://www.discovery.com/cgi-bin/nasa/worldmap.pl/littlesea.gif"
showcam
End Sub


Private Sub geneve_Click()
Form1.Text1.Text = "http://194.158.29.50/fullsize.jpg"
showcam
End Sub

Private Sub hess_Click()
Form1.Text1.Text = "http://128.169.144.124/17th.jpg"
showcam
End Sub

Private Sub lake_Click()
Form1.Text1.Text = "http://web.alberta.com/skycams/PROGS/webcam.cgi?picture=louise"
showcam
End Sub

Private Sub moonphase_Click()
Form1.Text1.Text = "http://tycho.usno.navy.mil/cgi-bin/phase.gif"
showcam
End Sub

Private Sub nasdaq_Click()
Form1.Text1.Text = "http://chart.neural.com/scripts/chart/dllchart.dll?MfcISAPICommand=Chart&sym1=$compq&dres=min&size=325x180&cbckg=FFFFFF&csym1=blk&csym2=red&csym3=dgrn&csym4=blu&csym5=mag&cbcku=ffffcc&cbckl=ffffcc&cbckd=ffffcc&ctxtu=blu&ctxtd=blu&ctxtl=blu&cind1a=red&cind2=blk&cind3=red&cind3a=lblu&cind4=7200AA&cind5=AA681A&cind6=mag&cavg1=lred&cavg2=cyn&height=180&width=325&source=WALLST&ignore=1220001646"
showcam
End Sub

Private Sub niagara_Click()
Form1.Text1.Text = "http://www.computan.on.ca/corp/fallsview/fallsmain.jpg"
showcam
End Sub

Private Sub showcam()
If Form1.Combo1.Text = "10 seconds" Then
   Form1.Timer1.Interval = 10000
End If
If Form1.Combo1.Text = "30 seconds" Then
   Form1.Timer1.Interval = 30000
End If
If Form1.Combo1.Text = "1 minute" Then
   Form1.Timer1.Interval = 100000
End If
If Form1.Combo1.Text = "5 minutes" Then
   Form1.Timer1.Interval = 500000
End If
If Form1.Combo1.Text = "10 minutes" Then
   Form1.Timer1.Interval = 1000000
End If
If Form1.Combo1.Text = "20 minutes" Then
   Form1.Timer1.Interval = 2000000
End If
If Form1.Combo1.Text = "30 minutes" Then
   Form1.Timer1.Interval = 3000000
End If
Form1.Timer1.Enabled = True
Form1.WebBrowser1.Navigate (Form1.Text1.Text)
Form1.Label3.Caption = "Connecting..."
Form1.Label5.Caption = "On"
Form1.Shape4.FillColor = vbGreen
End Sub

Private Sub ohio_Click()
Form1.Text1.Text = "http://www.ps.ohio-state.edu/webcams/webcam2.jpg"
showcam
End Sub

Private Sub puerto_Click()
Form1.Text1.Text = "http://www.lanacion.com.ar/camara/clarita.jpg"
showcam
End Sub

Private Sub puppy_Click()
Form1.Text1.Text = "http://www.discovery.com/cams/puppy/images/puppy.jpg?ct=3916cbda"
showcam
End Sub

Private Sub rhode_Click()
Form1.Text1.Text = "http://manncam1.ric.edu/fullsize.jpg"
showcam
End Sub

Private Sub robo_Click()
Form1.Text1.Text = "http://www.robosky.com/widefield.jpg"
showcam
End Sub

Private Sub rushmore_Click()
Form1.Text1.Text = "http://lightning.state.sd.us/webcam/rushmore.jpg"
showcam
End Sub

Private Sub salzburg_Click()
Form1.Text1.Text = "http://www.salzburg.com/airport/bilder/fullsize.jpg"
showcam
End Sub

Private Sub sao_Click()
Form1.Text1.Text = "http://www.uol.com.br/aliwebcam2/avpaulista.jpg"
showcam
End Sub

Private Sub seal_Click()
Form1.Text1.Text = "http://www.discovery.com/cams/seal/images/seal.jpg?ct=3916ceb6"
showcam
End Sub

Private Sub stockholms_Click()
Form1.Text1.Text = "http://webcam.connection.se/flygtorget/live_bromma.jpg"
showcam
End Sub

Private Sub suncam_Click()
Form1.Text1.Text = "http://camimg.discovery.com/space/soho/dopsun.gif?ct=382873d4"
showcam
End Sub

Private Sub thetaxicam_Click()
Form1.Text1.Text = "http://www.ultimatetaxi.com/cam/taxi.jpg"
showcam
End Sub

Private Sub white_Click()
Form1.Text1.Text = "http://24.113.75.182/cgi-bin/fullsize.jpg"
showcam
End Sub
