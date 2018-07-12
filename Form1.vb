Public Class Form1
    Private MyECB As New ECBExchanges
#Region "Form Events"

    ''' <summary>
    ''' Behandelt das Load Event der Form
    ''' </summary>
    ''' <param name="sender">Auslösendes Steuerelement</param>
    ''' <param name="e">System.EventArgs</param>
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Liste der Europäischen Zentralbank aus dem Internet auslesen und
        ' in ListView einfügen
        Call ListeLaden(False)
    End Sub

#End Region

#Region "MenuBar"

    ''' <summary>
    ''' Behandelt das Klick Event des Menupunktes "Lokal Kurse aktualisieren"
    ''' </summary>
    ''' <param name="sender">Auslösendes Steuerelement</param>
    ''' <param name="e">System.EventArgs</param>
    Private Sub LokalAktualisierenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LokalAktualisierenToolStripMenuItem.Click
        ' Liste der Europäischen Zentralbank aus lokaler XML Datei auslesen und
        ' in ListView einfügen
        Call ListeLaden(True)
    End Sub

    ''' <summary>
    ''' Behandelt das Klick Event des Menupunktes "&Internet Kurse aktualisieren"
    ''' </summary>
    ''' <param name="sender">Auslösendes Steuerelement</param>
    ''' <param name="e">System.EventArgs</param>
    Private Sub InternetAktualiToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InternetAktualiToolStripMenuItem.Click
        ' Liste der Europäischen Zentralbank aus dem Internet auslesen und
        ' in ListView einfügen
        Call ListeLaden(False)
    End Sub

    ''' <summary>
    ''' Behandelt das Klick Event des Menupunktes "Beenden"
    ''' </summary>
    ''' <param name="sender">Auslösendes Steuerelement</param>
    ''' <param name="e">System.EventArgs</param>
    Private Sub BeendenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BeendenToolStripMenuItem.Click
        ' Anwendung beenden
        Me.Close()
    End Sub

#End Region

#Region "Funktionen/Methoden"

    ''' <summary>
    ''' Methode zum auslesen der Kursdaten über Klasse und belegen der ListView
    ''' </summary>
    ''' <param name="lokal">Boolsche Variable ob XML Datei aus Internet (=False)
    ''' oder aus einer Lokalen Datei (=True) gelesen werden soll</param>
    Private Sub ListeLaden(Optional ByVal lokal As Boolean = False)

        Dim InetDatei As String = "http://www.ecb.int/stats/eurofxref/eurofxref-daily.xml"
        Dim str As String
        Dim strTmp As String
        Dim subItems() As ListViewItem.ListViewSubItem
        Dim item As ListViewItem = Nothing

        ' Wenn lokal dann kommt FileOpenDialog zum wählen der eurofxref-daily.xml Datei
        If lokal Then
            With OpenFileDialog1
                ' Default Erweiterung
                .DefaultExt = ".xml"
                ' Filter für auswählbare Dateien
                .Filter = "XML Dateien (*.xml)|*.xml"
                ' Anfangsverzeichnis
                .InitialDirectory = "%USERPROFILE%"
                ' Dateiname
                .FileName = "%USERPROFILE%\eurofxref-daily.xml"
                ' FileOpenDialog anzeigen
                If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                    ' Dateiname mit Pfad belegen
                    strTmp = .FileName
                Else
                    ' falls keine Auswahl dann abbrechen
                    Exit Sub
                End If
            End With
        Else
            ' Internet URL belegen
            strTmp = InetDatei
        End If

        ' XML Datei mit den Kursinformationen über Klasse auswerten und in Liste
        ' schreiben
        Dim ECBList As List(Of ECBType) = MyECB.getECBCurrencyExchanges(strTmp)

        ' ListView leeren
        ListView1.Clear()
        ' Wenn Liste nicht leer ist dann ListView belegen
        If ECBList IsNot Nothing Then
            ' Beschriftung der Form mit Datum der Wechselkurse
            Me.Text = "Wechselkurse der EZB vom " & MyECB.Datum

            With ListView1
                ' ListView Darstellung ist Detail
                .View = View.Details
                ' Spalten erzeugen
                .Columns.Add("Euro")
                .Columns.Add("Währung")
                .Columns.Add("Land")

                ' Schleife über alle Listen Items
                For i As Integer = 0 To ECBList.Count - 1
                    If ECBList(i).DisplayName.Length > 0 Then
                        ' Ländername bestimmen
                        str = ECBList(i).DisplayName.Substring(ECBList(i).DisplayName.IndexOf("(") + 1, ECBList(i).DisplayName.Length - (ECBList(i).DisplayName.IndexOf("(") + 2))
                    Else
                        ' Für nicht vom System erkannte Länder manuell Ländername ergänzen
                        Select Case ECBList(i).ECBCurrency.ToUpper.Trim
                            Case "RUB"
                                str = "Russland"
                            Case "BGN"
                                str = "Bulgarien"
                            Case "RON"
                                str = "Rumänien"
                            Case Else
                                str = String.Empty
                        End Select
                    End If

                    ' Neuen ListView Eintrag erstellen
                    item = New ListViewItem("1 Euro =", 0)
                    ' Zweite und dritte Spalte des Eintrags erstellen
                    subItems = New ListViewItem.ListViewSubItem() _
                                    {New ListViewItem.ListViewSubItem(item, _
                                     CDbl(ECBList(i).ECBRate.Replace(".", ",")).ToString("#0.0000") & " " & ECBList(i).ECBCurrency), _
                                     New ListViewItem.ListViewSubItem(item, str)}

                    ' Alle Spalten des Eintrags zusammmenfassen
                    item.SubItems.AddRange(subItems)
                    ' Zeile in ListView aufnehmen
                    .Items.Add(item)
                Next

                ' Bei allen Spalten einen AutoRisize vornehmen 
                .Columns(0).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
                .Columns(1).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
                .Columns(2).AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent)
            End With
        End If

    End Sub

#End Region

End Class
