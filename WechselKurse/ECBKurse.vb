Imports System.Globalization

''' <summary>
''' ECB Struktur
''' </summary>
Public Structure ECBType
    Public ECBCurrency As String
    Public ECBRate As String
    Public DisplayName As String
End Structure

''' <summary>
''' Klasse zum auslesen der XML Datei der 
''' Europäischen Zentralbank mit den Wechselkursen
''' </summary>
Public Class ECBExchanges

#Region "Variablen"

    Private _Datum As String

#End Region

#Region "Eigenschaften"

    ''' <summary>
    ''' ReadOnly Eigenschaft für Datum als String
    ''' </summary>
    ''' <returns>Das ermittelte Datum als String</returns>
    Public ReadOnly Property Datum() As String
        Get
            Return _Datum
        End Get
    End Property

#End Region

#Region "Funktionen/Methoden"

    ''' <summary>
    ''' Funktion um geneische Liste aus XML Datei zu belegen
    ''' </summary>
    ''' <param name="WebAddress">Ort oder URL der XML Datei</param>
    ''' <returns>Generische List der Stucture ECBType</returns>
    Public Function getECBCurrencyExchanges(ByVal WebAddress As String) _
      As List(Of ECBType)

        Try
            Dim xr As XElement = XElement.Load(WebAddress)
            Dim xn As XNamespace = xr.Attribute("xmlns").Value

            ' Linq Abfrage um Liste mit den Werten der XML Datei zu belegen
            Dim xECBs = From ECB In xr.Descendants(xn + "Cube") _
              Where ECB.Attribute("currency") IsNot Nothing _
              AndAlso ECB.Attribute("rate") IsNot Nothing _
              Select New ECBType With { _
                .ECBCurrency = ECB.Attribute("currency").Value, _
                .ECBRate = ECB.Attribute("rate").Value, _
                .DisplayName = CurrencyName(.ECBCurrency)}

            ' Linq Abfrage um Datum aus XML Datei zu belegen
            Dim d = From ECB In xr.Descendants(xn + "Cube") _
              Where ECB.Attribute("time") IsNot Nothing _
              Select _
                Datum = DateTime.Parse(ECB.Attribute("time").Value).ToShortDateString

            _Datum = d(0).ToString

            ' Die Liste 
            Return xECBs.ToList
        Catch ex As Exception
            ' Bei Ausnahme wird Nothing zurückgegeben
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Funktion um Ländernamen aus Währungsabkürzungen zu erhalten
    ''' </summary>
    ''' <param name="isoCode">Abkürzung der Währung</param>
    ''' <returns>Ländernamen als String</returns>
    Private Function CurrencyName(ByVal isoCode As String) As String
        Dim cultures As CultureInfo() = CultureInfo.GetCultures( _
          CultureTypes.SpecificCultures)

        ' Schleife über alle CultureInfos des Systems
        For Each ci As CultureInfo In cultures
            ' RegionInfo abrufen
            Dim ri As New RegionInfo(ci.LCID)
            ' Falls Region Info gleich der übergebenen Kennung ist wird der Name 
            ' zurückgegeben
            If ri.ISOCurrencySymbol = isoCode Then
                Return ci.DisplayName
            End If
        Next

        ' Falls keine CultureInfo gefunden Leerstring zurückgeben
        Return String.Empty
    End Function

#End Region

End Class

