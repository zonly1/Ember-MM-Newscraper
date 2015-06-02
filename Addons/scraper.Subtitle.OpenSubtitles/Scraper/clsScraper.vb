' ################################################################################
' #                             EMBER MEDIA MANAGER                              #
' ################################################################################
' ################################################################################
' # This file is part of Ember Media Manager.                                    #
' #                                                                              #
' # Ember Media Manager is free software: you can redistribute it and/or modify  #
' # it under the terms of the GNU General Public License as published by         #
' # the Free Software Foundation, either version 3 of the License, or            #
' # (at your option) any later version.                                          #
' #                                                                              #
' # Ember Media Manager is distributed in the hope that it will be useful,       #
' # but WITHOUT ANY WARRANTY; without even the implied warranty of               #
' # MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the                #
' # GNU General Public License for more details.                                 #
' #                                                                              #
' # You should have received a copy of the GNU General Public License            #
' # along with Ember Media Manager.  If not, see <http://www.gnu.org/licenses/>. #
' ################################################################################

Imports EmberAPI
Imports NLog
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

Namespace OpenSubtitle

    Public Class Scraper

#Region "Fields"

        Shared logger As Logger = NLog.LogManager.GetCurrentClassLogger()
        Private _MySettings As OpenSubtitle.Scraper.sMySettings_ForScraper

#End Region 'Fields

#Region "Methods"

        Public Function GetSubtitles(ByVal strFilename As String, ByVal imdbID As String, ByRef Settings As sMySettings_ForScraper) As List(Of MediaContainers.Subtitle)
            Dim alSubtitles As New List(Of MediaContainers.Subtitle) 'main subtitle list
            Dim alSubtitlesP As New List(Of MediaContainers.Subtitle) 'preferred language subtitle list
            Dim alSubtitlesE As New List(Of MediaContainers.Subtitle) 'english subtitle list

            Dim OSDBClient As OSDBnet.IAnonymousClient = OSDBnet.Osdb.Login("OSTestUserAgent")
            Dim Results As IList(Of OSDBnet.Subtitle) = OSDBClient.SearchSubtitlesFromImdb("ger", imdbID)

            For Each subtitle In Results
                alSubtitles.Add(New MediaContainers.Subtitle With { _
                                .LongLang = subtitle.LanguageName, _
                                .Title = subtitle.SubtitleFileName, _
                                .URL = subtitle.SubTitleDownloadLink.AbsoluteUri})
            Next

            'Subtitles sorting
            'For Each xPoster As MediaContainers.Image In alPostersO.OrderBy(Function(p) (p.LongLang))
            '    alPostersOs.Add(xPoster)
            'Next
            'alPosters.AddRange(alPostersP)
            'alPosters.AddRange(alPostersE)
            'alPosters.AddRange(alPostersOs)
            'alPosters.AddRange(alPostersN)

            Return alSubtitles
        End Function

        Private Function ComputeMovieHash(ByVal filename As String) As Byte()
            Dim result As Byte()
            Using input As Stream = File.OpenRead(filename)
                result = ComputeMovieHash(input)
            End Using
            Return result
        End Function

        Private Function ComputeMovieHash(ByVal input As Stream) As Byte()
            Dim lhash As System.Int64, streamsize As Long
            streamsize = input.Length
            lhash = streamsize

            Dim i As Long = 0
            Dim buffer As Byte() = New Byte(Marshal.SizeOf(GetType(Long)) - 1) {}
            While i < 65536 / Marshal.SizeOf(GetType(Long)) AndAlso (input.Read(buffer, 0, Marshal.SizeOf(GetType(Long))) > 0)
                i += 1

                lhash += BitConverter.ToInt64(buffer, 0)
            End While

            input.Position = Math.Max(0, streamsize - 65536)
            i = 0
            While i < 65536 / Marshal.SizeOf(GetType(Long)) AndAlso (input.Read(buffer, 0, Marshal.SizeOf(GetType(Long))) > 0)
                i += 1
                lhash += BitConverter.ToInt64(buffer, 0)
            End While
            input.Close()
            Dim result As Byte() = BitConverter.GetBytes(lhash)
            Array.Reverse(result)
            Return result
        End Function

        Private Function ToHexadecimal(ByVal bytes As Byte()) As String
            Dim hexBuilder As New StringBuilder()
            For i As Integer = 0 To bytes.Length - 1
                hexBuilder.Append(bytes(i).ToString("x2"))
            Next
            Return hexBuilder.ToString()
        End Function

        Private Sub Main(ByVal strFilename As String)
            If File.Exists(strFilename) Then
                Dim moviehash As Byte() = ComputeMovieHash(strFilename)
                Console.WriteLine("The hash of the movie-file is: {0}", ToHexadecimal(moviehash))
            End If
        End Sub

#End Region 'Methods

#Region "Nested Types"

        Structure sMySettings_ForScraper

#Region "Fields"

            Dim APIKey As String
            Dim PrefLanguage As String
            Dim PrefLanguageOnly As Boolean
            Dim GetEnglishImages As Boolean
            Dim GetBlankImages As Boolean

#End Region 'Fields

        End Structure

#End Region 'Nested Types

    End Class

End Namespace