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

Imports System.IO
Imports EmberAPI
Imports ScraperModule.RottenTomatoes
Imports NLog
Imports System.Diagnostics

Public Class RottenTomatoes_Data
    Implements Interfaces.ScraperModule_Data_Movie


#Region "Fields"

    Shared logger As Logger = NLog.LogManager.GetCurrentClassLogger()

    Public Shared ConfigOptions_Movie As New Structures.ScrapeOptions_Movie
    Public Shared ConfigScrapeModifier_Movie As New Structures.ScrapeModifier_Movie_MovieSet
    Public Shared _AssemblyName As String

    Private TMDBId As String

    ''' <summary>
    ''' Scraping Here
    ''' </summary>
    ''' <remarks></remarks>
    Private strPrivateAPIKey As String = String.Empty
    Private _MySettings_Movie As New sMySettings
    Private _Name As String = "RottenTomatoes_Data"
    Private _ScraperEnabled_Movie As Boolean = False
    Private _setup_Movie As frmSettingsHolder_Movie

#End Region 'Fields

#Region "Events"

    Public Event ModuleSettingsChanged_Movie() Implements Interfaces.ScraperModule_Data_Movie.ModuleSettingsChanged
    Public Event ScraperEvent_Movie(ByVal eType As Enums.ScraperEventType_Movie, ByVal Parameter As Object) Implements Interfaces.ScraperModule_Data_Movie.ScraperEvent
    Public Event ScraperSetupChanged_Movie(ByVal name As String, ByVal State As Boolean, ByVal difforder As Integer) Implements Interfaces.ScraperModule_Data_Movie.ScraperSetupChanged
    Public Event SetupNeedsRestart_Movie() Implements Interfaces.ScraperModule_Data_Movie.SetupNeedsRestart

#End Region 'Events

#Region "Properties"

    ReadOnly Property ModuleName() As String Implements Interfaces.ScraperModule_Data_Movie.ModuleName
        Get
            Return _Name
        End Get
    End Property

    ReadOnly Property ModuleVersion() As String Implements Interfaces.ScraperModule_Data_Movie.ModuleVersion
        Get
            Return System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileVersion.ToString
        End Get
    End Property

    Property ScraperEnabled_Movie() As Boolean Implements Interfaces.ScraperModule_Data_Movie.ScraperEnabled
        Get
            Return _ScraperEnabled_Movie
        End Get
        Set(ByVal value As Boolean)
            _ScraperEnabled_Movie = value
        End Set
    End Property

#End Region 'Properties

#Region "Methods"

    Private Sub Handle_ModuleSettingsChanged_Movie()
        RaiseEvent ModuleSettingsChanged_Movie()
    End Sub

    Private Sub Handle_SetupNeedsRestart_Movie()
        RaiseEvent SetupNeedsRestart_Movie()
    End Sub

    Private Sub Handle_SetupScraperChanged_Movie(ByVal state As Boolean, ByVal difforder As Integer)
        ScraperEnabled_Movie = state
        RaiseEvent ScraperSetupChanged_Movie(String.Concat(Me._Name, "_Movie"), state, difforder)
    End Sub

    Sub Init_Movie(ByVal sAssemblyName As String) Implements Interfaces.ScraperModule_Data_Movie.Init
        _AssemblyName = sAssemblyName
        LoadSettings_Movie()
    End Sub

    Function InjectSetupScraper_Movie() As Containers.SettingsPanel Implements Interfaces.ScraperModule_Data_Movie.InjectSetupScraper
        Dim SPanel As New Containers.SettingsPanel
        _setup_Movie = New frmSettingsHolder_Movie
        LoadSettings_Movie()
        _setup_Movie.API = _setup_Movie.txtApiKey.Text
        _setup_Movie.Lang = _setup_Movie.cbPrefLanguage.Text
        _setup_Movie.chkEnabled.Checked = _ScraperEnabled_Movie
        _setup_Movie.cbPrefLanguage.Text = _MySettings_Movie.PrefLanguage
        _setup_Movie.chkCast.Checked = ConfigOptions_Movie.bCast
        _setup_Movie.chkCollectionID.Checked = ConfigOptions_Movie.bCollectionID
        _setup_Movie.chkCountry.Checked = ConfigOptions_Movie.bCountry
        _setup_Movie.chkDirector.Checked = ConfigOptions_Movie.bDirector
        _setup_Movie.chkFallBackEng.Checked = _MySettings_Movie.FallBackEng
        _setup_Movie.chkGenre.Checked = ConfigOptions_Movie.bGenre
        _setup_Movie.chkGetAdultItems.Checked = _MySettings_Movie.GetAdultItems
        _setup_Movie.chkCertification.Checked = ConfigOptions_Movie.bMPAA
        _setup_Movie.chkOriginalTitle.Checked = ConfigOptions_Movie.bOriginalTitle
        _setup_Movie.chkPlot.Checked = ConfigOptions_Movie.bPlot
        _setup_Movie.chkRating.Checked = ConfigOptions_Movie.bRating
        _setup_Movie.chkRelease.Checked = ConfigOptions_Movie.bRelease
        _setup_Movie.chkRuntime.Checked = ConfigOptions_Movie.bRuntime
        _setup_Movie.chkStudio.Checked = ConfigOptions_Movie.bStudio
        _setup_Movie.chkTagline.Checked = ConfigOptions_Movie.bTagline
        _setup_Movie.chkTitle.Checked = ConfigOptions_Movie.bTitle
        _setup_Movie.chkTrailer.Checked = ConfigOptions_Movie.bTrailer
        _setup_Movie.chkVotes.Checked = ConfigOptions_Movie.bVotes
        _setup_Movie.chkWriters.Checked = ConfigOptions_Movie.bWriters
        _setup_Movie.chkYear.Checked = ConfigOptions_Movie.bYear
        _setup_Movie.txtApiKey.Text = strPrivateAPIKey

        If Not String.IsNullOrEmpty(strPrivateAPIKey) Then
            _setup_Movie.btnUnlockAPI.Text = Master.eLang.GetString(443, "Use embedded API Key")
            _setup_Movie.lblEMMAPI.Visible = False
            _setup_Movie.txtApiKey.Enabled = True
        End If

        _setup_Movie.orderChanged()

        SPanel.Name = String.Concat(Me._Name, "_Movie")
        SPanel.Text = "RottenTomatoes"
        SPanel.Prefix = "RottenTomatoesMovieInfo_"
        SPanel.Order = 110
        SPanel.Parent = "pnlMovieData"
        SPanel.Type = Master.eLang.GetString(36, "Movies")
        SPanel.ImageIndex = If(_ScraperEnabled_Movie, 9, 10)
        SPanel.Panel = _setup_Movie.pnlSettings

        AddHandler _setup_Movie.SetupScraperChanged, AddressOf Handle_SetupScraperChanged_Movie
        AddHandler _setup_Movie.ModuleSettingsChanged, AddressOf Handle_ModuleSettingsChanged_Movie
        AddHandler _setup_Movie.SetupNeedsRestart, AddressOf Handle_SetupNeedsRestart_Movie
        Return SPanel
    End Function

    Sub LoadSettings_Movie()
        ConfigOptions_Movie.bCast = clsAdvancedSettings.GetBooleanSetting("DoCast", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bCert = clsAdvancedSettings.GetBooleanSetting("DoCert", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bCollectionID = clsAdvancedSettings.GetBooleanSetting("DoCollectionID", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bCountry = clsAdvancedSettings.GetBooleanSetting("DoCountry", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bDirector = clsAdvancedSettings.GetBooleanSetting("DoDirector", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bFullCrew = clsAdvancedSettings.GetBooleanSetting("DoFullCrews", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bFullCrew = clsAdvancedSettings.GetBooleanSetting("FullCrew", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bGenre = clsAdvancedSettings.GetBooleanSetting("DoGenres", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bMPAA = clsAdvancedSettings.GetBooleanSetting("DoMPAA", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bOriginalTitle = clsAdvancedSettings.GetBooleanSetting("DoOriginalTitle", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bMusicBy = clsAdvancedSettings.GetBooleanSetting("DoMusic", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bOtherCrew = clsAdvancedSettings.GetBooleanSetting("DoOtherCrews", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bOutline = clsAdvancedSettings.GetBooleanSetting("DoOutline", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bPlot = clsAdvancedSettings.GetBooleanSetting("DoPlot", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bProducers = clsAdvancedSettings.GetBooleanSetting("DoProducers", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bRating = clsAdvancedSettings.GetBooleanSetting("DoRating", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bRelease = clsAdvancedSettings.GetBooleanSetting("DoRelease", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bRuntime = clsAdvancedSettings.GetBooleanSetting("DoRuntime", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bStudio = clsAdvancedSettings.GetBooleanSetting("DoStudio", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bTagline = clsAdvancedSettings.GetBooleanSetting("DoTagline", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bTitle = clsAdvancedSettings.GetBooleanSetting("DoTitle", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bTop250 = clsAdvancedSettings.GetBooleanSetting("DoTop250", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bTrailer = clsAdvancedSettings.GetBooleanSetting("DoTrailer", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bVotes = clsAdvancedSettings.GetBooleanSetting("DoVotes", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bWriters = clsAdvancedSettings.GetBooleanSetting("DoWriters", True, , Enums.Content_Type.Movie)
        ConfigOptions_Movie.bYear = clsAdvancedSettings.GetBooleanSetting("DoYear", True, , Enums.Content_Type.Movie)

        strPrivateAPIKey = clsAdvancedSettings.GetSetting("APIKey", "", , Enums.Content_Type.Movie)
        _MySettings_Movie.FallBackEng = clsAdvancedSettings.GetBooleanSetting("FallBackEn", False, , Enums.Content_Type.Movie)
        _MySettings_Movie.GetAdultItems = clsAdvancedSettings.GetBooleanSetting("GetAdultItems", False, , Enums.Content_Type.Movie)
        _MySettings_Movie.APIKey = If(String.IsNullOrEmpty(strPrivateAPIKey), "44810eefccd9cb1fa1d57e7b0d67b08d", strPrivateAPIKey)
        _MySettings_Movie.PrefLanguage = clsAdvancedSettings.GetSetting("PrefLanguage", "en", , Enums.Content_Type.Movie)
        ConfigScrapeModifier_Movie.DoSearch = True
        ConfigScrapeModifier_Movie.Meta = True
        ConfigScrapeModifier_Movie.NFO = True
    End Sub

    Sub SaveSettings_Movie()
        Using settings = New clsAdvancedSettings()
            settings.SetBooleanSetting("DoCast", ConfigOptions_Movie.bCast, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoCert", ConfigOptions_Movie.bCert, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoCollectionID", ConfigOptions_Movie.bCollectionID, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoCountry", ConfigOptions_Movie.bCountry, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoDirector", ConfigOptions_Movie.bDirector, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoFanart", ConfigScrapeModifier_Movie.Fanart, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoFullCrews", ConfigOptions_Movie.bFullCrew, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoGenres", ConfigOptions_Movie.bGenre, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoMPAA", ConfigOptions_Movie.bMPAA, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoOriginalTitle", ConfigOptions_Movie.bOriginalTitle, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoMusic", ConfigOptions_Movie.bMusicBy, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoOtherCrews", ConfigOptions_Movie.bOtherCrew, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoOutline", ConfigOptions_Movie.bOutline, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoPlot", ConfigOptions_Movie.bPlot, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoPoster", ConfigScrapeModifier_Movie.Poster, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoProducers", ConfigOptions_Movie.bProducers, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoRating", ConfigOptions_Movie.bRating, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoRelease", ConfigOptions_Movie.bRelease, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoRuntime", ConfigOptions_Movie.bRuntime, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoStudio", ConfigOptions_Movie.bStudio, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoTagline", ConfigOptions_Movie.bTagline, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoTitle", ConfigOptions_Movie.bTitle, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoTop250", ConfigOptions_Movie.bTop250, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoTrailer", ConfigOptions_Movie.bTrailer, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoVotes", ConfigOptions_Movie.bVotes, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoWriters", ConfigOptions_Movie.bWriters, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoYear", ConfigOptions_Movie.bYear, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("FallBackEn", _MySettings_Movie.FallBackEng, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("FullCrew", ConfigOptions_Movie.bFullCrew, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("GetAdultItems", _MySettings_Movie.GetAdultItems, , , Enums.Content_Type.Movie)
            settings.SetSetting("APIKey", _setup_Movie.txtApiKey.Text, , , Enums.Content_Type.Movie)
            settings.SetSetting("PrefLanguage", _MySettings_Movie.PrefLanguage, , , Enums.Content_Type.Movie)
        End Using
    End Sub

    Sub SaveSetupScraper_Movie(ByVal DoDispose As Boolean) Implements Interfaces.ScraperModule_Data_Movie.SaveSetupScraper
        ConfigOptions_Movie.bCast = _setup_Movie.chkCast.Checked
        ConfigOptions_Movie.bCert = _setup_Movie.chkCertification.Checked
        ConfigOptions_Movie.bCollectionID = _setup_Movie.chkCollectionID.Checked
        ConfigOptions_Movie.bCountry = _setup_Movie.chkCountry.Checked
        ConfigOptions_Movie.bDirector = _setup_Movie.chkDirector.Checked
        ConfigOptions_Movie.bFullCrew = True
        ConfigOptions_Movie.bGenre = _setup_Movie.chkGenre.Checked
        ConfigOptions_Movie.bMPAA = _setup_Movie.chkCertification.Checked
        ConfigOptions_Movie.bOriginalTitle = _setup_Movie.chkOriginalTitle.Checked
        ConfigOptions_Movie.bMusicBy = False
        ConfigOptions_Movie.bOtherCrew = False
        ConfigOptions_Movie.bOutline = _setup_Movie.chkPlot.Checked
        ConfigOptions_Movie.bPlot = _setup_Movie.chkPlot.Checked
        ConfigOptions_Movie.bProducers = _setup_Movie.chkDirector.Checked
        ConfigOptions_Movie.bRating = _setup_Movie.chkRating.Checked
        ConfigOptions_Movie.bRelease = _setup_Movie.chkRelease.Checked
        ConfigOptions_Movie.bRuntime = _setup_Movie.chkRuntime.Checked
        ConfigOptions_Movie.bStudio = _setup_Movie.chkStudio.Checked
        ConfigOptions_Movie.bTagline = _setup_Movie.chkTagline.Checked
        ConfigOptions_Movie.bTitle = _setup_Movie.chkTitle.Checked
        ConfigOptions_Movie.bTop250 = False
        ConfigOptions_Movie.bTrailer = _setup_Movie.chkTrailer.Checked
        ConfigOptions_Movie.bVotes = _setup_Movie.chkVotes.Checked
        ConfigOptions_Movie.bWriters = _setup_Movie.chkWriters.Checked
        ConfigOptions_Movie.bYear = _setup_Movie.chkYear.Checked
        _MySettings_Movie.FallBackEng = _setup_Movie.chkFallBackEng.Checked
        _MySettings_Movie.GetAdultItems = _setup_Movie.chkGetAdultItems.Checked
        _MySettings_Movie.PrefLanguage = _setup_Movie.cbPrefLanguage.Text
        SaveSettings_Movie()
        If DoDispose Then
            RemoveHandler _setup_Movie.SetupScraperChanged, AddressOf Handle_SetupScraperChanged_Movie
            RemoveHandler _setup_Movie.ModuleSettingsChanged, AddressOf Handle_ModuleSettingsChanged_Movie
            _setup_Movie.Dispose()
        End If
    End Sub

    Function GetMovieStudio(ByRef DBMovie As Structures.DBMovie, ByRef sStudio As List(Of String)) As Interfaces.ModuleResult Implements Interfaces.ScraperModule_Data_Movie.GetMovieStudio
        Return New Interfaces.ModuleResult With {.breakChain = False}
    End Function

    Function GetTMDBID(ByVal sIMDBID As String, ByRef sTMDBID As String) As Interfaces.ModuleResult Implements Interfaces.ScraperModule_Data_Movie.GetTMDBID
        Return New Interfaces.ModuleResult With {.breakChain = False}
    End Function
    ''' <summary>
    '''  Scrape MovieDetails from RottenTomatoes
    ''' </summary>
    ''' <param name="DBMovie">Movie to be scraped. DBMovie as ByRef to use existing data for identifing movie and to fill with IMDB/TMDB ID for next scraper</param>
    ''' <param name="nMovie">New scraped movie data</param>
    ''' <param name="Options">What kind of data is being requested from the scrape(global scraper settings)</param>
    ''' <returns>Structures.DBMovie Object (nMovie) which contains the scraped data</returns>
    ''' <remarks></remarks>
    Function Scraper(ByRef oDBMovie As Structures.DBMovie, ByRef nMovie As MediaContainers.Movie, ByRef ScrapeType As Enums.ScrapeType, ByRef Options As Structures.ScrapeOptions_Movie) As Interfaces.ModuleResult Implements Interfaces.ScraperModule_Data_Movie.Scraper
        logger.Trace("Started RottenTomatoes Scraper")

        LoadSettings_Movie()

        Dim Settings As RottenTomatoes.Scraper.sMySettings_ForScraper
        Settings.ApiKey = _MySettings_Movie.APIKey
        Settings.FallBackEng = _MySettings_Movie.FallBackEng
        Settings.GetAdultItems = _MySettings_Movie.GetAdultItems
        Settings.PrefLanguage = _MySettings_Movie.PrefLanguage

        Dim _scraper As New RottenTomatoes.Scraper(Settings)
        Dim filterOptions As Structures.ScrapeOptions_Movie = Functions.MovieScrapeOptionsAndAlso(Options, ConfigOptions_Movie)

        If Master.GlobalScrapeMod.NFO AndAlso Not Master.GlobalScrapeMod.DoSearch Then
            If Not String.IsNullOrEmpty(oDBMovie.Movie.ID) Then
                'IMDB-ID already available -> scrape and save data into an empty movie container (nMovie)
                _scraper.GetMovieInfo(oDBMovie.Movie.ID, nMovie, filterOptions.bFullCrew, False, filterOptions, False)
            ElseIf Not String.IsNullOrEmpty(oDBMovie.Movie.TMDBID) Then
                'TMDB-ID already available -> scrape and save data into an empty movie container (nMovie)
                _scraper.GetMovieInfo(oDBMovie.Movie.TMDBID, nMovie, filterOptions.bFullCrew, False, filterOptions, False)
            ElseIf Not ScrapeType = Enums.ScrapeType.SingleScrape Then
                'no IMDB-ID or TMDB-ID for movie --> search first and try to get ID!
                If Not String.IsNullOrEmpty(oDBMovie.Movie.Title) Then
                    _scraper.GetSearchMovieInfo(oDBMovie.Movie.Title, oDBMovie, nMovie, ScrapeType, filterOptions)
                End If
                'if still no ID retrieved -> exit
                If String.IsNullOrEmpty(nMovie.TMDBID) Then Return New Interfaces.ModuleResult With {.breakChain = False, .Cancelled = True}
            End If
        End If

        If String.IsNullOrEmpty(nMovie.TMDBID) Then
            Select Case ScrapeType
                Case Enums.ScrapeType.FilterAuto, Enums.ScrapeType.FullAuto, Enums.ScrapeType.MarkAuto, Enums.ScrapeType.NewAuto, Enums.ScrapeType.MissAuto
                    nMovie = Nothing
                    Return New Interfaces.ModuleResult With {.breakChain = False}
            End Select
        End If

        If ScrapeType = Enums.ScrapeType.SingleScrape OrElse ScrapeType = Enums.ScrapeType.SingleAuto Then
            If String.IsNullOrEmpty(oDBMovie.Movie.ID) AndAlso String.IsNullOrEmpty(oDBMovie.Movie.TMDBID) Then
                Using dSearch As New dlgSearchResults_Movie(Settings, _scraper)
                    If dSearch.ShowDialog(nMovie, oDBMovie.Movie.Title, oDBMovie.Filename, filterOptions, oDBMovie.Movie.Year) = Windows.Forms.DialogResult.OK Then
                        _scraper.GetMovieInfo(nMovie.TMDBID, nMovie, filterOptions.bFullCrew, False, filterOptions, False)
                        'if a movie is found, set DoSearch back to "false" for following scrapers
                        Functions.SetScraperMod(Enums.ModType_Movie.DoSearch, False, False)
                    Else
                        nMovie = Nothing
                        Return New Interfaces.ModuleResult With {.breakChain = False, .Cancelled = True}
                    End If
                End Using
            End If
        End If

        'set new informations for following scrapers
        If Not String.IsNullOrEmpty(nMovie.Title) Then
            oDBMovie.Movie.Title = nMovie.Title
        End If
        If Not String.IsNullOrEmpty(nMovie.OriginalTitle) Then
            oDBMovie.Movie.OriginalTitle = nMovie.OriginalTitle
        End If
        If Not String.IsNullOrEmpty(nMovie.ID) Then
            oDBMovie.Movie.ID = nMovie.ID
        End If
        If Not String.IsNullOrEmpty(nMovie.IMDBID) Then
            oDBMovie.Movie.IMDBID = nMovie.IMDBID
        End If
        If Not String.IsNullOrEmpty(nMovie.TMDBID) Then
            oDBMovie.Movie.TMDBID = nMovie.TMDBID
        End If

        logger.Trace("Finished RottenTomatoes Scraper")
        Return New Interfaces.ModuleResult With {.breakChain = False}
    End Function

    Public Sub ScraperOrderChanged_Movie() Implements EmberAPI.Interfaces.ScraperModule_Data_Movie.ScraperOrderChanged
        _setup_Movie.orderChanged()
    End Sub

#End Region 'Methods

#Region "Nested Types"

    Structure sMySettings

#Region "Fields"
        Dim FallBackEng As Boolean
        Dim GetAdultItems As Boolean
        Dim APIKey As String
        Dim PrefLanguage As String
#End Region 'Fields

    End Structure

#End Region 'Nested Types

End Class