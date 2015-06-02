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
' ###############################################################################

Imports System.IO
Imports EmberAPI
Imports NLog
Imports System.Diagnostics

Public Class OpenSubtitle_Subtitle
    Implements Interfaces.ScraperModule_Subtitle_Movie


#Region "Fields"

    Shared logger As Logger = NLog.LogManager.GetCurrentClassLogger()

    Public Shared ConfigScrapeModifier_Movie As New Structures.ScrapeModifier_Movie_MovieSet
    Public Shared ConfigScrapeModifier_MovieSet As New Structures.ScrapeModifier_Movie_MovieSet
    Public Shared _AssemblyName As String

    ''' <summary>
    ''' Scraping Here
    ''' </summary>
    ''' <remarks></remarks>
    Private strPrivateAPIKey As String = String.Empty
    Private _MySettings_Movie As New sMySettings
    Private _Name As String = "OpenSubtitles_Subtitle"
    Private _ScraperEnabled_Movie As Boolean = False
    Private _setup_Movie As frmSettingsHolder_Movie

#End Region 'Fields

#Region "Events"

    'Movie part
    Public Event ModuleSettingsChanged_Movie() Implements Interfaces.ScraperModule_Subtitle_Movie.ModuleSettingsChanged
    Public Event MovieScraperEvent_Movie(ByVal eType As Enums.ScraperEventType_Movie, ByVal Parameter As Object) Implements Interfaces.ScraperModule_Subtitle_Movie.ScraperEvent
    Public Event SetupScraperChanged_Movie(ByVal name As String, ByVal State As Boolean, ByVal difforder As Integer) Implements Interfaces.ScraperModule_Subtitle_Movie.ScraperSetupChanged
    Public Event SetupNeedsRestart_Movie() Implements Interfaces.ScraperModule_Subtitle_Movie.SetupNeedsRestart
    Public Event ProgressUpdated_Movie(ByVal iPercent As Integer) Implements Interfaces.ScraperModule_Subtitle_Movie.ProgressUpdated

#End Region 'Events

#Region "Properties"

    ReadOnly Property ModuleName() As String Implements Interfaces.ScraperModule_Subtitle_Movie.ModuleName
        Get
            Return _Name
        End Get
    End Property

    ReadOnly Property ModuleVersion() As String Implements Interfaces.ScraperModule_Subtitle_Movie.ModuleVersion
        Get
            Return System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileVersion.ToString
        End Get
    End Property

    Property ScraperEnabled_Movie() As Boolean Implements Interfaces.ScraperModule_Subtitle_Movie.ScraperEnabled
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
        RaiseEvent SetupScraperChanged_Movie(String.Concat(Me._Name, "_Movie"), state, difforder)
    End Sub

    Sub Init_Movie(ByVal sAssemblyName As String) Implements Interfaces.ScraperModule_Subtitle_Movie.Init
        _AssemblyName = sAssemblyName
        LoadSettings_Movie()
    End Sub

    Function InjectSetupScraper_Movie() As Containers.SettingsPanel Implements Interfaces.ScraperModule_Subtitle_Movie.InjectSetupScraper
        Dim Spanel As New Containers.SettingsPanel
        _setup_Movie = New frmSettingsHolder_Movie
        LoadSettings_Movie()
        _setup_Movie.chkEnabled.Checked = _ScraperEnabled_Movie
        _setup_Movie.txtApiKey.Text = strPrivateAPIKey
        _setup_Movie.cbPrefLanguage.Text = _MySettings_Movie.PrefLanguage
        _setup_Movie.chkGetBlankImages.Checked = _MySettings_Movie.GetBlankImages
        _setup_Movie.chkGetEnglishImages.Checked = _MySettings_Movie.GetEnglishImages
        _setup_Movie.chkPrefLanguageOnly.Checked = _MySettings_Movie.PrefLanguageOnly
        _setup_Movie.Lang = _setup_Movie.cbPrefLanguage.Text
        _setup_Movie.API = _setup_Movie.txtApiKey.Text

        If Not String.IsNullOrEmpty(strPrivateAPIKey) Then
            _setup_Movie.btnUnlockAPI.Text = Master.eLang.GetString(443, "Use embedded API Key")
            _setup_Movie.lblEMMAPI.Visible = False
            _setup_Movie.txtApiKey.Enabled = True
        End If

        _setup_Movie.orderChanged()

        Spanel.Name = String.Concat(Me._Name, "_Movie")
        Spanel.Text = "OpenSubtitles"
        Spanel.Prefix = "OpenSubtitlesMovieSubtitle_"
        Spanel.Order = 110
        Spanel.Parent = "pnlMovieSubtitle"
        Spanel.Type = Master.eLang.GetString(36, "Movies")
        Spanel.ImageIndex = If(Me._ScraperEnabled_Movie, 9, 10)
        Spanel.Panel = Me._setup_Movie.pnlSettings

        AddHandler _setup_Movie.SetupScraperChanged, AddressOf Handle_SetupScraperChanged_Movie
        AddHandler _setup_Movie.ModuleSettingsChanged, AddressOf Handle_ModuleSettingsChanged_Movie
        AddHandler _setup_Movie.SetupNeedsRestart, AddressOf Handle_SetupNeedsRestart_Movie
        Return Spanel
    End Function

    Sub LoadSettings_Movie()

        strPrivateAPIKey = clsAdvancedSettings.GetSetting("APIKey", "", , Enums.Content_Type.Movie)
        _MySettings_Movie.APIKey = If(String.IsNullOrEmpty(strPrivateAPIKey), "44810eefccd9cb1fa1d57e7b0d67b08d", strPrivateAPIKey)
        _MySettings_Movie.GetBlankImages = clsAdvancedSettings.GetBooleanSetting("GetBlankImages", False, , Enums.Content_Type.Movie)
        _MySettings_Movie.GetEnglishImages = clsAdvancedSettings.GetBooleanSetting("GetEnglishImages", False, , Enums.Content_Type.Movie)
        _MySettings_Movie.PrefLanguage = clsAdvancedSettings.GetSetting("PrefLanguage", "en", , Enums.Content_Type.Movie)
        _MySettings_Movie.PrefLanguageOnly = clsAdvancedSettings.GetBooleanSetting("PrefLanguageOnly", False, , Enums.Content_Type.Movie)

        ConfigScrapeModifier_Movie.Poster = clsAdvancedSettings.GetBooleanSetting("DoPoster", True, , Enums.Content_Type.Movie)
        ConfigScrapeModifier_Movie.Fanart = clsAdvancedSettings.GetBooleanSetting("DoFanart", True, , Enums.Content_Type.Movie)

    End Sub

    Function Scraper(ByRef DBMovie As Structures.DBMovie, ByRef SubtitleList As List(Of MediaContainers.Subtitle)) As Interfaces.ModuleResult Implements Interfaces.ScraperModule_Subtitle_Movie.Scraper
        logger.Trace("Started scrape OpenSubtitles")

        LoadSettings_Movie()

        If String.IsNullOrEmpty(DBMovie.Movie.TMDBID) Then
            DBMovie.Movie.TMDBID = ModulesManager.Instance.GetMovieTMDBID(DBMovie.Movie.ID)
        End If

        If Not String.IsNullOrEmpty(DBMovie.Movie.TMDBID) Then
            Dim Settings As OpenSubtitle.Scraper.sMySettings_ForScraper
            Settings.GetBlankImages = _MySettings_Movie.GetBlankImages
            Settings.GetEnglishImages = _MySettings_Movie.GetEnglishImages
            Settings.APIKey = _MySettings_Movie.APIKey
            Settings.PrefLanguage = _MySettings_Movie.PrefLanguage
            Settings.PrefLanguageOnly = _MySettings_Movie.PrefLanguageOnly

            Dim _scraper As New OpenSubtitle.Scraper

            SubtitleList = _scraper.GetSubtitles(DBMovie.Filename, DBMovie.Movie.IMDBID, Settings)
        End If

        logger.Trace("Finished OpenSubtitles Scraper")
        Return New Interfaces.ModuleResult With {.breakChain = False}
    End Function

    Sub SaveSettings_Movie()
        Using settings = New clsAdvancedSettings()
            settings.SetBooleanSetting("DoPoster", ConfigScrapeModifier_Movie.Poster, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("DoFanart", ConfigScrapeModifier_Movie.Fanart, , , Enums.Content_Type.Movie)

            settings.SetSetting("APIKey", _setup_Movie.txtApiKey.Text, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("GetBlankImages", _MySettings_Movie.GetBlankImages, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("GetEnglishImages", _MySettings_Movie.GetEnglishImages, , , Enums.Content_Type.Movie)
            settings.SetBooleanSetting("PrefLanguageOnly", _MySettings_Movie.PrefLanguageOnly, , , Enums.Content_Type.Movie)
            settings.SetSetting("PrefLanguage", _MySettings_Movie.PrefLanguage, , , Enums.Content_Type.Movie)
        End Using
    End Sub

    Sub SaveSetupScraper_Movie(ByVal DoDispose As Boolean) Implements Interfaces.ScraperModule_Subtitle_Movie.SaveSetupScraper
        _MySettings_Movie.PrefLanguage = _setup_Movie.cbPrefLanguage.Text
        _MySettings_Movie.GetBlankImages = _setup_Movie.chkGetBlankImages.Checked
        _MySettings_Movie.GetEnglishImages = _setup_Movie.chkGetEnglishImages.Checked
        _MySettings_Movie.PrefLanguageOnly = _setup_Movie.chkPrefLanguageOnly.Checked
        SaveSettings_Movie()
        'ModulesManager.Instance.SaveSettings()
        If DoDispose Then
            RemoveHandler _setup_Movie.SetupScraperChanged, AddressOf Handle_SetupScraperChanged_Movie
            RemoveHandler _setup_Movie.ModuleSettingsChanged, AddressOf Handle_ModuleSettingsChanged_Movie
            RemoveHandler _setup_Movie.SetupNeedsRestart, AddressOf Handle_SetupNeedsRestart_Movie
            _setup_Movie.Dispose()
        End If
    End Sub

    Public Sub ScraperOrderChanged_Movie() Implements EmberAPI.Interfaces.ScraperModule_Subtitle_Movie.ScraperOrderChanged
        _setup_Movie.orderChanged()
    End Sub

#End Region 'Methods

#Region "Nested Types"

    Structure sMySettings

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