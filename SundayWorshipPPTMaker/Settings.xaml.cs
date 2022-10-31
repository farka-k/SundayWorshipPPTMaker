using System;
using System.Windows;
using System.Windows.Controls;
using System.IO;
using System.Diagnostics;
using System.Configuration;
using System.Collections.Specialized;
using Microsoft.Win32;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SundayWorshipPPTMaker
{
    ///	<summary>
    ///	템플릿 슬라이드 작업시 주요 슬라이드에 대한 슬라이드번호.
    ///	</summary>
    ///	<remarks>1-Base Indexing.</remarks>
    static class Constants
	{
		/// <summary>찬양 시작슬라이드. 제목 Shape가 있음.</summary>
		public const int PraiseEntry = 5;
		/// <summary>찬양 시작슬라이드 이동후 복사시 복사된 슬라이드번호는 6부터 시작.</summary>
		public const int PraiseSlidesInsertPos = 6;
		/// <summary>대표기도</summary>
		public const int PrayerNotice = 7;
		/// <summary말씀</summary>
		public const int BibleEntry = 8;
		/// <summary>설교 전 영상</summary>
		public const int VidBeforePreach = 10;
		/// <summary>설교제목</summary>
		public const int PreachEntry = 11;
		/// <summary>생일광고</summary>
		public const int AdBirthEntry = 23;
		/// <summary>생일자 명단</summary>
		public const int AdBirthList = 24;
	}

	/// <summary>
	/// Settings.xaml에 대한 상호 작용 논리
	/// </summary>
	public partial class Settings : Window
    {
		/// <summary>찬양 시작슬라이드. 제목 Shape가 있음.</summary>
		public int SettingsPraiseEntry;
		/// <summary>찬양 시작슬라이드 이동후 복사시 복사된 슬라이드번호는 6부터 시작.</summary>
		public int SettingsPraiseSlidesInsertPos;
		/// <summary>대표기도</summary>
		public int SettingsPrayerNotice;
		/// <summary말씀</summary>
		public int SettingsBibleEntry;
		/// <summary>설교 전 영상</summary>
		public int SettingsVidBeforePreach;
		/// <summary>설교제목</summary>
		public int SettingsPreachEntry;
		/// <summary>생일광고</summary>
		public int SettingsAdBirthEntry;
		/// <summary>생일자 명단</summary>
		public int SettingsAdBirthList;

		public string templateDirectory = AppDomain.CurrentDomain.BaseDirectory.ToString() + @"template\";
		public string templateFileName;
		public string templateFileFullPath;
		public Settings()
        {
            InitializeComponent();
			InitParams();
			SetParams();
		}

		/// <summary>
		/// app.config로부터 초기값 Load
		/// </summary>
		private void InitParams()
        {
			if (!Directory.Exists(templateDirectory))
            {
				Directory.CreateDirectory(templateDirectory);
            }

			//Initialize from app.config
			NumPraiseEntry.Value = int.Parse(ConfigurationManager.AppSettings.Get("PraiseEntry"));
			NumPraiseSlidesInsertPos.Value = int.Parse(ConfigurationManager.AppSettings.Get("PraiseSlidesInsertPosition"));
			NumPrayerNotice.Value = int.Parse(ConfigurationManager.AppSettings.Get("PrayerNotice"));
			NumBibleEntry.Value = int.Parse(ConfigurationManager.AppSettings.Get("BibleEntry"));
			NumVidBeforePreach.Value = int.Parse(ConfigurationManager.AppSettings.Get("VideoBeforePreach"));
			NumPreachEntry.Value = int.Parse(ConfigurationManager.AppSettings.Get("PreachEntry"));
			NumAdBirthEntry.Value = int.Parse(ConfigurationManager.AppSettings.Get("AdBirthEntry"));
			NumAdBirthList.Value = int.Parse(ConfigurationManager.AppSettings.Get("AdBirthList"));
			templateFileName = ConfigurationManager.AppSettings.Get("TemplateFileName");
			templateFileFullPath = templateDirectory + templateFileName;
			TxtTemplatePath.Text = templateFileFullPath;
			ShowFileErrorMessages();
		}

		/// <summary>
		/// Parameter가 유효한지 검사.
		/// </summary>
		/// <returns></returns>
		private bool ValidateParams()
        {
			string errorString = "";
			if (!File.Exists(TxtTemplatePath.Text))
            {
				errorString += "The template file does not exists.\n";
            }

			if (errorString.Length > 0)
            {
				errorString += "Click 'Ok' to set to default values or 'Cancel' to correct manually.";
				var msgBoxResult = MessageBox.Show(errorString, "Error", MessageBoxButton.OKCancel);
				if (msgBoxResult == MessageBoxResult.OK)
				{
					TxtTemplatePath.Text = templateDirectory + ConfigurationManager.AppSettings.Get("TemplateFileName");
					return true;
				}
				else
				{
					return false;
				}
            }
			return true;
        }

		/// <summary>
		/// Settings클래스 멤버 변수 확정
		/// </summary>
		private void SetParams()
		{ 
			SettingsPraiseEntry = NumPraiseEntry.Value.Value;
			SettingsPraiseSlidesInsertPos = NumPraiseSlidesInsertPos.Value.Value;
			SettingsPrayerNotice = NumPrayerNotice.Value.Value;
			SettingsBibleEntry = NumBibleEntry.Value.Value;
			SettingsVidBeforePreach = NumVidBeforePreach.Value.Value;
			SettingsPreachEntry = NumPreachEntry.Value.Value;
			SettingsAdBirthEntry = NumAdBirthEntry.Value.Value;
			SettingsAdBirthList = NumAdBirthList.Value.Value;
		}

		/// <summary>
		/// config를 갱신.
		/// </summary>
		private void UpdateConfig()
        {
			var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
			config.AppSettings.Settings["TemplateFileName"].Value = templateFileName;
			config.AppSettings.Settings["PraiseEntry"].Value = NumPraiseEntry.Text;
			config.AppSettings.Settings["PraiseSlidesInsertPosition"].Value = NumPraiseSlidesInsertPos.Text;
			config.AppSettings.Settings["PrayerNotice"].Value = NumPrayerNotice.Text;
			config.AppSettings.Settings["BibleEntry"].Value = NumBibleEntry.Text;
			config.AppSettings.Settings["VideoBeforePreach"].Value = NumVidBeforePreach.Text;
			config.AppSettings.Settings["PreachEntry"].Value = NumPreachEntry.Text;
			config.AppSettings.Settings["AdBirthEntry"].Value = NumAdBirthEntry.Text;
			config.AppSettings.Settings["AdBirthList"].Value = NumAdBirthList.Text;
			config.Save(ConfigurationSaveMode.Modified);
			ConfigurationManager.RefreshSection("appSettings");
		}
		
		/// <summary>
		/// 파라미터 값을 정해진 기본값으로 되돌린다.
		/// </summary>
		private void BtnDefault_Click(object sender, RoutedEventArgs e)
        {
			NumPraiseEntry.Value = Constants.PraiseEntry;
			NumPraiseSlidesInsertPos.Value = Constants.PraiseSlidesInsertPos;
			NumPrayerNotice.Value = Constants.PrayerNotice;
			NumBibleEntry.Value = Constants.BibleEntry;
			NumVidBeforePreach.Value = Constants.VidBeforePreach;
			NumPreachEntry.Value = Constants.PreachEntry;
			NumAdBirthEntry.Value = Constants.AdBirthEntry;
			NumAdBirthList.Value = Constants.AdBirthList;
			templateFileName = ConfigurationManager.AppSettings.Get("TemplateFileName");
		}

		/// <summary>
		/// 유효성검사, app.config갱신, 최종 변수 확정 후 창을 숨긴다.
		/// </summary>
        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
			if (!ValidateParams()) {
				return;
			}
			UpdateConfig();
			SetParams();
			Hide();
        }

		/// <summary>
		/// 템플릿 파일을 선택한다. 
		/// 선택한 파일이 템플릿 폴더 외부에 있는 파일이면 템플릿 폴더로 복사.
		/// </summary>
        private void BtnBrowseTemplate_Click(object sender, RoutedEventArgs e)
        {
			OpenFileDialog ofd = new OpenFileDialog();
			ofd.Filter = "Presentation Files(*.ppt;*.pptx)|*.ppt;*.pptx|All Files(*.*)|*.*";
			ofd.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory + @"template\";
			if (ofd.ShowDialog() == true)
			{
				string fileName = ofd.SafeFileName;
				
				if (!ofd.FileName.StartsWith(templateDirectory))
                {
					if (File.Exists(templateDirectory + fileName))
				    {
						int extIndex = fileName.LastIndexOf(".");
						string ext = fileName.Substring(extIndex);
						fileName = fileName.Substring(0, extIndex) + "_new" + ext;
					}
					File.Copy(ofd.FileName, templateDirectory + fileName);
                }
				templateFileName = fileName;
				templateFileFullPath = templateDirectory + templateFileName;
				TxtTemplatePath.Text = templateFileFullPath;
				ShowFileErrorMessages();
			}
		}
		
		/// <summary>
		/// 템플릿 폴더 열기
		/// </summary>
        private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
        {
			ProcessStartInfo startInfo = new ProcessStartInfo { 
				Arguments = templateDirectory, 
				FileName = "explorer.exe" 
			};
			Process.Start(startInfo);
        }
        
		/// <summary>
		/// 선택된 ppt 열기
		/// </summary>
		private void BtnOpenTemplateFile_Click(object sender, RoutedEventArgs e)
        {
			PowerPoint.Application pptApp = new PowerPoint.Application();
			PowerPoint.Presentations pptPres = pptApp.Presentations;
			PowerPoint.Presentation presentation = pptPres.Open(TxtTemplatePath.Text);
		}

		private void ShowFileErrorMessages()
        {
			if (!File.Exists(templateFileFullPath))
				TxtTemplateFileError.Text = "파일이 존재하지 않습니다. 템플릿 파일을 다시 선택하세요.";
			else
				TxtTemplateFileError.Text = "";
        }
    }
}
