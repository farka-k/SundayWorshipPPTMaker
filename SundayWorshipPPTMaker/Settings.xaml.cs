using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.IO;
using System.Diagnostics;
using System.Configuration;
using System.Collections.Specialized;
using Microsoft.Win32;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SundayWorshipPPTMaker
{
	/// <summary>
	/// Settings.xaml에 대한 상호 작용 논리
	/// </summary>
	public partial class Settings : Window
	{
		internal SlideSizeType SlideSizeType;
		internal OCREngineType OCREngineType;
		internal slideSize SlideSize;
		internal string ImgUriCreed;
		internal string ImgUriPray;
		internal string ImgUriLordsPrayer;
		internal string ImgUriAfterPraise;
		internal string ImgUriBeforeAd;
		internal string ImgUriAfterAd;

		public Settings()
		{
			InitializeComponent();
			InitParams();
			SetParams();
		}

		public string GetSlideSizeTypeString { 
			get
            {
				if (SlideSizeType == SlideSizeType.Normal)
					return "4x3";
				else
					return "16x9";
            } 
		}

		private void InitImages()
        {
			var imgFullPath = Constants.BaseDirectory + ConfigurationManager.AppSettings.Get("ImgUriCreed");
            if (File.Exists(imgFullPath))
				ImgCreed.Source=new BitmapImage(new Uri(imgFullPath));
			else
				ImgCreed.Source = new BitmapImage(new Uri(Properties.Resources.BGUriCross01, UriKind.Relative));

			imgFullPath = Constants.BaseDirectory + ConfigurationManager.AppSettings.Get("ImgUriPray");
			if (File.Exists(imgFullPath))
				ImgPray.Source = new BitmapImage(new Uri(imgFullPath));
			else
				ImgPray.Source = new BitmapImage(new Uri(Properties.Resources.BGUriPray01, UriKind.Relative));

			imgFullPath = Constants.BaseDirectory + ConfigurationManager.AppSettings.Get("ImgUriLordsPrayer");
			if (File.Exists(imgFullPath))
				ImgLordsPrayer.Source = new BitmapImage(new Uri(imgFullPath));
			else
				ImgLordsPrayer.Source = new BitmapImage(new Uri(Properties.Resources.BGUriCross02, UriKind.Relative));

			imgFullPath = Constants.BaseDirectory + ConfigurationManager.AppSettings.Get("ImgUriAfterPraise");
			if (File.Exists(imgFullPath))
				ImgAfterPraise.Source = new BitmapImage(new Uri(imgFullPath));
			else
				ImgAfterPraise.Source= new BitmapImage(new Uri(Properties.Resources.BGUriCutImg01, UriKind.Relative));

			imgFullPath = Constants.BaseDirectory + ConfigurationManager.AppSettings.Get("ImgUriBeforeAd");
			if (File.Exists(imgFullPath))
				ImgBeforeAd.Source = new BitmapImage(new Uri(imgFullPath));
			else
				ImgBeforeAd.Source= new BitmapImage(new Uri(Properties.Resources.BGUriCutImg02, UriKind.Relative));

			imgFullPath = Constants.BaseDirectory + ConfigurationManager.AppSettings.Get("ImgUriAfterAd");
			if (File.Exists(imgFullPath))
				ImgAfterAd.Source = new BitmapImage(new Uri(imgFullPath));
			else
				ImgAfterAd.Source= new BitmapImage(new Uri(Properties.Resources.BGUriCutImg03, UriKind.Relative));
		}

		/// <summary>
		/// app.config로부터 초기값 Load
		/// </summary>
		private void InitParams()
        {
			//Initialize from app.config
			var SizeCode = int.Parse(ConfigurationManager.AppSettings.Get("SlideSize"));
			if (SizeCode == (int)SlideSizeType.WideScreen)
				RadioSlideSizeWideScreen.IsChecked = true;
			else
				RadioSlideSizeNormal.IsChecked = true;

			var OCRCode = int.Parse(ConfigurationManager.AppSettings.Get("OCREngine"));
			if (OCRCode == (int)OCREngineType.Clova)
				RadioOCRClova.IsChecked = true;
			else
				RadioOCRTesseract.IsChecked = true;

			InitImages();
		}

		/// <summary>
		/// Parameter가 유효한지 검사.
		/// </summary>
		/// <returns></returns>
		/*private bool ValidateParams()
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
        }*/

		/// <summary>
		/// Settings클래스 멤버 변수 확정
		/// </summary>
		private void SetParams()
		{ 
			if (RadioSlideSizeWideScreen.IsChecked==true)
            {
				SlideSizeType = SlideSizeType.WideScreen;
				SlideSize.Width = Utils.CMToPoint(Constants.SlideSize16x9Width);
				SlideSize.Height = Utils.CMToPoint(Constants.SlideSizeHeight);
			}
            else
            {
				SlideSizeType = SlideSizeType.Normal;
				SlideSize.Width = Utils.CMToPoint(Constants.SlideSize4x3Width);
				SlideSize.Height = Utils.CMToPoint(Constants.SlideSizeHeight);
            }

			if (RadioOCRClova.IsChecked == true) OCREngineType = OCREngineType.Clova;
			else OCREngineType = OCREngineType.Tesseract;
			ImgUriCreed = new Uri(ImgCreed.Source.ToString()).AbsolutePath;
			ImgUriPray = new Uri(ImgPray.Source.ToString()).AbsolutePath;
			ImgUriLordsPrayer = new Uri(ImgLordsPrayer.Source.ToString()).AbsolutePath;
			ImgUriAfterPraise = new Uri(ImgAfterPraise.Source.ToString()).AbsolutePath;
			ImgUriBeforeAd = new Uri(ImgBeforeAd.Source.ToString()).AbsolutePath;
			ImgUriAfterAd = new Uri(ImgAfterAd.Source.ToString()).AbsolutePath;
		}

		/// <summary>
		/// config를 갱신.
		/// </summary>
		private void UpdateConfig()
        {
			var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
			config.AppSettings.Settings["SlideSize"].Value = SlideSizeType.ToString();
			config.AppSettings.Settings["OCREngine"].Value = OCREngineType.ToString();
			config.AppSettings.Settings["ImgUriCreed"].Value = new Uri(ImgCreed.Source.ToString()).AbsolutePath;
			config.AppSettings.Settings["ImgUriPray"].Value = new Uri(ImgPray.Source.ToString()).AbsolutePath;
			config.AppSettings.Settings["ImgUriLordsPrayer"].Value = new Uri(ImgLordsPrayer.Source.ToString()).AbsolutePath;
			config.AppSettings.Settings["ImgUriAfterPraise"].Value = new Uri(ImgAfterPraise.Source.ToString()).AbsolutePath;
			config.AppSettings.Settings["ImgUriBeforeAd"].Value = new Uri(ImgBeforeAd.Source.ToString()).AbsolutePath;
			config.AppSettings.Settings["ImgUriAfterAd"].Value = new Uri(ImgAfterAd.Source.ToString()).AbsolutePath;
			config.Save(ConfigurationSaveMode.Modified);
			ConfigurationManager.RefreshSection("appSettings");
		}
		
		/// <summary>
		/// 파라미터 값을 정해진 기본값으로 되돌린다.
		/// </summary>
		private void BtnDefault_Click(object sender, RoutedEventArgs e)
        {
			/*NumPraiseEntry.Value = Constants.PraiseEntry;
			NumPraiseSlidesInsertPos.Value = Constants.PraiseSlidesInsertPos;
			NumPrayerNotice.Value = Constants.PrayerNotice;
			NumBibleEntry.Value = Constants.BibleEntry;
			NumVidBeforePreach.Value = Constants.VidBeforePreach;
			NumPreachEntry.Value = Constants.PreachEntry;
			NumAdBirthEntry.Value = Constants.AdBirthEntry;
			NumAdBirthList.Value = Constants.AdBirthList;
			templateFileName = ConfigurationManager.AppSettings.Get("TemplateFileName");*/
		}

		/// <summary>
		/// 유효성검사, app.config갱신, 최종 변수 확정 후 창을 숨긴다.
		/// </summary>
        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
			SetParams();
			UpdateConfig();
			Hide();
        }

        private void ImgCreed_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlCreed.Opacity = 1;
        }

        private void ImgCreed_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlCreed.Opacity = 0;
        }

        private void ImgCtrlCreed_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlCreed.Opacity = 1;
        }

        private void ImgPray_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlPray.Opacity = 1;
        }

        private void ImgPray_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlPray.Opacity = 0;
		}

        private void ImgCtrlPray_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlPray.Opacity = 1;
        }

        private void ImgLordsPrayer_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlLordsPrayer.Opacity = 1;
		}

        private void ImgLordsPrayer_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlLordsPrayer.Opacity = 0;
		}

        private void ImgCtrlLordsPrayer_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlLordsPrayer.Opacity = 1;
        }

        private void ImgAfterPraise_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlAfterPraise.Opacity = 1;
        }

        private void ImgAfterPraise_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlAfterPraise.Opacity = 0;
		}

        private void ImgCtrlAfterPraise_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlAfterPraise.Opacity = 1;
        }

        private void ImgBeforeAd_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlBeforeAd.Opacity = 1;
        }

        private void ImgBeforeAd_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlBeforeAd.Opacity = 0;
		}

        private void ImgCtrlBeforeAd_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlBeforeAd.Opacity = 1;
        }

        private void ImgAfterAd_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlAfterAd.Opacity = 1;
		}

        private void ImgAfterAd_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlAfterAd.Opacity = 0;
		}

        private void ImgCtrlAfterAd_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
			ImgCtrlAfterAd.Opacity = 1;
        }

        private void BrowseLocalImage(object sender, RoutedEventArgs e)
        {
			OpenFileDialog ofd = new OpenFileDialog();
			ofd.Filter = "Image Files (*.jpg;*.png;*.bmp)|*.jpg;*.png;*.bmp|All Files (*.*)|*.*";
			ofd.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory+@"Resources\";
            if (ofd.ShowDialog() == true)
            {
				string fileName = ofd.SafeFileName;
				if (!ofd.FileName.StartsWith(Constants.ResourceDirectory))
				{
					if (File.Exists(Constants.ResourceDirectory + fileName))
					{
						int extIndex = fileName.LastIndexOf(".");
						string ext = fileName.Substring(extIndex);
						fileName = fileName.Substring(0, extIndex) + "_new" + ext;
					}
					File.Copy(ofd.FileName, Constants.ResourceDirectory + fileName);
				}
				
				var parentObjName = ((StackPanel)((Button)sender).Parent).Name;
				if (parentObjName.Contains("Creed"))
					ImgCreed.Source = new BitmapImage(new Uri(Constants.ResourceDirectory + fileName));
				if (parentObjName.Contains("Pray"))
					ImgPray.Source = new BitmapImage(new Uri(Constants.ResourceDirectory + fileName));
				if (parentObjName.Contains("Lords"))
					ImgLordsPrayer.Source = new BitmapImage(new Uri(Constants.ResourceDirectory + fileName));
				if (parentObjName.Contains("Praise"))
					ImgAfterPraise.Source = new BitmapImage(new Uri(Constants.ResourceDirectory + fileName));
				if (parentObjName.Contains("Before"))
					ImgBeforeAd.Source = new BitmapImage(new Uri(Constants.ResourceDirectory + fileName));
				if (parentObjName.Contains("rAd"))
					ImgAfterAd.Source = new BitmapImage(new Uri(Constants.ResourceDirectory + fileName));
			}
        }

        private void BrowseWebImage(object sender, RoutedEventArgs e)
        {

        }

        private void SetImageToDefault(object sender, RoutedEventArgs e)
        {
			var parentObjName = ((StackPanel)((Button)sender).Parent).Name;
			if (parentObjName.Contains("Creed"))
				ImgCreed.Source = new BitmapImage(new Uri(Properties.Resources.BGUriCross01, UriKind.Relative));
			if (parentObjName.Contains("Pray"))
				ImgPray.Source = new BitmapImage(new Uri(Properties.Resources.BGUriPray01, UriKind.Relative));
			if (parentObjName.Contains("Lords"))
				ImgLordsPrayer.Source = new BitmapImage(new Uri(Properties.Resources.BGUriCross02, UriKind.Relative));
			if (parentObjName.Contains("Praise"))
				ImgAfterPraise.Source = new BitmapImage(new Uri(Properties.Resources.BGUriCutImg01, UriKind.Relative));
			if (parentObjName.Contains("Before"))
				ImgBeforeAd.Source = new BitmapImage(new Uri(Properties.Resources.BGUriCutImg02, UriKind.Relative));
			if (parentObjName.Contains("rAd"))
				ImgAfterAd.Source = new BitmapImage(new Uri(Properties.Resources.BGUriCutImg03, UriKind.Relative));
		}

        /// <summary>
        /// 템플릿 파일을 선택한다. 
        /// 선택한 파일이 템플릿 폴더 외부에 있는 파일이면 템플릿 폴더로 복사.
        /// </summary>
        /*private void BtnBrowseTemplate_Click(object sender, RoutedEventArgs e)
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
		}*/

        /// <summary>
        /// 템플릿 폴더 열기
        /// </summary>
        /*private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
        {
			ProcessStartInfo startInfo = new ProcessStartInfo { 
				Arguments = templateDirectory, 
				FileName = "explorer.exe" 
			};
			Process.Start(startInfo);
        }*/

        /// <summary>
        /// 선택된 ppt 열기
        /// </summary>
        /*private void BtnOpenTemplateFile_Click(object sender, RoutedEventArgs e)
        {
			/*PowerPoint.Application pptApp = new PowerPoint.Application();
			PowerPoint.Presentations pptPres = pptApp.Presentations;
			PowerPoint.Presentation presentation = pptPres.Open(TxtTemplatePath.Text);
		}*/

        /*private void ShowFileErrorMessages()
        {
			if (!File.Exists(templateFileFullPath))
				TxtTemplateFileError.Text = "파일이 존재하지 않습니다. 템플릿 파일을 다시 선택하세요.";
			else
				TxtTemplateFileError.Text = "";
        }*/
    }
}
