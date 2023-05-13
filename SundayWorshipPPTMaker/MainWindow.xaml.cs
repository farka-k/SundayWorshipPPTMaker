using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Text.RegularExpressions;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.Diagnostics;

using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using HwpObjectLib;
using System.Data.SQLite;
using Tesseract;

namespace SundayWorshipPPTMaker
{
	public partial class MainWindow : Window
	{
		Settings settings;
		public List<string> books=new List<string>();
		public List<string> abbr=new List<string>();
		public List<int> numOfChapters = new List<int>();
		public string workFolder;
		private DateTime dt;
		private Jubo jubo;
		string OutputDirectory = @"\Out\";
		private SQLiteConnection conn = null;
		private SQLiteCommand command = null;
		private SQLiteDataReader rdr = null;

		/// <summary>
		/// MainWindow 초기화 코드
		/// </summary>
		public MainWindow()
		{
			InitializeComponent();

			//GetBibleBooks();
			GetBibleBooksDB();
			InitComponentsValues();

			
			jubo = new Jubo(dt.ToString("yy. M. d"));

            RegisterHWPSecurityModule();
			
		}

		/// <summary>
		/// Hwp Object Library 보안 모듈을 레지스트리에 등록한다.
		/// </summary>
		private static void RegisterHWPSecurityModule()
        {
			Microsoft.Win32.Registry.SetValue(Properties.Resources.HncRoot, "FilePathChecker", AppDomain.CurrentDomain.BaseDirectory + @"\FilePathCheckerModuleExample.dll");
		}
		/// <summary>
		/// Get the list of Bible Books from embedded text file
		/// </summary>
		private void GetBibleBooks()
        {
			Assembly _assembly;
			StreamReader _textStreamReader = null;
			try
			{
				_assembly = Assembly.GetExecutingAssembly();
				_textStreamReader = new StreamReader(_assembly.GetManifestResourceStream(this.GetType().Namespace+".BibleBooks.txt"));
			}
			catch
			{
				MessageBox.Show("Error accessing resources!");
				Close();
			}
			finally
			{
				string line;
				while ((line = _textStreamReader.ReadLine()) != null)
				{
					string[] s = line.Split();
					books.Add(s[0]);
					abbr.Add(s[1]);
				}
			}
		}

		private void GetBibleBooksDB()
        {
			if (!File.Exists("RevisedKorBible.db"))
            {
				GetBibleBooks();
				return;
            }
			conn = new SQLiteConnection("Data Source=RevisedKorBible.db;Version=3");
			conn.Open();
			command = conn.CreateCommand();
			command.CommandText = String.Format("select Name,Abbr,Chapters from Books");
			rdr = command.ExecuteReader();

            while (rdr.Read())
            {
				books.Add(rdr.GetString(0));
				abbr.Add(rdr.GetString(1));
				numOfChapters.Add(rdr.GetInt32(2));
            }
        }

		private void InitComponentsValues()
        {
			GetBibleBooksDB();
			LoadLogoImage();
			CmbStartBook.ItemsSource = books;
			CmbEndBook.ItemsSource = books;
			CmbStartBook.SelectedIndex = 0;
			CmbEndBook.SelectedIndex = 0;
			NumStartChapter.Value = 1;
			NumStartPassage.Value = 1;
			NumEndChapter.Value = 1;
			NumEndPassage.Value = 1;

			dt = Utils.GetComingSundayDate();
			TxtOutputFileName.Text = dt.ToString("yyyy.MM.dd") + " 고등부 예배.pptx";

			settings = new Settings();
		}

		/// <summary>
		/// 메인화면의 로고 이미지 로딩.
		/// </summary>
		private void LoadLogoImage()
        {
			BitmapImage image = new BitmapImage(new Uri("pack://application:,,,/Resources/logo02.png"));
			imageLogo.Source = image;
        }

		/// <summary>
		/// 새 PPT에 필요한 모든 파일이 저장된 폴더를 작업폴더로 설정한다.
		/// </summary>
		private void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
		{
			var dlg = new CommonOpenFileDialog();
			dlg.IsFolderPicker = true;

			if (dlg.ShowDialog() == CommonFileDialogResult.Ok)
			{
				TxtOutputFolder.Text = dlg.FileName;
				workFolder = dlg.FileName;
				jubo.SetPathInfo(workFolder);

				HwpObject hwp = new HwpObject();
				hwp.RegisterModule("FilePathCheckDLL", "FilePathChecker");

				if (hwp.Open(jubo.FullPath, "", ""))
				{
					string txt = (string)hwp.GetTextFile("TEXT", "");
					if (!System.IO.Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + OutputDirectory))
						System.IO.Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + OutputDirectory);
					System.IO.File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + OutputDirectory + dt.ToString("yy-MM-dd") + ".txt", txt);
					ManualMode.IsChecked = false;

					jubo.Parse(txt);
					MessageBox.Show(jubo.GetJuboInfo());
				}
				else
				{
					/*var engine = new TesseractEngine(@"./tessdata", "kor", EngineMode.Default);
					var img = Pix.LoadFromFile(workFolder + @"/230430_2cr1.jpg");
					var page = engine.Process(img);
					var text = page.GetText();
					List<string> wordList = text.Trim().Split(new string[] { " ", "\r\n" }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
					string new_text="";
					foreach(var item in wordList)
                    {
						new_text += item;
                    }
					MessageBox.Show(new_text);*/
					MessageBox.Show("File Open Failed.\nChange to manual mode.");
					ManualMode.IsChecked = true;
					return;
				}
				
				TxtPrayer.Text = jubo.PrayerName;
				TxtTitle.Text = jubo.PreachTitle;
				
				//Update Fields
				CmbStartBook.SelectedIndex=books.IndexOf(jubo.BVSStart.book);
				NumStartChapter.Value = jubo.BVSStart.chapter;
				NumStartPassage.Value = jubo.BVSStart.passage;
				//TxtStartChapter.Text = jubo.BVSStart.chapter.ToString();
				//TxtStartPassage.Text = jubo.BVSStart.passage.ToString();
				CmbEndBook.SelectedIndex = books.IndexOf(jubo.BVSEnd.book);
				NumEndChapter.Value = jubo.BVSEnd.chapter;
				NumEndPassage.Value = jubo.BVSEnd.passage;
				//TxtEndChapter.Text = jubo.BVSEnd.chapter.ToString();
				//TxtEndPassage.Text = jubo.BVSEnd.passage.ToString();

				//생일 필드
				//생일자 있으면
				if (jubo.IsBirthday)
				{
					CbBirth.IsChecked = true;
					TxtBirthList.Text = jubo.GetBirthAdString();
				}
				else CbBirth.IsChecked = false;
			}
		}
		/// <summary>
		/// 하나의 파일 이름을 불러온다.
		/// </summary>
		/// <remarks>이벤트를 발생시킨 버튼 이름에 따라 FileDialog의 형태가 바뀐다.</remarks>
		private void BtnBrowseSingleFile_Click(object sender, RoutedEventArgs e)
		{

			OpenFileDialog ofd = new OpenFileDialog();
			string sender_name = ((Button)sender).Name;
			ofd.InitialDirectory = workFolder;

			if (sender_name == "BtnBrowseVid")
			{
				ofd.Filter = "Video Files(*.avi;*.flv;*.mp4;*.wmv;*.mkv)|*.avi;*.flv;*.mp4;*.wmv;*.mkv|All Files(*.*)|*.*";
			}
			else { 
				ofd.Filter = "Presentation Files(*.ppt;*.pptx)|*.ppt;*.pptx|All Files(*.*)|*.*";
			}
			if (ofd.ShowDialog() == true)
			{
				if (sender_name == "BtnBrowsePreach")
					TxtPreachLocation.Text = ofd.FileName;
				else    //sender_name=="BtnBrowseVid"
					TxtVidLocation.Text = ofd.FileName;
			}
		}
		/// <summary>
		/// 선택된 아이템의 위치를 위로 한 칸 올린다.
		/// </summary>
		private void BtnOrderUp_Click(object sender, RoutedEventArgs e)
		{
            MoveItem(SongList, -1);
		}
		/// <summary>
		/// 선택된 아이템의 위치를 아래로 한 칸 내린다.
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void BtnOrderDown_Click(object sender, RoutedEventArgs e)
		{
            MoveItem(SongList, 1);
		}
		/// <summary>
		/// 방향에 따라 아이템 위치 변경을 수행.
		/// </summary>
		/// <param name="lb">ListBox Object</param>
		/// <param name="direction">-1이면 위로, 1이면 아래로 이동</param>
		private static void MoveItem(ListBox lb, int direction)
		{
			if (lb.SelectedItems.Count != 1)
				return;     //No selected item or Selected multiple items
			int newIdx = lb.SelectedIndex + direction;
			if (newIdx < 0 || newIdx >= lb.Items.Count)
				return;     //Index out of range

			object selected = lb.SelectedItem;

			lb.Items.Remove(selected);
			lb.Items.Insert(newIdx, selected);
			lb.SelectedIndex = newIdx;
		}
		/// <summary>
		/// 찬양 PPT 목록에 아이템을 추가한다.
		/// </summary>
		private void BtnAddSongs_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog ofd = new OpenFileDialog();
			ofd.Multiselect = true;
			ofd.InitialDirectory = workFolder;
			ofd.Filter = "Presentation Files(*.ppt;*.pptx)|*.ppt;*.pptx|All Files(*.*)|*.*";
			if (ofd.ShowDialog() == true)
			{
				foreach (string path in ofd.FileNames)
				{
					if (!SongList.Items.Contains(path))
						SongList.Items.Add(path);
				}
			}
		}
		/// <summary>
		/// 찬양 PPT 목록에서 선택된 아이템을 삭제한다.
		/// </summary>
		private void BtnDelSongs_Click(object sender, RoutedEventArgs e)
		{
			int count = SongList.SelectedItems.Count;
			for (int i = 0; i < count; i++)
			{
				SongList.Items.RemoveAt(SongList.SelectedIndex);
			}
		}
		/// <summary>
		/// 시작 범위의 책과 끝 범위의 책을 같이 변경한다.
		/// </summary>
		private void CmbBook_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
            if (((ComboBox)sender).Name == "CmbStartBook") 
				CmbEndBook.SelectedIndex = CmbStartBook.SelectedIndex;
			else
				if (CmbStartBook.SelectedIndex > CmbEndBook.SelectedIndex)
					CmbStartBook.SelectedIndex = CmbEndBook.SelectedIndex;
			
			int num= numOfChapters[((ComboBox)sender).SelectedIndex];
			NumStartChapter.Maximum = num;
			NumEndChapter.Maximum = num;
			
		}
		/// <summary>
		/// 끝 범위의 책이 시작 범위 보다 앞서지 않게 한다.
		/// </summary>
		private void CmbEndBook_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (CmbStartBook.SelectedIndex > CmbEndBook.SelectedIndex)
				CmbStartBook.SelectedIndex = CmbEndBook.SelectedIndex;
		}

		private void NumStartChapter_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
		{
			conn = new SQLiteConnection("Data Source=RevisedKorBible.db;Verseion=3");
			conn.Open();
			command = conn.CreateCommand();
			command.CommandText = String.Format("select count(*) from {0} where Chapter={1}", CmbStartBook.Text, ((Xceed.Wpf.Toolkit.IntegerUpDown)sender).Value);
			rdr = command.ExecuteReader();
			while (rdr.Read())
			{
				int num = rdr.GetInt32(0);
				NumStartPassage.Maximum = num;
			}
			rdr.Close();

			NumEndChapter.Minimum = NumStartChapter.Value;
		}

		private void NumStartPassage_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
		{
			if (NumStartChapter.Value == NumEndChapter.Value)
				NumEndPassage.Minimum = NumStartPassage.Value;
		}

		private void NumEndChapter_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
		{
			conn = new SQLiteConnection("Data Source=RevisedKorBible.db;Verseion=3");
			conn.Open();
			command = conn.CreateCommand();
			command.CommandText = String.Format("select count(*) from {0} where Chapter={1}", CmbStartBook.Text, ((Xceed.Wpf.Toolkit.IntegerUpDown)sender).Value);
			rdr = command.ExecuteReader();
			while (rdr.Read())
			{
				int num = rdr.GetInt32(0);
				NumEndPassage.Maximum = num;
			}
			rdr.Close();

			if (NumStartChapter.Value != NumEndChapter.Value)
				NumEndPassage.Minimum = 1;
		}

		private void NumEndPassage_ValueChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
		{

		}

		/// <summary>
		/// CbBirth체크박스가 체크되면 BirthList영역을 표시. 해제시 숨김 처리.
		/// </summary>
		private void CheckBox_CheckChanged(object sender, RoutedEventArgs e)
		{
			bool check = (CbBirth.IsChecked == true);
			if (check) BirthList.Visibility = Visibility.Visible;
			else BirthList.Visibility = Visibility.Hidden;
		}
		
		/// <summary>
		/// Deprecated: 말씀 범위에 해당하는 모든 구절을 반환.
		/// </summary>
		/// <param name="start">시작 범위를 나타내는 BVS클래스</param>
		/// <param name="passagesNum">전체 절 갯수</param>
		/// <returns>범위 구절을 담은 string LIst</returns>
		/// <remarks>구절 Resource의 각 line은 다음과 같은 형식으로 되어있다.
		/// <br/> {약어}{장}:{절}{Space}{구절}
		/// </remarks>
		private List<string> GetAllBibleVerse(BibleVerseSkeleton start, int passagesNum)
        {
			List<string> verses = new List<string>();
			
			string bookAbbr = abbr[books.IndexOf(start.book)];
			string verseAbbr = bookAbbr + start.chapter.ToString() + ":" + start.passage.ToString();

			Assembly _assembly;
			StreamReader _textStreamReader = null;
			try
			{
				_assembly = Assembly.GetExecutingAssembly();
				_textStreamReader = new StreamReader(_assembly.GetManifestResourceStream("SundayWorshipPPTMaker.RevisedKorBible.txt"));
			}
			catch
			{
				MessageBox.Show("Error accessing resources!");
				return verses;
			}
			finally
			{
				int index = 0;
				bool found = false;
				string line;
				while ((line = _textStreamReader.ReadLine()) != null)
				{
                    if (!found)
                    {
						if (verseAbbr == line.Split().ElementAt(0))
						{
							found = true;
							verses.Add(line.Substring(line.IndexOf(" ") + 1));
							index++;
						}
					}
					else if (found && index <= passagesNum)
					{
						verses.Add(line.Substring(line.IndexOf(" ") + 1));
						index++;
					}
					else break;

				}
			}
			return verses;
        }
		
		private List<Tuple<string,string>> GetAllBibleVerseDB(BibleVerseSkeleton start, BibleVerseSkeleton end)
        {
			List<Tuple<string,string>> verses = new List<Tuple<string, string>>();
			conn = new SQLiteConnection("Data Source=RevisedKorBible.db;Verseion=3");
			conn.Open();
			command = conn.CreateCommand();
			command.CommandText = String.Format("select * from {0} where rowid between ", start.book) +
				String.Format("(select rowid from {0} where Chapter={1} and Passage={2}) and (select rowid from {0} where Chapter={3} and Passage={4})",
				start.book, start.chapter, start.passage, end.chapter, end.passage);
			rdr = command.ExecuteReader();
            while (rdr.Read())
            {
				verses.Add(new Tuple<string, string>(rdr.GetInt32(0).ToString() + ":" + rdr.GetInt32(1).ToString(), rdr.GetString(2)));
            }
			rdr.Close();
			return verses;
		}

		/// <summary>
		/// PPT생성 작업 전 빠진 부분 체크.
		/// </summary>
		/// <returns>에러가 없으면 <c>false</c>, 있으면 <c>true</c></returns>
		private bool CheckErrorBeforeTask()
        {
			string errorMsg = Properties.Resources.TaskErrorCheckMessage;
			int errorCheckNum = 1;
			if (!Directory.Exists(TxtOutputFolder.Text))
			{
				errorMsg += "\n" + errorCheckNum++.ToString() + ". " + Properties.Resources.InvalidWorkingDirectory;
			}
			/*if (!File.Exists(settings.templateFileFullPath))
			{
				errorMsg += "\n" + errorCheckNum++.ToString() + ". " + Properties.Resources.InvalidPresentationTemplate;
			}*/
			if (SongList.Items.Count == 0)
			{
				errorMsg += "\n" + errorCheckNum++.ToString() + ". " + Properties.Resources.NoPraise;
			}
			if (!File.Exists(TxtPreachLocation.Text))
			{
				errorMsg += "\n" + errorCheckNum++.ToString() + ". " + Properties.Resources.InvalidPreachPPT;
			}
			if (ManualMode.IsChecked==true)
            {
				if (TxtPrayer.Text.Length == 0)
					errorMsg += "\n" + errorCheckNum++.ToString() + ". " + Properties.Resources.NoPrayer;
				if (TxtTitle.Text.Length == 0)
					errorMsg += "\n" + errorCheckNum++.ToString() + ". " + Properties.Resources.NoTitle;
            }
			if (TxtOutputFileName.Text.StartsWith(".pptx"))
			{
				errorMsg += "\n" + errorCheckNum++.ToString() + ". " + Properties.Resources.InvalidFileName;
			}
			if (!TxtOutputFileName.Text.EndsWith(".pptx"))
            {
				TxtOutputFileName.Text += ".pptx";
            }

			if (errorCheckNum == 1) return false;
			else
			{
				MessageBox.Show(errorMsg);
				return true;
			}
		}

		private void AddRandomCover(
			ref PowerPoint.Application pptApp,
			ref PowerPoint.Presentations pptPres, 
			ref PowerPoint.Presentation presentation)
        {
			Random rnd = new Random();
			int coverIndex = rnd.Next(1, 23);
			PowerPoint.Presentation covers=pptPres.Open(AppDomain.CurrentDomain.BaseDirectory + "cover.pptx", 
				WithWindow: MsoTriState.msoFalse);
			covers.Slides[coverIndex].Copy();
			presentation.Windows[1].Activate();
			pptApp.CommandBars.ExecuteMso("PasteSourceFormatting");
			presentation.Slides[1].Delete();
			covers.Close();

			presentation.Slides[1].Shapes[1].LockAspectRatio = MsoTriState.msoFalse;
			presentation.Slides[1].Shapes[1].Height = Utils.CMToPoint(Constants.CoverImageHeight);
			presentation.Slides[1].Shapes[1].Width = Utils.CMToPoint(Constants.CoverImageWidth);
			presentation.Slides[1].Shapes[1].Top = Utils.CMToPoint(Constants.CoverImageTop);
			presentation.Slides[1].Shapes[1].Left = Utils.CMToPoint(Constants.CoverImageLeft);

			presentation.Slides[1].Shapes[2].LockAspectRatio = MsoTriState.msoFalse;
			presentation.Slides[1].Shapes[2].Top = Utils.CMToPoint(Constants.CoverCommentTop);
			presentation.Slides[1].Shapes[2].Left = Utils.CMToPoint(Constants.CoverCommentLeft);


			presentation.Slides[1].Shapes[2].TextFrame2.TextRange.Lines[1, 1].Text = dt.ToString("yyyy.MM.dd");
			presentation.Slides[1].Shapes[2].TextFrame2.TextRange.Lines[1, 1].Font.Size = 40;
			presentation.Slides[1].Shapes[2].TextFrame2.TextRange.Lines[2, 1].Text = "XX주일";
			presentation.Slides[1].Shapes[2].TextFrame2.TextRange.Lines[2, 1].Font.Size = 48;

		}

		private void AddCreed(ref PowerPoint.Presentation presentation, ref int last_idx)
        {
			PowerPoint.Slide currentSlide;
			AddCutSlide(ref presentation, ref last_idx, 
				AppDomain.CurrentDomain.BaseDirectory+Properties.Resources.BGUriCross01,
				out currentSlide);

			currentSlide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,
				0, 0, settings.SlideSize.Width, settings.SlideSize.Height);
			var currentShape = currentSlide.Shapes[1];
			currentShape.Fill.ForeColor.RGB = 0;
			currentShape.Fill.Transparency = 0.45f;
			currentShape.Line.ForeColor.RGB = 0;

			currentSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
				Utils.CMToPoint(6.83f), Utils.CMToPoint(1.32f),
				Utils.CMToPoint(20.2f), Utils.CMToPoint(2.82f));
			currentShape = currentSlide.Shapes[2];
			SetTextEffectOptions(ref currentShape, "사도신경", Constants.FontKopubDotumBold, 60, TextEmphasis.Bold, 0xE8DEB7,
				MsoParagraphAlignment.msoAlignCenter);
			currentSlide.Shapes[2].TextFrame2.TextRange.Font.Glow.Color.RGB = 0xB94D7E;
			currentSlide.Shapes[2].TextFrame2.TextRange.Font.Glow.Radius = 18;
			currentSlide.Shapes[2].TextFrame2.TextRange.Font.Glow.Transparency = 0.6f;
			currentSlide.Shapes[2].TextFrame2.TextRange.Font.Spacing = 6;

			currentSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
				Utils.CMToPoint(0.53f), Utils.CMToPoint(4.52f),
				Utils.CMToPoint(32.8f), Utils.CMToPoint(13.2f));
			currentShape = currentSlide.Shapes[3];
			currentShape.LockAspectRatio = MsoTriState.msoFalse;
			SetTextEffectOptions(ref currentShape,
				"나는 전능하신 아버지 하나님, 천지의 창조주를 믿습니다.\n"+
				"나는 그의 유일하신 아들, 우리 주 예수 그리스도를 믿습니다.\n"+
				"그는 성령으로 잉태되어 동정녀 마리아에게서 나시고,\n"+
				"본디오 빌라도에게 고난을 받아 십자가에 못 박혀 죽으시고,\n"+
				"장사된 지 사흘 만에 죽은 자 가운데서 다시 살아나셨으며,",
				Constants.FontNanumSquareBold, 36, fontFillColor: 0xffffff,
				paragraphAlignment: MsoParagraphAlignment.msoAlignJustify,
				verticalAnchor: MsoVerticalAnchor.msoAnchorMiddle,
				autoSize: MsoAutoSize.msoAutoSizeNone, lineSpace: 1.5f);
			currentSlide.Shapes[3].TextFrame2.TextRange.Font.Glow.Color.RGB = 0xB94D7E;
			currentSlide.Shapes[3].TextFrame2.TextRange.Font.Glow.Radius = 5;
			currentSlide.Shapes[3].TextFrame2.TextRange.Font.Glow.Transparency = 0.6f;

			currentSlide.Duplicate();
			presentation.Slides[++last_idx].Shapes[3].TextFrame2.TextRange.Text =
				"하늘에 오르시어 전능하신 아버지 하나님 우편에 앉아 계시다가\n"+
				"거기로부터 살아있는 자와 죽은 자를 심판하러 오십니다.\n"+
				"나는 성령을 믿으며, 거룩한 공교회와 성도의 교제와\n"+
				"죄를 용서 받는 것과 몸의 부활과 영생을 믿습니다.\n아멘.";
		}

		private void AddLordsPrayer(ref PowerPoint.Presentation presentation, ref int last_idx)
        {
			PowerPoint.Slide currentSlide;
			AddCutSlide(ref presentation, ref last_idx,
				AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriCross02,
				out currentSlide);
			currentSlide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,
				0, 0, settings.SlideSize.Width, settings.SlideSize.Height);
			currentSlide.Shapes[1].Fill.ForeColor.RGB = 0;
			currentSlide.Shapes[1].Fill.Transparency = 0.45f;
			currentSlide.Shapes[1].Line.ForeColor.RGB = 0;

			currentSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
				Utils.CMToPoint(6.83f), Utils.CMToPoint(1.32f),
				Utils.CMToPoint(20.2f), Utils.CMToPoint(2.82f));
			var currentShape = currentSlide.Shapes[2];
			SetTextEffectOptions(ref currentShape, "주기도문", Constants.FontKopubDotumBold, 60,
				TextEmphasis.Bold, 0xE8DEB7, MsoParagraphAlignment.msoAlignCenter);
			currentShape.TextFrame2.TextRange.Font.Spacing = 6;
			currentShape.TextFrame2.TextRange.Font.Glow.Color.RGB = 0x34E0A8;
			currentShape.TextFrame2.TextRange.Font.Glow.Radius = 18;
			currentShape.TextFrame2.TextRange.Font.Glow.Transparency = 0.6f;

			currentSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
				Utils.CMToPoint(1.53f), Utils.CMToPoint(4.52f),
				Utils.CMToPoint(31.4f), Utils.CMToPoint(13.2f));
			currentShape = currentSlide.Shapes[3];
			SetTextEffectOptions(ref currentShape,
				"하늘에 계신 우리 아버지,\n아버지의 이름을 거룩하게 하시며\n"+
				"아버지의 나라가 오게 하시며,\n아버지의 뜻이 하늘에서와 같이 땅에서도 이루어지게 하소서.\n"+
				"오늘 우리에게 일용할 양식을 주시고,\n우리가 우리에게 잘못한 사람을 용서하여 준 것같이",
				Constants.FontNanumSquareBold, 36, fontFillColor: 0xffffff,
				paragraphAlignment:MsoParagraphAlignment.msoAlignJustify, verticalAnchor: MsoVerticalAnchor.msoAnchorMiddle,
				autoSize:MsoAutoSize.msoAutoSizeNone, lineSpace: 1.5f);
			currentShape.TextFrame2.TextRange.Font.Glow.Color.RGB = 0xF4C71D;
			currentShape.TextFrame2.TextRange.Font.Glow.Radius = 5;
			currentShape.TextFrame2.TextRange.Font.Glow.Transparency = 0.6f;
			currentSlide.Duplicate();
			presentation.Slides[++last_idx].Shapes[3].TextFrame2.TextRange.Text =
				"우리 죄를 용서하여 주시고,\n우리를 시험에 빠지지 않게 하시고\n"+
				"악에서 구하소서.\n나라와 권능과 영광이\n영원히 아버지의 것입니다.\n아멘.";
		}
		private void AddSongSlides(ref PowerPoint.Application pptApp,
			ref PowerPoint.Presentations pptPres,
			ref PowerPoint.Presentation presentation,
			ref int last_idx)
        {
			PowerPoint.Slide currentSlide;
			AddCutSlide(ref presentation, ref last_idx,
				AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriIntro,
				out currentSlide);
			currentSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
				Utils.CMToPoint(4.34f), Utils.CMToPoint(1.92f),
				Utils.CMToPoint(21.59f), Utils.CMToPoint(2.2f));
			var currentShape = currentSlide.Shapes[1];
			SetTextEffectOptions(ref currentShape, "경배와 찬양", "KT&G 상상제목 B", 80,
				verticalAnchor: MsoVerticalAnchor.msoAnchorMiddle, autoSize: MsoAutoSize.msoAutoSizeNone);

			for (int i =0; i < SongList.Items.Count; i++)
			{
				PowerPoint.Presentation item = pptPres.Open((string)SongList.Items[i], WithWindow: MsoTriState.msoFalse);
				item.Slides.Range().Copy();

				presentation.Windows[1].Activate();
				presentation.Windows[1].View.GotoSlide(last_idx);
				pptApp.CommandBars.ExecuteMso("PasteSourceFormatting");
				//찬양 제목 슬라이드
				currentSlide = presentation.Slides[++last_idx];
				currentSlide.Duplicate();
				currentSlide.Shapes.Range().Delete();
				AddSongTitle(ref currentSlide,i);

				last_idx += item.Slides.Count;
				item.Close();
			}
			//presentation.Slides[settings.SettingsPraiseEntry].Delete();
			
			AddCutSlide(ref presentation, ref last_idx,
				AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriCutImg01,
				out currentSlide);
		}
		private void AddSongTitle(ref PowerPoint.Slide currentSlide, int index)
        {
			currentSlide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle,
					Utils.CMToPoint(7.13f), Utils.CMToPoint(8.15f), Utils.CMToPoint(19.6f), Utils.CMToPoint(2.79f));
			var currentShape = currentSlide.Shapes[1];
			currentShape.Line.Visible = MsoTriState.msoFalse;
			AdjustShadow(ref currentShape, 0, 0.4f, 100, 28, 45, 16);
			currentShape.Duplicate();
			
			currentShape = currentSlide.Shapes[2];
			currentShape.Left = Utils.CMToPoint(7.13f);
			currentShape.Top = Utils.CMToPoint(8.15f);
			AdjustShadow(ref currentShape, 0xffffff, 0.3f, 100, 28, 225, 16);
			currentShape.Fill.TwoColorGradient(MsoGradientStyle.msoGradientHorizontal, 1);
			var gradStops = currentShape.Fill.GradientStops;
			gradStops[1].Color.RGB = 0xccc6af;
			gradStops[2].Color.RGB = 0xf8f1de;
			gradStops.Insert(0xeceae0, 0.22f);
			gradStops.Insert(0xf9f4f3, 0.56f);

			SetTextEffectOptions(ref currentShape, System.IO.Path.GetFileNameWithoutExtension(SongList.Items[index] as string),
				Constants.FontSongTitle, 40, TextEmphasis.Shadow, 0,
				MsoParagraphAlignment.msoAlignCenter, MsoVerticalAnchor.msoAnchorMiddle,
				MsoAutoSize.msoAutoSizeNone, shadowOptions: new ShadowOptions(transparency: 0.6f, distance: 3));
		}

		private void AddSlide(ref PowerPoint.Presentation presentation,ref int last_idx,out PowerPoint.Slide currentSlide)
        {
			presentation.Windows[1].Activate();
			presentation.Windows[1].View.GotoSlide(last_idx);
			presentation.Slides.AddSlide(++last_idx, presentation.SlideMaster.CustomLayouts[7]);
			currentSlide = presentation.Slides[last_idx];
        }

		private void AddCutSlide(ref PowerPoint.Presentation presentation, ref int last_idx, string path, out PowerPoint.Slide currentSlide)
        {
			AddSlide(ref presentation, ref last_idx, out currentSlide);
			currentSlide.FollowMasterBackground = MsoTriState.msoFalse;
			currentSlide.Background.Fill.UserPicture(path);
        }

		private void SetTextEffectOptions(ref PowerPoint.Shape currentShape, string text, string fontName = "굴림", float fontSize = 18,
			TextEmphasis emphasis = TextEmphasis.None, int fontFillColor = 0,
			MsoParagraphAlignment paragraphAlignment = MsoParagraphAlignment.msoAlignLeft,
			MsoVerticalAnchor verticalAnchor = MsoVerticalAnchor.msoAnchorTop,
			MsoAutoSize autoSize = MsoAutoSize.msoAutoSizeShapeToFitText,
			float lineSpace = 1, MsoTriState fontLineVisible = MsoTriState.msoFalse,
			int fontLineColor = 0, float fontLineWeight = 1, ShadowOptions shadowOptions=null)
        {
			currentShape.TextFrame2.AutoSize = autoSize;
			currentShape.TextFrame2.TextRange.Text = text;
			currentShape.TextFrame2.TextRange.Font.NameFarEast = fontName;
			currentShape.TextFrame2.TextRange.Font.NameAscii = "(한글 글꼴 사용)";
			currentShape.TextFrame2.TextRange.Font.Size = fontSize;
			if ((emphasis & TextEmphasis.Bold) == TextEmphasis.Bold) currentShape.TextFrame2.TextRange.Font.Bold = MsoTriState.msoTrue;
			else currentShape.TextFrame2.TextRange.Font.Bold = MsoTriState.msoFalse;
			if ((emphasis & TextEmphasis.Italic) == TextEmphasis.Italic) currentShape.TextFrame2.TextRange.Font.Italic = MsoTriState.msoTrue;
			else currentShape.TextFrame2.TextRange.Font.Italic = MsoTriState.msoFalse;
			if ((emphasis & TextEmphasis.UnderLine) == TextEmphasis.UnderLine) currentShape.TextFrame2.TextRange.Font.UnderlineStyle = MsoTextUnderlineType.msoUnderlineSingleLine;
			else currentShape.TextFrame2.TextRange.Font.UnderlineStyle = MsoTextUnderlineType.msoNoUnderline;
			if ((emphasis & TextEmphasis.StrikeThrough) == TextEmphasis.StrikeThrough) currentShape.TextFrame2.TextRange.Font.StrikeThrough = MsoTriState.msoTrue;
			else currentShape.TextFrame2.TextRange.Font.StrikeThrough = MsoTriState.msoFalse;
			if ((emphasis & TextEmphasis.Shadow) == TextEmphasis.Shadow)
			{
				var textFrame = currentShape.TextFrame2;
				if (shadowOptions == null) shadowOptions = new ShadowOptions();
				AdjustShadow(ref textFrame, shadowOptions.Color, shadowOptions.Transparency, shadowOptions.Size,
					shadowOptions.Blur, shadowOptions.Angle, shadowOptions.Distance, shadowOptions.Style);
			}

			currentShape.TextFrame2.TextRange.ParagraphFormat.Alignment = paragraphAlignment;
			currentShape.TextFrame2.VerticalAnchor = verticalAnchor;
			currentShape.TextFrame2.TextRange.ParagraphFormat.SpaceWithin = lineSpace;
			currentShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = fontFillColor;
			currentShape.TextFrame2.TextRange.Font.Line.Visible = fontLineVisible;
			if (fontLineVisible == MsoTriState.msoTrue)
            {
				currentShape.TextFrame2.TextRange.Font.Line.ForeColor.RGB = fontLineColor;
				currentShape.TextFrame2.TextRange.Font.Line.Weight = fontLineWeight;
            }

        }

		private void EditPrayer(ref PowerPoint.Presentation presentation, ref int last_idx)
        {
			PowerPoint.Slide currentSlide;
			AddCutSlide(ref presentation, ref last_idx, 
				AppDomain.CurrentDomain.BaseDirectory+Properties.Resources.BGUriPray01,
				out currentSlide);
			currentSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
				Utils.CMToPoint(3.6f), Utils.CMToPoint(2.12f), Utils.CMToPoint(20), Utils.CMToPoint(4.36f));
			var currentShape = currentSlide.Shapes[1];
			SetTextEffectOptions(ref currentShape, "대표기도", Constants.FontSequenceTitle, 96,
				TextEmphasis.Bold | TextEmphasis.Shadow, 0xffffff);
			currentSlide.Shapes[1].TextFrame2.TextRange.Font.Line.Visible = MsoTriState.msoTrue;
			currentSlide.Shapes[1].TextFrame2.TextRange.Font.Line.ForeColor.RGB = 0x074898;
			currentSlide.Shapes[1].TextFrame2.TextRange.Font.Line.Weight = 1.5f;


			currentSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
				Utils.CMToPoint(6.38f), Utils.CMToPoint(5.33f), Utils.CMToPoint(21.1f), Utils.CMToPoint(3.2f));
			currentShape = currentSlide.Shapes[2];
			SetTextEffectOptions(ref currentShape, TxtPrayer.Text, Constants.FontSequenceTitle, 60,
				TextEmphasis.None, 0xffffff, lineSpace:1.5f);
			currentSlide.Shapes[2].TextFrame2.TextRange.Font.Line.Visible = MsoTriState.msoTrue;
			currentSlide.Shapes[2].TextFrame2.TextRange.Font.Line.ForeColor.RGB = 0x404040;
			currentSlide.Shapes[2].TextFrame2.TextRange.Font.Line.Weight = 1.25f;
		}

		private void EditBibleVerseCover(ref PowerPoint.Shape currentShape,out string verseString)
        {
			verseString = CmbStartBook.Text + " ";
			verseString += NumStartChapter.Text;
			if (CmbStartBook.Text == "시편")
				verseString += "편 ";
			else
				verseString += "장 ";

			//Case Multi passages through multiple chapters
			if (NumStartChapter.Text != NumEndChapter.Text)
            {
				verseString += NumStartChapter.Text + ":" + NumStartPassage.Text + "-" 
					+ NumEndChapter.Text + ":" + NumEndPassage.Text;
            }
			//Case Single Passage
			else if (NumStartPassage.Text == NumEndPassage.Text)
            {
				verseString += NumStartPassage.Text + "절";
            }
			//Default
            else
            {
				verseString += NumStartPassage.Text + "-" + NumEndPassage.Text + "절";
            }

			currentShape.TextFrame.TextRange.Text = verseString;
		}

		private void AddBibleVerseSlides(ref PowerPoint.Presentation presentation, ref int last_idx, string verseString)
        {
			PowerPoint.Slide currentSlide;
			AddCutSlide(ref presentation, ref last_idx,
				AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriDefault, out currentSlide);
			currentSlide.Shapes.AddPicture2(AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriBibleTextLight,
				MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0);
			currentSlide.Shapes.AddPicture2(AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.ImgUriBibleIcon,
				MsoTriState.msoFalse, MsoTriState.msoTrue,
				Utils.CMToPoint(3.19f), Utils.CMToPoint(1.17f), Utils.CMToPoint(1.39f), Utils.CMToPoint(1.39f));
			currentSlide.Shapes.AddPicture2(AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.ImgUriSeal,
				MsoTriState.msoFalse, MsoTriState.msoTrue,
				Utils.CMToPoint(15.69f), Utils.CMToPoint(17.72f), Utils.CMToPoint(2.54f), Utils.CMToPoint(0.87f));
			currentSlide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle,
				Utils.CMToPoint(2.53f), Utils.CMToPoint(2.99f), Utils.CMToPoint(29.14f), Utils.CMToPoint(14.22f));
			currentSlide.Shapes[4].Fill.ForeColor.RGB = 0xf3fbff;
			currentSlide.Shapes[4].Fill.Transparency = 0.6f;
			currentSlide.Shapes[4].Line.Visible = MsoTriState.msoFalse;
			currentSlide.Shapes[4].Adjustments[1] = 0.07707f;

			currentSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
				Utils.CMToPoint(4.58f), Utils.CMToPoint(1.27f), Utils.CMToPoint(21.16f), Utils.CMToPoint(1.35f));
			var currentShape = currentSlide.Shapes[5];
			SetTextEffectOptions(ref currentShape, verseString.Replace('-', '~'), Constants.FontNanumSquareBold, 36, 0, 0xe0dedc,
				verticalAnchor: MsoVerticalAnchor.msoAnchorMiddle, autoSize: MsoAutoSize.msoAutoSizeNone,
				shadowOptions: new ShadowOptions(transparency: 0.5f, angle: 90, distance: 1));
			currentShape.TextFrame2.TextRange.Font.Spacing = 1.5f;
			currentSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
				Utils.CMToPoint(3.53f), Utils.CMToPoint(3.67f), Utils.CMToPoint(27.1f), Utils.CMToPoint(6.67f));
			currentShape = currentSlide.Shapes[6];
			SetTextEffectOptions(ref currentShape, "", Constants.FontNanumSquareExBold, 36,
				autoSize: MsoAutoSize.msoAutoSizeNone, lineSpace: 1.5f);
			currentShape.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent = Utils.CMToPoint(-2.06f);
			currentShape.TextFrame2.TextRange.ParagraphFormat.LeftIndent = Utils.CMToPoint(2.06f);


			//절 수
			List <Tuple<string, string>> verses;
			if (ManualMode.IsChecked == true)
			{
				verses = GetAllBibleVerseDB(
					new BibleVerseSkeleton(CmbStartBook.Text, int.Parse(NumStartChapter.Text), int.Parse(NumStartPassage.Text)),
					new BibleVerseSkeleton(CmbStartBook.Text, int.Parse(NumEndChapter.Text), int.Parse(NumEndPassage.Text))
					);
			}
			else verses = GetAllBibleVerseDB(jubo.BVSStart, jubo.BVSEnd);
			int passagesNum = verses.Count();

			for (int i = 0; i < passagesNum - 1; i++)
			{
				currentSlide.Duplicate();
			}
			for (int i = 0; i < passagesNum; i++)
			{
				int slideIndex = currentSlide.SlideIndex;
				string[] va = verses[i].Item1.Split(':');
				if (int.Parse(va[0])!=int.Parse(NumStartChapter.Text))
				{
					presentation.Slides[slideIndex + i].Shapes[6].TextFrame.TextRange.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone;
					presentation.Slides[slideIndex + i].Shapes[6].TextFrame.TextRange.Text = va[1] + ". " + verses[i].Item2;
				}
				else
				{
					presentation.Slides[slideIndex + i].Shapes[6].TextFrame.TextRange.Text = verses[i].Item2;
					presentation.Slides[slideIndex + i].Shapes[6].TextFrame.TextRange.ParagraphFormat.Bullet.StartValue = int.Parse(va[1]);
				}
			}
			last_idx += passagesNum - 1;
		}
		
		private void AddBibleCover(ref PowerPoint.Presentation presentation,
			ref int last_idx, out string verseString)
        {
			PowerPoint.Slide currentSlide;
			AddCutSlide(ref presentation, ref last_idx,
				AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriDefault, out currentSlide);
			currentSlide.Shapes.AddPicture2(AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriBibleCover,
				MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0);

			currentSlide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle,
				Utils.CMToPoint(4.12f), Utils.CMToPoint(3.75f), Utils.CMToPoint(25.64f), Utils.CMToPoint(10.95f));
			currentSlide.Shapes[2].Fill.ForeColor.RGB = 0xf3fbff;
			currentSlide.Shapes[2].Fill.Transparency = 0.6f;
			currentSlide.Shapes[2].Line.Visible = MsoTriState.msoFalse;
			currentSlide.Shapes[2].Adjustments[1] = 0.07707f;

			currentSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
				Utils.CMToPoint(3.51f), Utils.CMToPoint(5.74f), Utils.CMToPoint(26.86f), Utils.CMToPoint(3.28f));
			var currentShape = currentSlide.Shapes[3];
			SetTextEffectOptions(ref currentShape, "성경봉독", "함초롬돋움", 72,
				TextEmphasis.Bold | TextEmphasis.Shadow, 0x3b3734,
				MsoParagraphAlignment.msoAlignCenter, MsoVerticalAnchor.msoAnchorMiddle,
				shadowOptions: new ShadowOptions(transparency: 0.5f, angle: 90, distance: 1));
			currentShape.TextFrame2.MarginLeft = Utils.CMToPoint(0.1f);
			currentShape.TextFrame2.MarginRight = Utils.CMToPoint(0.1f);
			currentShape.TextFrame2.MarginTop = Utils.CMToPoint(0.1f);
			currentShape.TextFrame2.MarginBottom = Utils.CMToPoint(0.1f);
			currentShape.Duplicate();
			currentShape.TextFrame2.TextRange.Font.Spacing = 3;

			currentShape = currentSlide.Shapes[4];
			SetTextEffectOptions(ref currentShape, "", Constants.FontNanumSquareExBold, 40, 0, 0x3b3734,
				MsoParagraphAlignment.msoAlignCenter, MsoVerticalAnchor.msoAnchorMiddle,
				shadowOptions: new ShadowOptions(transparency: 0.5f, angle: 90, distance: 1));
			currentShape.Left = Utils.CMToPoint(3.51f);
			currentShape.Top = Utils.CMToPoint(10.91f);

			EditBibleVerseCover(ref currentShape, out verseString);
			AddSeal(ref currentSlide);
		}

		private void AddSeal(ref PowerPoint.Slide currentSlide)
        {
			currentSlide.Shapes.AddPicture2(
				AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.ImgUriSeal,
				MsoTriState.msoFalse, MsoTriState.msoTrue,
				Utils.CMToPoint(15.63f), Utils.CMToPoint(17.72f), Utils.CMToPoint(2.61f), Utils.CMToPoint(0.87f));
        }
		private void AddVideo(ref PowerPoint.Presentation presentation, ref int last_idx)
        {
			if (System.IO.File.Exists(TxtVidLocation.Text))
            {
				PowerPoint.Slide currentSlide;
				AddSlide(ref presentation, ref last_idx, out currentSlide);
				currentSlide.Shapes.AddMediaObject2(TxtVidLocation.Text);
            }
		}

		private void AddPreachSlides(ref PowerPoint.Application pptApp,
			ref PowerPoint.Presentations pptPres,
			ref PowerPoint.Presentation presentation,
			ref int last_idx)
        {
			AddPreachCover(ref presentation, ref last_idx);

			PowerPoint.Presentation preachPPT = pptPres.Open(TxtPreachLocation.Text, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);
			preachPPT.Slides.Range().Copy();

			presentation.Windows[1].Activate();
			presentation.Windows[1].View.GotoSlide(last_idx);
			pptApp.CommandBars.ExecuteMso("PasteSourceFormatting");

			last_idx += preachPPT.Slides.Count;
			preachPPT.Close();
		}
		
		private void AddPreachCover(ref PowerPoint.Presentation presentation,
			ref int last_idx)
        {
			PowerPoint.Slide currentSlide;
			AddCutSlide(ref presentation, ref last_idx,
				AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriPreach,
				out currentSlide);
			
			currentSlide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,
				0, Utils.CMToPoint(2.54f),
				settings.SlideSize.Width, Utils.CMToPoint(Constants.PaneHeight));
			var currentShape = currentSlide.Shapes[1];
			currentShape.Fill.ForeColor.RGB = 0;
			currentShape.Fill.Transparency = 0.45f;
			currentShape.Line.Visible = MsoTriState.msoFalse;

			SetTextEffectOptions(ref currentShape, "[말씀선포]\n" + TxtTitle.Text + "\n\n\n" + "유주원 전도사",
				"맑은 고딕", fontFillColor: 0xffffff, paragraphAlignment: MsoParagraphAlignment.msoAlignCenter, 
				verticalAnchor:MsoVerticalAnchor.msoAnchorMiddle, autoSize:MsoAutoSize.msoAutoSizeNone);
			currentSlide.Shapes[1].TextFrame2.TextRange.Lines[1, 1].Font.Size = 32;
			currentSlide.Shapes[1].TextFrame2.TextRange.Lines[1, 1].Font.Fill.ForeColor.RGB = 0x00c0ff;
			currentSlide.Shapes[1].TextFrame2.TextRange.Lines[2, 1].Font.Size = 54;
			currentSlide.Shapes[1].TextFrame2.TextRange.Lines[3, 1].Font.Size = 20;
			currentSlide.Shapes[1].TextFrame2.TextRange.Lines[3, 1].Font.Size = 20;
			currentSlide.Shapes[1].TextFrame2.TextRange.Lines[5, 1].Font.Size = 48;
		}

		private void AddAfterPreach(ref PowerPoint.Presentation presentation,
			ref int last_idx)
        {
			PowerPoint.Slide currentSlide;
			AddCutSlide(ref presentation, ref last_idx,
				AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriPray02, out currentSlide);
			currentSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
				Utils.CMToPoint(9.54f), Utils.CMToPoint(3.72f),
				Utils.CMToPoint(14.16f), Utils.CMToPoint(4.02f));
			var currentShape = currentSlide.Shapes[1];
			SetTextEffectOptions(ref currentShape, "결단 기도", Constants.FontSequenceTitle, 88,
				TextEmphasis.Bold | TextEmphasis.Shadow, 0x262626, MsoParagraphAlignment.msoAlignCenter,
				shadowOptions: new ShadowOptions(color: 0xd57a9f, transparency: 0.55f, blur: 6.93f, angle: 84, distance: 4));

			AddCutSlide(ref presentation, ref last_idx,
				AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriDedication, out currentSlide);
			currentSlide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,
				0, Utils.CMToPoint(3.32f), settings.SlideSize.Width, Utils.CMToPoint(5.2f));
			currentShape = currentSlide.Shapes[1];
			currentShape.Fill.ForeColor.RGB = 0;
			currentShape.Fill.Transparency = 0.26f;
			currentShape.Line.Visible = MsoTriState.msoFalse;
			SetTextEffectOptions(ref currentShape, "봉헌", Constants.FontSequenceTitle, 80,
				TextEmphasis.Bold | TextEmphasis.Shadow, 0xffffff,
				MsoParagraphAlignment.msoAlignCenter, MsoVerticalAnchor.msoAnchorMiddle,
				MsoAutoSize.msoAutoSizeNone, fontLineVisible: MsoTriState.msoTrue,
				fontLineColor: 0x074898, fontLineWeight: 1.25f,
				shadowOptions: new ShadowOptions(transparency: 0.7f, angle: 90, distance: 1.8f));
			currentSlide.Duplicate();
			presentation.Slides[last_idx+1].Shapes[1].TextFrame2.TextRange.InsertAfter("기도");

			AddCutSlide(ref presentation, ref last_idx,
				AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriDSong, out currentSlide);
			currentSlide.Duplicate();
			currentSlide.Duplicate();
			for(int i = 0; i < 3; i++)
            {
				presentation.Slides[last_idx + i].Shapes.AddPicture2(
					AppDomain.CurrentDomain.BaseDirectory + "/Resources/dsong0" + (i + 1).ToString() + ".png",
					MsoTriState.msoFalse, MsoTriState.msoTrue, Utils.CMToPoint(4.23f), 0);
				presentation.Slides[last_idx + i].Shapes[1].LockAspectRatio = MsoTriState.msoFalse;
			}

			last_idx += 3;
		}

		private void AddAdvSlides(ref PowerPoint.Presentation presentation, ref int last_idx)
        {
			PowerPoint.Slide currentSlide;
			AddCutSlide(ref presentation, ref last_idx,
				AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriCutImg02, out currentSlide);

			AddSlide(ref presentation, ref last_idx, out currentSlide);
			MakeAdvTemplate(ref currentSlide, SlideContentsType.Cover);

			AddSlide(ref presentation, ref last_idx, out currentSlide);
			MakeAdvTemplate(ref currentSlide, SlideContentsType.Main);
			currentSlide.Duplicate();
			currentSlide=presentation.Slides[++last_idx];
			currentSlide.Shapes[2].TextFrame2.TextRange.Text = "사랑의 편지 신청 받고 있습니다!";
			currentSlide.Shapes[2].TextFrame2.TextRange.Words[1, 2].Font.Fill.ForeColor.RGB = 0x0000ff;

			EditBirthDaySlides(ref presentation, ref last_idx);
			AddCutSlide(ref presentation, ref last_idx,
				AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriDSong, out currentSlide);
			currentSlide.Duplicate();
			for (int i = 0; i < 2; i++)
			{
				presentation.Slides[last_idx + i].Shapes.AddPicture2(
					AppDomain.CurrentDomain.BaseDirectory + "/Resources/bsong0" + (i + 1).ToString() + ".png",
					MsoTriState.msoFalse, MsoTriState.msoTrue, Utils.CMToPoint(4.23f), 0);
				presentation.Slides[last_idx + i].Shapes[1].LockAspectRatio = MsoTriState.msoFalse;
			}
			last_idx++;

			//추가 광고


			AddCutSlide(ref presentation, ref last_idx,
				AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriCutImg03, out currentSlide);
		}

		private void MakeAdvTemplate(ref PowerPoint.Slide currentSlide, SlideContentsType slideContentsType)
        {
			//커버
			currentSlide.FollowMasterBackground = MsoTriState.msoFalse;
			currentSlide.Background.Fill.TwoColorGradient(MsoGradientStyle.msoGradientHorizontal, 1);
			var gradStops = currentSlide.Background.Fill.GradientStops;
			gradStops[1].Color.RGB = 0x99e6ff;
			gradStops[2].Color.RGB = 0x99e6ff;
			gradStops.Insert(0xccf2ff, 0.43f);
			gradStops.Insert(0xccf2ff, 0.60f);

			if (slideContentsType == SlideContentsType.Cover)
            {
				currentSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
					Utils.CMToPoint(9.24f), Utils.CMToPoint(7.49f), Utils.CMToPoint(15.38f), Utils.CMToPoint(4.02f));
				var currentShape = currentSlide.Shapes[1];
				SetTextEffectOptions(ref currentShape, "광      고", Constants.FontNanumSquareExBold, 88,
					TextEmphasis.Bold | TextEmphasis.Shadow, 0, MsoParagraphAlignment.msoAlignCenter);
				currentShape.TextFrame2.TextRange.Font.Spacing = 1.8f;
				currentShape.TextFrame2.TextRange.Font.Fill.TwoColorGradient(MsoGradientStyle.msoGradientVertical, 1);
				var fontGradStops = currentShape.TextFrame2.TextRange.Font.Fill.GradientStops;
				fontGradStops[1].Color.RGB = 0x629bb8;
				fontGradStops[2].Color.RGB = 0x3E628c;
            }
            else
            {
				currentSlide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle,
					Utils.CMToPoint(1.83f), Utils.CMToPoint(1.12f), Utils.CMToPoint(30.2f), Utils.CMToPoint(8));
				var currentShape = currentSlide.Shapes[1];
				currentShape.Adjustments[1] = Constants.RoundedRectangleRadius;
				currentShape.Line.Visible = MsoTriState.msoFalse;
				AdjustShadow(ref currentShape, 0, 0.6f, 100, 30, 45, 16);
				currentShape.Duplicate();

				currentShape = currentSlide.Shapes[2];
				currentShape.Left = Utils.CMToPoint(1.83f);
				currentShape.Top = Utils.CMToPoint(1.12f);
				currentShape.Fill.TwoColorGradient(MsoGradientStyle.msoGradientDiagonalUp, 1);
				var shapeGradStops = currentShape.Fill.GradientStops;
				shapeGradStops[1].Color.RGB = 0x99E6ff;
				shapeGradStops[2].Color.RGB = 0xf3d3ea;
				shapeGradStops.Insert(0xccf2ff, 0.39f);
				shapeGradStops.Insert(0xfffffa, 0.64f);
				AdjustShadow(ref currentShape, 0xffffff, 0.45f, 100, 30, 225, 16);
				SetTextEffectOptions(ref currentShape, "생일인 친구, 새 친구 모두 환영합니다!", Constants.FontNanumSquareBold, 50,
					verticalAnchor: MsoVerticalAnchor.msoAnchorMiddle, autoSize: MsoAutoSize.msoAutoSizeNone);
			}
		}

		private void AdjustShadow(ref PowerPoint.TextFrame2 textFrame,
			int color = 0, float transparency = 0.57f, float size = 100, float blur = 3, double angle = 45, float distance = 3,
			MsoShadowStyle shadowStyle = MsoShadowStyle.msoShadowStyleOuterShadow)
        {
			textFrame.TextRange.Font.Shadow.Style = shadowStyle;
			textFrame.TextRange.Font.Shadow.ForeColor.RGB = color;
			textFrame.TextRange.Font.Shadow.Transparency = transparency;
			textFrame.TextRange.Font.Shadow.Size = size;
			textFrame.TextRange.Font.Shadow.Blur = blur;
			textFrame.TextRange.Font.Shadow.OffsetX = (float)(distance * Math.Cos(angle));
			textFrame.TextRange.Font.Shadow.OffsetY = (float)(distance * Math.Sin(angle));
		}

		private void AdjustShadow(ref PowerPoint.Shape shape,
			int color = 0, float transparency = 0.6f, float size = 100, float blur = 4, double angle = 45, float distance = 3,
			MsoShadowStyle shadowStyle = MsoShadowStyle.msoShadowStyleOuterShadow)
		{
			shape.Shadow.Style = shadowStyle;
			shape.Shadow.ForeColor.RGB = color;
			shape.Shadow.Transparency = transparency;
			shape.Shadow.Size = size;
			shape.Shadow.Blur = blur;
			shape.Shadow.OffsetX = (float)(distance * Math.Cos(angle));
			shape.Shadow.OffsetY = (float)(distance * Math.Sin(angle));
		}

		private void EditBirthDaySlides(ref PowerPoint.Presentation presentation, ref int last_idx)
		{
			if (CbBirth.IsChecked == true)
			{
				PowerPoint.Slide currentSlide;
				AddCutSlide(ref presentation, ref last_idx,
					AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriBirthday01, out currentSlide);
				AddCutSlide(ref presentation, ref last_idx,
					AppDomain.CurrentDomain.BaseDirectory + Properties.Resources.BGUriBirthday02, out currentSlide);

				//생일자 명단 입력
				currentSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
					Utils.CMToPoint(3.73f), Utils.CMToPoint(1.32f), Utils.CMToPoint(26.2f), Utils.CMToPoint(2.57f));
				string titleText;
				if (Utils.IsLastSundayOfMonth())
					titleText = "🎉 " + DateTime.Now.Month.ToString() + "월 생일";
				else
					titleText = "🎉 " + "이번 주 생일자";

				var currentShapes = currentSlide.Shapes[1];
				SetTextEffectOptions(ref currentShapes, "titleText", Constants.FontNanumSquareExBold, 54, paragraphAlignment: MsoParagraphAlignment.msoAlignCenter);

				currentSlide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
					Utils.CMToPoint(3.93f), Utils.CMToPoint(6.33f), Utils.CMToPoint(26.6f), Utils.CMToPoint(2.64f));
				currentShapes = currentSlide.Shapes[2];
				SetTextEffectOptions(ref currentShapes, TxtBirthList.Text.Replace(", ", "\n"),
					Constants.FontNanumSquareBold, 44, paragraphAlignment: MsoParagraphAlignment.msoAlignCenter,
					verticalAnchor: MsoVerticalAnchor.msoAnchorMiddle,
					autoSize: MsoAutoSize.msoAutoSizeNone);
			}
		}

		private void SavePPT(ref PowerPoint.Presentation presentation, out string tempFilePath, out string finalFilePath)
        {
			string fileName = @"\" + TxtOutputFileName.Text;
			if (!Directory.Exists(TxtOutputFolder.Text))
			{
				//tempFilePath=Documents
			}
			tempFilePath = TxtOutputFolder.Text + @"\new";
			finalFilePath = TxtOutputFolder.Text + fileName;

			if (!Directory.Exists(finalFilePath))
			{
				File.Delete(finalFilePath);
			}
			presentation.Export(tempFilePath, "pptx");
			presentation.Close();
		}

		private void OpenFinalPPT(ref PowerPoint.Presentations pptPres, string tempFilePath, string finalFilePath)
        {
			//Open the created Presentation
			System.IO.File.Move(tempFilePath + ".pptx", finalFilePath);
			pptPres.Open(finalFilePath);
		}

		/// <summary>
		/// 새 PPT를 작성한다. PowerPoint가 실행되고 모든 작업을 마친 파일을 연다. 저장 위치는 작업폴더와 같음.
		/// </summary>
		private void StartTask(object sender, RoutedEventArgs e)
		{
			if (CheckErrorBeforeTask()) return;

			PowerPoint.Application pptApp = new PowerPoint.Application();
			PowerPoint.Presentations pptPres = pptApp.Presentations;
			PowerPoint.Presentation presentation = //pptPres.Open(settings.templateFileFullPath);
				pptPres.Add();
			presentation.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeCustom; 
			presentation.PageSetup.SlideWidth = Utils.CMToPoint(33.867f);
			presentation.PageSetup.SlideHeight = Utils.CMToPoint(19.05f);

			PowerPoint.CustomLayout pcl = presentation.SlideMaster.CustomLayouts[7];
			presentation.Slides.AddSlide(1, pcl);

			int lastSlide = 1;
			AddRandomCover(ref pptApp, ref pptPres, ref presentation);
			AddCreed(ref presentation,ref lastSlide);
			AddSongSlides(ref pptApp, ref pptPres, ref presentation, ref lastSlide);
			EditPrayer(ref presentation, ref lastSlide);
			string argVerseString;
			AddBibleCover(ref presentation, ref lastSlide, out argVerseString);
			AddBibleVerseSlides(ref presentation, ref lastSlide, argVerseString);
			AddVideo(ref presentation, ref lastSlide);
			AddPreachSlides(ref pptApp, ref pptPres, ref presentation, ref lastSlide);

			AddAfterPreach(ref presentation, ref lastSlide);
			AddLordsPrayer(ref presentation, ref lastSlide);
			AddAdvSlides(ref presentation, ref lastSlide);

			
			string argTempFilePath, argFinalFilePath;
			SavePPT(ref presentation, out argTempFilePath, out argFinalFilePath);
			OpenFinalPPT(ref pptPres, argTempFilePath, argFinalFilePath);
		}
		
		/// <summary>
		/// Settings Window를 표시한다.
		/// </summary>
        private void BtnShowSettings_Click(object sender, RoutedEventArgs e)
        {
			settings.ShowDialog();
        }
		/// <summary>
		/// 프로그램 종료 전 Settings Window를 닫는다.
		/// </summary>
        private void DisposeSettingWindow(object sender, System.ComponentModel.CancelEventArgs e)
        {
			settings.Close();
        }

        private void BtnShowHelp_Click(object sender, RoutedEventArgs e)
        {
			HelpPopup.IsOpen = true;
        }

        private void LinkUri_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
			Process.Start(
				new ProcessStartInfo(e.Uri.AbsoluteUri){ UseShellExecute=true,}
			);
			e.Handled = true;
        }

        private void ManualMode_Checked(object sender, RoutedEventArgs e)
        {
			((CheckBox)sender).IsEnabled = true;
        }

        private void ManualMode_Unchecked(object sender, RoutedEventArgs e)
        {
			((CheckBox)sender).IsEnabled = false;
        }
    }
}
