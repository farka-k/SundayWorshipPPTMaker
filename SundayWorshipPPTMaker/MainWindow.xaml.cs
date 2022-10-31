using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Text.RegularExpressions;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;

using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using HwpObjectLib;
using System.Data.SQLite;

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
		/// 주보 파일에 대한 Path와 추출 데이터.
		/// </summary>
		public class Jubo
		{
			/// <summary>주보 파일 이름</summary>
			public string FileName { get; set; }
			/// <summary>디렉토리명</summary>
			public string Directory { get; set; }
			/// <summary>디렉토리+파일이름</summary>
			public string FullPath { get; set; }
			
			/// <summary>말씀 시작 범위</summary>
			public BibleVerseSkeleton BVSStart { get; set; }
			/// <summary>말씀 끝 범위</summary>
			public BibleVerseSkeleton BVSEnd { get; set; }
			/// <summary>대표기도자</summary>
			public string PrayerName { get; set; }
			/// <summary>설교제목</summary>
			public string PreachTitle { get; set; }
			/// <summary>생일자 유무</summary>
			/// <value>생일자가 있으면 <c>true</c>, 없으면 <c>false</c></value>
			public bool IsBirthday { get; set; }
			/// <summary>생일자 이름 리스트</summary>
			/// <remarks><c>BirthPersonList</c>의 원소는 
			/// <c>BirthDateList</c>의 원소와 매칭된다
			/// </remarks>
			public List<string> BirthPersonList { get; set; }
			/// <summary>생일 날짜 리스트</summary>
			/// <remarks><c>BirthDateList</c>의 원소는 
			/// <c>BirthPersonList</c>의 원소와 매칭된다
			/// </remarks>
			public List<string> BirthDateList { get; set; }
			
			///	<summary>생성자</summary>
			/// <param name="fileName">주보 파일 이름</param>
			public Jubo(string fileName) 
			{
				FileName = fileName+".hwp";
				BVSStart = new BibleVerseSkeleton();
				BVSEnd = new BibleVerseSkeleton();
				BirthPersonList = new List<string>();
				BirthDateList = new List<string>();
			}

			/// <summary>
			/// 주보 파일에서 추출한 텍스트에서 말씀범위,대표기도자,설교제목,생일자에 대한 정보를 찾는다.
			/// </summary>
			/// <param name="source"></param>
			public void Parse(string source) {
				List<string> wordList = source.Trim().Split(new string[]{" ","\r\n"}, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
				int idx_worship_start = wordList.IndexOf("사도신경");
				int idx_worship_end = wordList.IndexOf("주기도문");
				int idx_ad=wordList.IndexOf("소식");		//not use yet
				int idx_birth=wordList.IndexOf("생일자");

				//대표기도
				int prayer=wordList.IndexOf("대표기도", idx_worship_start, idx_worship_end - idx_worship_start);
				PrayerName = wordList[prayer + 1] + " " + wordList[prayer + 2];

				//성경봉독
				int verse = prayer + 4;
				BVSStart.book = wordList[verse];
				BVSEnd.book = BVSStart.book;
				int ch_length = wordList[verse + 1].Length;

				string[] range = wordList[verse + 1].Split('-', '~', ':');
				if (range.Length == 4)
                {
					BVSStart.chapter = int.Parse(range[0]);
					BVSStart.passage = int.Parse(range[1]);
					BVSEnd.chapter = int.Parse(range[2]);
					BVSEnd.passage = int.Parse(range[3]);
                }

				else {
							if (char.IsLetter(wordList[verse + 1][ch_length - 1])) { wordList[verse + 1] = wordList[verse + 1].Substring(0, ch_length - 1); }
							BVSStart.chapter = int.Parse(wordList[verse + 1]);
							BVSEnd.chapter = BVSStart.chapter;
							int ps_length = wordList[verse + 2].Length;
							if (char.IsLetter(wordList[verse + 2][ps_length - 1])) { wordList[verse + 2] = wordList[verse + 2].Substring(0, ps_length - 1); }
							range = wordList[verse + 2].Split('-', '~', ':');
							BVSStart.passage = int.Parse(range[0]);
							BVSEnd.passage = int.Parse(range[1]);
				}

				//말씀선포
				int preach= wordList.IndexOf("말씀선포", idx_worship_start, idx_worship_end - idx_worship_start);
				int consecr = wordList.IndexOf("봉헌", preach, idx_worship_end - preach);
				for(int i = preach + 1; i < consecr - 2; i++)
				{
					PreachTitle += wordList[i] + " ";
				}
				PreachTitle.Trim();

				//생일자
				for(int iter = idx_birth + 1; ; iter+=2)
				{
					if (char.IsDigit(wordList[iter][0]))
					{
						BirthDateList.Add(wordList[iter]);
						//선생님일 경우
						if (wordList[iter + 2] == "T")
						{
							iter++;
							BirthPersonList.Add(wordList[iter] + " " + wordList[iter + 1]);
						}
						else
							BirthPersonList.Add(wordList[iter + 1]);
					}
					else
					{
						break;
					}
				}
				if (BirthPersonList.Count > 0)
				{
					IsBirthday = true;
				}
			}
			
			/// <summary>
			/// 주보 파일의 위치를 저장.
			/// </summary>
			/// <param name="directory"></param>
			public void SetPathInfo(string directory)
			{
				Directory = directory;
				FullPath = Directory + @"\" + FileName;
			}

			///	<summary>주보 파일을 파싱한 string을 반환.</summary>
			///	<remarks>파일위치,말씀범위,대표기도자,설교제목,생일자</remarks>
			///	<returns>string</returns>
			public string GetJuboInfo()
			{
				string info="";
				info += this.FullPath + "\n";
				info += BVSStart.ToString() + "~" + BVSEnd.ToString() + "\n";
				info += PrayerName + "\n";
				info += PreachTitle + "\n";
				info += "생일자:\n";
				int n = BirthPersonList.Count;
				for(int i = 0; i < n; i++)
				{
					if (i != 0) info += ", ";
					info += BirthPersonList[i] + " " + BirthDateList[i];
				}
				
				return info;
			}

			/**	<summary>생일 광고용 string 반환</summary>
			*	<returns>생일자 명단,날짜 데이터를 {이름}({월}.{일}{요일})의 형태로 ','로 구분한 string을 반환한다.<br/>
			*	<example>
			*	생일자:{홍길동,김서방},날짜:{11/9수,11/13주일} -> 홍길동(11/9수), 김서방(11/13주일)
			*	</example>
			*	</returns>
			*/
			public string GetBirthAdString()
			{
				string text = "";
				for(int i = 0; i < BirthPersonList.Count; i++)
				{
					if (i != 0) text += ", ";
					string date = BirthDateList[i];
					int len = date.Length;
					if (date[^2] == '일')
					{
						date = date.Remove(len-4,2);
					}
					else
					{
						date = date.Remove(len-3,1);
					}

					text += BirthPersonList[i] + "(" + date;
				}
				return text;
			}
		}

		/// <summary>
		/// 성경 구조 정보
		/// </summary>
		public class BibleVerseSkeleton
		{
			/// <summary>책 이름. ex) 창세기,출애굽기,레위기,...</summary>
			public string book { get; set; }
			/// <summary>장 혹은 편.</summary>
			public int chapter { get; set; }
			/// <summary>절.</summary>
			public int passage { get; set; }
			public BibleVerseSkeleton() { }
			/// <summary>Initialize class with specific book, chapter, passage.</summary>
			/// <param name="book">책 이름</param>
			/// <param name="chapter">장</param>
			/// <param name="passage">절</param>
			public BibleVerseSkeleton(string book, int chapter, int passage)
			{
				this.book = book;
				this.chapter = chapter;
				this.passage = passage;
			}
			public BibleVerseSkeleton(BibleVerseSkeleton bvs)
			{
				this.book = bvs.book;
				this.chapter = bvs.chapter;
				this.passage = bvs.passage;
			}

			/** <remarks>override basic ToString() method.</remarks>
			 * <returns>책 이름,장,절에 대한 string.<br/>
			 * <example><br/>ex) 창세기 1:1</example>
			 * </returns>
			 */
			public override string ToString()
			{
				return book + " " + chapter.ToString() + ":" + passage.ToString();
			}
		}

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
		/// 다가오는 일요일의 날짜를 구한다.
		/// </summary>
		/// <returns>해당 날짜를 나타내는 DateTime Object</returns>
		private static DateTime GetComingSundayDate()
        {
			int daysRemain = (7 - (int)DateTime.Now.DayOfWeek) % 7;
			DateTime dt = DateTime.Now.AddDays(daysRemain);
			return dt;
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

			dt = GetComingSundayDate();
			TxtOutputFileName.Text = dt.ToString("yyyy.MM.dd") + " 고등부 예배.pptx";

			settings = new Settings();
		}

		/// <summary>
		/// 메인화면의 로고 이미지 로딩.
		/// </summary>
		private void LoadLogoImage()
        {
			BitmapImage image = new BitmapImage(new Uri("pack://application:,,,/Resources/logo01.png"));
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
					MessageBox.Show("Opened");
					string txt = (string)hwp.GetTextFile("TEXT", "");
					if (!System.IO.Directory.Exists(AppDomain.CurrentDomain.BaseDirectory + OutputDirectory))
						System.IO.Directory.CreateDirectory(AppDomain.CurrentDomain.BaseDirectory + OutputDirectory);
					System.IO.File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + OutputDirectory + dt.ToString("yy-MM-dd") + ".txt", txt);

					jubo.Parse(txt);
					MessageBox.Show(jubo.GetJuboInfo());
				}
				else
				{
					MessageBox.Show("File Open Failed");
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
		private bool CheckErrorBeforeDoTask()
        {
			string errorMsg = Properties.Resources.TaskErrorCheckMessage;
			int errorCheckNum = 1;
			if (!Directory.Exists(TxtOutputFolder.Text))
			{
				errorMsg += "\n" + errorCheckNum++.ToString() + ". " + Properties.Resources.InvalidWorkingDirectory;
			}
			if (!File.Exists(settings.templateFileFullPath))
			{
				errorMsg += "\n" + errorCheckNum++.ToString() + ". " + Properties.Resources.InvalidPresentationTemplate;
			}
			if (SongList.Items.Count == 0)
			{
				errorMsg += "\n" + errorCheckNum++.ToString() + ". " + Properties.Resources.NoPraise;
			}
			if (!File.Exists(TxtPreachLocation.Text))
			{
				errorMsg += "\n" + errorCheckNum++.ToString() + ". " + Properties.Resources.InvalidPreachPPT;
			}
			if (!TxtOutputFileName.Text.EndsWith(".pptx") || TxtOutputFileName.Text.StartsWith(".pptx"))
			{
				errorMsg += "\n" + errorCheckNum++.ToString() + ". " + Properties.Resources.InvalidFileName;
			}
			if (errorCheckNum == 1) return false;
			else
			{
				MessageBox.Show(errorMsg);
				return true;
			}
		}
		/// <summary>
		/// 새 PPT를 작성한다. PowerPoint가 실행되고 모든 작업을 마친 파일을 연다. 저장 위치는 작업폴더와 같음.
		/// </summary>
		/// <remarks>인덱스 참조 편의를 위해 예배 진행 순서의 역순으로 수행.</remarks>
		private void DoTask(object sender, RoutedEventArgs e)
		{
			if (CheckErrorBeforeDoTask()) return;
			
			PowerPoint.Application pptApp = new PowerPoint.Application();
			PowerPoint.Presentations pptPres = pptApp.Presentations;
			PowerPoint.Presentation presentation = pptPres.Open(settings.templateFileFullPath);

			if (CbBirth.IsChecked == true)
			{
				//생일자 명단 입력
				presentation.Slides[settings.SettingsAdBirthList].Shapes[1].TextFrame.TextRange.Text = 
					TxtBirthList.Text.Replace(", ","\n");
			}
			else
			{
				//생일 영역 삭제
				presentation.Slides[settings.SettingsAdBirthEntry].Delete();
				presentation.Slides[settings.SettingsAdBirthEntry].Delete();
			}

			//설교
			PowerPoint.Presentation preachPPT = pptPres.Open(TxtPreachLocation.Text, WithWindow: MsoTriState.msoFalse);
			preachPPT.Slides.Range().Copy();

			presentation.Windows[1].Activate();
			presentation.Windows[1].View.GotoSlide(settings.SettingsPreachEntry);
			pptApp.CommandBars.ExecuteMso("PasteSourceFormatting");
			presentation.Slides[settings.SettingsPreachEntry].Shapes[2].TextFrame2.TextRange.Lines[2,1].Text = TxtTitle.Text;
			preachPPT.Close();

			//설교 전 영상
			if (System.IO.File.Exists(TxtVidLocation.Text))
				presentation.Slides[settings.SettingsVidBeforePreach].Shapes.AddMediaObject2(TxtVidLocation.Text);

			//말씀
			//3:제목 4:구절
			string verseString = "";
			verseString += jubo.BVSStart.book +" ";
			verseString += jubo.BVSStart.chapter.ToString();
			if (jubo.BVSStart.book == "시편")
				verseString += "편 ";
			else
				verseString += "장 ";
			verseString += jubo.BVSStart.passage.ToString() + "-" + jubo.BVSEnd.passage.ToString() + "절";

			presentation.Slides[settings.SettingsBibleEntry].Shapes[4].TextFrame.TextRange.Text = verseString;

			//6:범위 3:본문
			//List<string> verses=GetAllBibleVerse(jubo.BVSStart, passagesNum);
			List<Tuple<string,string>> verses = GetAllBibleVerseDB(jubo.BVSStart, jubo.BVSEnd);
			int passagesNum = verses.Count();
            presentation.Slides[settings.SettingsBibleEntry + 1].Shapes[6].TextFrame.TextRange.Text = verseString.Replace('-', '~');
            for (int i = 0; i < passagesNum-1; i++)
            {
                presentation.Slides[settings.SettingsBibleEntry + 1].Duplicate();
            }
				
            for (int i = 0; i < passagesNum; i++)
            {
				string[] va = verses[i].Item1.Split(':');
				if (int.Parse(va[0]) != jubo.BVSStart.chapter)
                {
					presentation.Slides[settings.SettingsBibleEntry + 1 + i].Shapes[3].TextFrame.TextRange.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone;
					presentation.Slides[settings.SettingsBibleEntry + 1 + i].Shapes[3].TextFrame.TextRange.Text = va[1] + ". " + verses[i].Item2;
				}
                else
                {
					presentation.Slides[settings.SettingsBibleEntry + 1 + i].Shapes[3].TextFrame.TextRange.Text = verses[i].Item2;
					presentation.Slides[settings.SettingsBibleEntry + 1 + i].Shapes[3].TextFrame.TextRange.ParagraphFormat.Bullet.StartValue = int.Parse(va[1]);
                }
            }
			

            //기도
            presentation.Slides[settings.SettingsPrayerNotice].Shapes[2].TextFrame.TextRange.Text= TxtPrayer.Text;
			
			//찬양
			for (int i = SongList.Items.Count - 1; i >= 0; i--)
			{
				PowerPoint.Presentation item = pptPres.Open((string)SongList.Items[i], WithWindow: MsoTriState.msoFalse);
				item.Slides.Range().Copy();

				presentation.Windows[1].Activate();
				presentation.Windows[1].View.GotoSlide(settings.SettingsPraiseEntry);
				pptApp.CommandBars.ExecuteMso("PasteSourceFormatting");
				//찬양 제목 슬라이드
				presentation.Slides[settings.SettingsPraiseSlidesInsertPos].Duplicate();
				presentation.Slides[settings.SettingsPraiseSlidesInsertPos].Shapes.Range().Delete();
				presentation.Slides[settings.SettingsPraiseEntry].Shapes.Range().Copy();
				presentation.Slides[settings.SettingsPraiseSlidesInsertPos].Shapes.Paste();
				presentation.Slides[settings.SettingsPraiseSlidesInsertPos].Shapes[1].TextFrame.TextRange.Text = 
					System.IO.Path.GetFileNameWithoutExtension(SongList.Items[i] as string);

				item.Close();
			}
			presentation.Slides[settings.SettingsPraiseEntry].Delete();
			
			//Save
			string fileName = @"\"+TxtOutputFileName.Text;
			string tempFilePath;
			if (!Directory.Exists(TxtOutputFolder.Text))
			{
				//tempFilePath=Documents
			}
			tempFilePath = TxtOutputFolder.Text + @"\new";
			string finalFilePath = TxtOutputFolder.Text + fileName;

            if (!Directory.Exists(finalFilePath))
            {
				File.Delete(finalFilePath);
            }
			presentation.Export(tempFilePath, "pptx");
			presentation.Close();

			//Open the created Presentation
			System.IO.File.Move(tempFilePath + ".pptx", finalFilePath);
			pptPres.Open(finalFilePath);
		}
		
		/// <summary>
		/// Settings Window를 표시한다.
		/// </summary>
        private void btnShowSettings_Click(object sender, RoutedEventArgs e)
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
    }
}
