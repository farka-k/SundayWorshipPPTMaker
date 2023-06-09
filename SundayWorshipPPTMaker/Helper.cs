using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Threading.Tasks;
using System.Text.Json.Serialization;
using System.Windows.Media;
using System.Globalization;
using System.Windows.Markup;
using System.Windows.Data;
using Tesseract;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using HtmlAgilityPack;

namespace SundayWorshipPPTMaker
{
	///	<summary>
	///	Template Constants
	///	</summary>
	public static class Constants
	{
		public static string BaseDirectory = AppDomain.CurrentDomain.BaseDirectory;
		public static string ResourceDirectory = AppDomain.CurrentDomain.BaseDirectory + @"Resources\";
		public const float SlideSize16x9Width = 33.867f;
		public const float SlideSizeHeight = 19.05f;
		public const float SlideSize4x3Width = 25.4f;
		public const float CoverImageWidth = 27.6f;
		public const float CoverImageHeight = 9.4f;
		public const float CoverImageTop = 4.0f;
		public const float CoverImageLeft = 3.13f;
		public const float CoverLogoWidth = 3.84f;
		public const float CoverLogoHeight = 3.37f;
		public const float CoverLogoTop = 2.71f;
		public const float CoverLogoLeft = 15.01f;
		public const float CoverCommentWidth = 16.86f;
		public const float CoverCommentHeight = 4.02f;
		public const float CoverCommentTop = 12.53f;
		public const float CoverCommentLeft = 8.5f;
		public const float PaneHeight = 14.32f;
		public const float HorizontalMargin = 0.26f;
		public const float VerticalMargin = 0.13f;

		public const string FontNanumSquareBold = "나눔스퀘어 Bold";
		public const string FontNanumSquareExBold = "나눔스퀘어라운드 ExtraBold";
		public const string FontMalgunGothic = "맑은 고딕";
		public const string FontKopubDotumBold = "Kopub돋움체 Bold";
		public const string FontSequenceTitle = "KT&G 상상제목 B";
		public const string FontSongTitle = "HY궁서B";
		public const float RoundedRectangleRadius = 0.07707f;
	}
	public struct slideSize
	{
		public float Width;
		public float Height;
		public slideSize(float width, float height)
		{
			Width = width;
			Height = height;
		}
	}

	public enum SlideContentsType { Cover, Main }
	public enum SlideSizeType { WideScreen, Normal }
	public enum OCREngineType { Clova, Tesseract }
	public enum TextEmphasis
	{
		None = 0b_0000_0000,
		Bold = 0b_0000_0001,
		Italic = 0b_0000_0010,
		UnderLine = 0b_0000_0100,
		Shadow = 0b_0000_1000,
		StrikeThrough = 0b_0001_0000
	}
	public class ShadowOptions
	{
		public ShadowOptions(MsoShadowStyle style = MsoShadowStyle.msoShadowStyleOuterShadow,
			int color = 0, float transparency = 0.57f, float size = 100, float blur = 3, double angle = 45, float distance = 3)
		{
			Style = style;
			Color = color;
			Transparency = transparency;
			Size = size;
			Blur = blur;
			Angle = angle;
			Distance = distance;
		}
		public MsoShadowStyle Style { get; set; }
		public int Color { get; set; }
		public float Transparency { get; set; }
		public float Size { get; set; }
		public float Blur { get; set; }
		public double Angle { get; set; }
		public float Distance { get; set; }
	}

	public static class Utils
	{
		public static float CMToPoint(float cm)
		{
			return (cm / 2.54f) * 72;
		}
		public static float PointToCM(float pt)
		{
			return (pt / 72) * 2.54f;
		}
		public static double CMToPixel(double cm)
		{
			return cm * (96 / 2.54);
		}
		public static double PixelToCM(double px)
		{
			return px / (96 / 2.54);
		}
		/// <summary>
		/// 다가오는 일요일의 날짜를 구한다.
		/// </summary>
		/// <returns>해당 날짜를 나타내는 DateTime Object</returns>
		public static DateTime GetComingSundayDate()
		{
			int daysRemain = (7 - (int)DateTime.Now.DayOfWeek) % 7;
			DateTime dt = DateTime.Now.AddDays(daysRemain);
			return dt;
		}

		public static bool IsLastSundayOfMonth()
		{
			DateTime dt = GetComingSundayDate();
			var next = dt.AddDays(7);
			if (next.Month != dt.Month) return true;
			else return false;
		}
	}

	public class FirstDegreeFunctionConverter : IValueConverter
	{
		public double A { get; set; }
		public double B { get; set; }

		#region IValueConverter Members

		public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
		{
			double a = GetDoubleValue(parameter, A);

			double x = GetDoubleValue(value, 0.0);

			return (a * x) + B;
		}

		public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
		{
			double a = GetDoubleValue(parameter, A);

			double y = GetDoubleValue(value, 0.0);

			return (y - B) / a;
		}

		#endregion


		private double GetDoubleValue(object parameter, double defaultValue)
		{
			double a;
			if (parameter != null)
				try
				{
					a = System.Convert.ToDouble(parameter);
				}
				catch
				{
					a = defaultValue;
				}
			else
				a = defaultValue;
			return a;
		}
	}

	public class ClovaOCRRequstFormat
	{
		[JsonInclude]
		public string version { get; set; }
		[JsonInclude]
		public string requestId { get; set; }
		[JsonInclude]
		public long timestamp { get; set; }
		public string lang { get; set; }
		[JsonInclude]
		public List<RequestImageJsonModel> images { get; set; }
	}

	public class ClovaOCRResponseFormat
	{
		//public string version { get; set; }
		//public string requestId { get; set; }
		//public long timestamp { get; set; }
		public List<ResponseImageJsonModel> images { get; set; }
	}

	public class RequestImageJsonModel
	{
		[JsonInclude]
		public string format { get; set; }
		[JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
		public string url { get; set; }
		[JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
		public string data { get; set; }
		[JsonInclude]
		public string name { get; set; }
		public List<int> templateIds { get; set; }
	}

	public class ResponseImageJsonModel
	{
		public string uid { get; set; }
		public string name { get; set; }
		public string inferResult { get; set; }
		public string message { get; set; }
		public MatchedTemplateModel matchedTemplate { get; set; }
		public ValidationResultModel validationResult { get; set; }
		public List<ImageFieldModel> fields { get; set; }
		public TitleModel title { get; set; }
	}

	public class MatchedTemplateModel
	{
		public int id { get; set; }
		public string name { get; set; }
	}

	public class ValidationResultModel
	{
		public string result { get; set; }
		public string message { get; set; }
	}

	public class ImageFieldModel
	{
		public string name { get; set; }
		public string valueType { get; set; }
		public string inferText { get; set; }
		public float inferConfidence { get; set; }
		public BoundingModel bounding { get; set; }
	}

	public class TitleModel
	{
		public string name { get; set; }
		public BoundingModel bounding { get; set; }
		public string inferText { get; set; }
		public float inferConfidence { get; set; }
	}

	public class BoundingModel
	{
		public float top { get; set; }
		public float left { get; set; }
		public float width { get; set; }
		public float height { get; set; }
	}

	public class SongItem
	{
		public string Title { get; set; }
		public string Path { get; set; }
		public bool PPTEnable { get; set; }
		public List<string>? Verse { get; set; }
		public LyricSlideOptions? Options { get; set; }
	}

	public class LyricSlideOptions
    {
		public int BackgroundColor { get; set; }
		public int FontColor { get; set; }
		public string FontName { get; set; }
		public float FontSize { get; set; }
		public TextEmphasis Emphasis { get; set; }
		public MsoVerticalAnchor VerticalAlignment{ get; set; }
		public float Offset { get; set; }
    }
	
	public class GoogleCustomSearchAPIResponseModel
    {
		public List<SearchResultItemModel> items { get; set; }
    }
	public class SearchResultItemModel
    {
		public string title { get; set; }
		public string link { get; set; }
    }

	public static class LyricsCollector
    {
		public static string GetLyrics(string url)
        {
			string fullLyrics = String.Empty;
			HtmlWeb web = new HtmlWeb();
			HtmlDocument htmlDoc = web.Load(url);

			var contentBody = htmlDoc.DocumentNode.SelectSingleNode("//div[@id='hyrendContentBody']");
			var lyricsSection = contentBody.SelectSingleNode("article/section[@class='sectionPadding contents lyrics']");
			var lyricsContainer = lyricsSection.SelectSingleNode("div/div[@class='lyricsContainer']");
			fullLyrics = lyricsContainer.SelectSingleNode("p/xmp").InnerText;

			return fullLyrics;
        }
    }

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
		/// <summary>찬양인도자</summary>
		public string WorshipLeader { get; set; }
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
		
		public string AdStrings { get; set; }

		///	<summary>생성자</summary>
		/// <param name="fileName">주보 파일 이름</param>
		public Jubo(string fileName)
		{
			FileName = fileName + ".hwp";
			BVSStart = new BibleVerseSkeleton();
			BVSEnd = new BibleVerseSkeleton();
			BirthPersonList = new List<string>();
			BirthDateList = new List<string>();
		}

		public void SetPrayerName(ref List<string> wordList, ref int cur_idx)
		{
			//cur_idx = wordList.IndexOf("대표기도", cur_idx, end_idx - cur_idx);
			PrayerName = wordList[cur_idx + 1] + " " + wordList[cur_idx + 2];
		}
		public void SetBibleVerseRange(ref List<string> wordList, ref int cur_idx, ref int end_idx)
		{
			int idx_read_bible = cur_idx + 3;
			cur_idx = wordList.IndexOf("말씀선포", idx_read_bible, end_idx - idx_read_bible);
			int textRange = cur_idx - 3 - idx_read_bible;

			string bibleRangeText = "";
			for (int i = idx_read_bible + 1; i < cur_idx - 2; i++)
			{
				bibleRangeText += wordList[i] + " ";
			}
			bibleRangeText=bibleRangeText.Trim();
			SetBibleVerseRange(bibleRangeText);
		}

		public void SetBibleVerseRange(string text)
        {
			string patternSinglePassage = @"[ㄱ-ㅎ|ㅏ-ㅣ|가-힣]+\s*\d+[장편:]\s*\d+[절]?";
			string patternMultiPassages = patternSinglePassage + @"\s*[-~]\s*\d+[절]?";
			string patternMultiChapters = patternSinglePassage + @"\s*[-~]\s*\d+[장편:]\s*\d+[절]?";

			Regex rxSingleNumber = new Regex(@"\d+");
			BVSStart.book = text.Split()[0];
			BVSEnd.book = BVSStart.book;
			MatchCollection matches = rxSingleNumber.Matches(text);
			int idx = Regex.IsMatch(text, @"요한[1-3]서") ? 1 : 0;
			if (Regex.IsMatch(text, patternMultiChapters))
			{
				BVSStart.chapter = int.Parse(matches[idx++].Value);
				BVSStart.passage = int.Parse(matches[idx++].Value);
				BVSEnd.chapter = int.Parse(matches[idx++].Value);
				BVSEnd.passage = int.Parse(matches[idx++].Value);
			}
			else if (Regex.IsMatch(text, patternMultiPassages))
			{
				BVSStart.chapter = int.Parse(matches[idx++].Value);
				BVSEnd.chapter = BVSStart.chapter;
				BVSStart.passage = int.Parse(matches[idx++].Value);
				BVSEnd.passage = int.Parse(matches[idx++].Value);
			}
			else if (Regex.IsMatch(text, patternSinglePassage))
			{
				BVSStart.chapter = int.Parse(matches[idx++].Value);
				BVSStart.passage = int.Parse(matches[idx++].Value);
				BVSEnd.chapter = BVSStart.chapter;
				BVSEnd.passage = BVSStart.passage;
			}
			else
			{
				MessageBox.Show("Parsing Fail");
			}
		}
		public void SetPreachTitle(ref List<string> wordList, ref int cur_idx, ref int end_idx)
		{
			PreachTitle = "";
			int preach = cur_idx;
			cur_idx = wordList.IndexOf("봉헌", preach, end_idx - preach);
			for (int i = preach + 1; i < cur_idx - 2; i++)
			{
				PreachTitle += wordList[i] + " ";
			}
			PreachTitle.Trim();
		}
		public void CheckBirthday(ref List<string> wordList)
		{
			BirthPersonList.Clear();
			BirthDateList.Clear();
			int idx_birth = wordList.FindIndex(0, s => s.Contains("생일자"));
			for (int iter = idx_birth + 1; ; iter += 2)
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

		private void SetWorshipLeader(ref List<string> wordList, ref int cur_idx, ref int idx_prayer)
        {
			WorshipLeader = wordList[idx_prayer - 2] + " " + wordList[idx_prayer - 1];
			cur_idx = idx_prayer;
        }

		private void SetAdStrings(ref List<string> wordList, ref int idx_ad,ref int idx_ad_end)
        {
			int start = idx_ad + 1;			
			var adString = String.Empty;
			int num;
			for(int i = start+1; ; i++)
            {
                if (i == idx_ad_end)
                {
					AdStrings += adString.Trim();
					break;
                }
				if(int.TryParse(wordList[i],out num))
                {
					AdStrings += adString.Trim() + "\n";
					adString = String.Empty;
                } 
				else adString += wordList[i] + " ";
            }			
        }

		/// <summary>
		/// 주보 파일에서 추출한 텍스트에서 말씀범위,대표기도자,설교제목,생일자에 대한 정보를 찾는다.
		/// </summary>
		/// <param name="source"></param>
		public void Parse(string source)
		{
			List<string> wordList = source.Trim().Split(new string[] { " ", "\r\n" }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
			int cur_idx = wordList.IndexOf("사도신경");
			int idx_worship_end = wordList.IndexOf("주기도문");
			int idx_prayer = wordList.IndexOf("대표기도", cur_idx, idx_worship_end - cur_idx);
			int idx_ad = wordList.IndexOf("소식");      //not use yet
			int idx_ad_end = wordList.IndexOf("지난주");

			SetWorshipLeader(ref wordList, ref cur_idx, ref idx_prayer);
			SetPrayerName(ref wordList, ref idx_prayer);
			SetBibleVerseRange(ref wordList, ref cur_idx, ref idx_worship_end);
			SetPreachTitle(ref wordList, ref cur_idx, ref idx_worship_end);
			CheckBirthday(ref wordList);
			SetAdStrings(ref wordList, ref idx_ad, ref idx_ad_end);
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
			string info = "";
			info += this.FullPath + "\n";
			info += BVSStart.ToString() + "~" + BVSEnd.ToString() + "\n";
			info += PrayerName + "\n";
			info += PreachTitle + "\n";
			info += "생일자:\n";
			int n = BirthPersonList.Count;
			for (int i = 0; i < n; i++)
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
			for (int i = 0; i < BirthPersonList.Count; i++)
			{
				if (i != 0) text += ", ";
				string date = BirthDateList[i];
				int len = date.Length;
				if (date[^2] == '일')
				{
					date = date.Remove(len - 4, 2);
				}
				else
				{
					date = date.Remove(len - 3, 1);
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
}