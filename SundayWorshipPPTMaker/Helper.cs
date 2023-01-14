using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Threading.Tasks;

namespace SundayWorshipPPTMaker
{
	///	<summary>
	///	템플릿 슬라이드 작업 상수
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

		public const float CoverImageHeight = 9.4f;
		public const float CoverImageWidth = 27.6f;
		public const float CoverImageTop = 2.52f;
		public const float CoverImageLeft = 3.13f;
		public const float CoverCommentHeight = 4.02f;
		public const float CoverCommentWidth = 16.86f;
		public const float CoverCommentTop = 12.53f;
		public const float CoverCommentLeft = 8.5f;
	}

	public class Utils
	{
		public static float CMToPoint(float cm)
		{
			return (cm / 2.54f) * 72;
		}
		public static float PointToCM(float pt)
		{
			return (pt / 72) * 2.54f;
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
			FileName = fileName + ".hwp";
			BVSStart = new BibleVerseSkeleton();
			BVSEnd = new BibleVerseSkeleton();
			BirthPersonList = new List<string>();
			BirthDateList = new List<string>();
		}

		public void SetPrayerName(ref List<string> wordList, ref int cur_idx, ref int end_idx)
		{
			cur_idx = wordList.IndexOf("대표기도", cur_idx, end_idx - cur_idx);
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
			bibleRangeText.Trim();
			string patternSinglePassage = @"[ㄱ-ㅎ|ㅏ-ㅣ|가-힣]+\s*\d+[장편:]\s*\d+[절]?";
			string patternMultiPassages = patternSinglePassage + @"\s*[-~]\s*\d+[절]?";
			string patternMultiChapters = patternSinglePassage + @"\s*[-~]\s*\d+[장편:]\s*\d+[절]?";

			Regex rxSingleNumber = new Regex(@"\d+");
			BVSStart.book = bibleRangeText.Split()[0];
			BVSEnd.book = BVSStart.book;
			MatchCollection matches = rxSingleNumber.Matches(bibleRangeText);
			int idx = Regex.IsMatch(bibleRangeText, @"요한[1-3]서") ? 1 : 0;
			if (Regex.IsMatch(bibleRangeText, patternMultiChapters))
			{
				BVSStart.chapter = int.Parse(matches[idx++].Value);
				BVSStart.passage = int.Parse(matches[idx++].Value);
				BVSEnd.chapter = int.Parse(matches[idx++].Value);
				BVSEnd.passage = int.Parse(matches[idx++].Value);
			}
			else if (Regex.IsMatch(bibleRangeText, patternMultiPassages))
			{
				BVSStart.chapter = int.Parse(matches[idx++].Value);
				BVSEnd.chapter = BVSStart.chapter;
				BVSStart.passage = int.Parse(matches[idx++].Value);
				BVSEnd.passage = int.Parse(matches[idx++].Value);
			}
			else if (Regex.IsMatch(bibleRangeText, patternSinglePassage))
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
			int idx_birth = wordList.IndexOf("생일자");
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
		/// <summary>
		/// 주보 파일에서 추출한 텍스트에서 말씀범위,대표기도자,설교제목,생일자에 대한 정보를 찾는다.
		/// </summary>
		/// <param name="source"></param>
		public void Parse(string source)
		{
			List<string> wordList = source.Trim().Split(new string[] { " ", "\r\n" }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
			int cur_idx = wordList.IndexOf("사도신경");
			int idx_worship_end = wordList.IndexOf("주기도문");
			int idx_ad = wordList.IndexOf("소식");      //not use yet

			SetPrayerName(ref wordList, ref cur_idx, ref idx_worship_end);
			SetBibleVerseRange(ref wordList, ref cur_idx, ref idx_worship_end);
			SetPreachTitle(ref wordList, ref cur_idx, ref idx_worship_end);
			CheckBirthday(ref wordList);
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
