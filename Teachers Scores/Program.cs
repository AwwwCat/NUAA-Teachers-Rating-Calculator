using ExcelDataReader;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Reflection.Emit;
using System.Text;


namespace TeacherApplication
{
	enum Datas { s1_1_1, s1_1_2, s1_1_3, s1_2, s1_3, s1_4, s2_1, s2_2, s3, as4_1, as4_2, as4_3, as4_4, as4_5, as4_6, as4_7, as4_8, as4_9, as4_10, as4_11, bs4_1, bs4_2, bs4_3, bs4_4, bs4_5, bs4_6, bs4_7, bs4_8, bs4_9, bs4_10, bs4_11, bs4_12, cs4_4, cs4_5_1, cs4_5_2, cs4_5_3, cs4_5_4, cs4_5_5, cs4_5_6, cs4_6, cs4_7, cs4_8, cs4_9, cs4_10, ds2_1, ds2_2, ds2_3, ds2_4, ds2_5, ds2_6, ds2_7, ds2_8, ds2_9, ds2_10, ds2_11, ds2_12, es2_1, es2_2_1, es2_2_2, es2_2_3, es2_2_4, es2_2_5, es2_2_6, es2_3, es2_4, es2_5, es2_6, es2_7, ds3, s5_1_1, s5_1_2, s5_1_3, s5_2, s5_3, as5_4, bs5_4, cs5_4, s5_5, s5_6, s6, max };
	enum Scores { s1, s2, s3, s4, s5, s6, final, max };
	struct Teachers
	{
		public string name;
		public int level;
		public string unit;
		public double[] datas;
		public double[] scores;
		public string[] comp;		// 是否完成达标

		public Teachers()
		{
			name = string.Empty;
			level = 0;
			unit = "未知";
			datas = new double[(int)Datas.max];
			scores = new double[(int)Scores.max];
			comp = new string[2];
		}
	};

	// Excel读取类
	class Excel			
	{
		DataSet result = new DataSet();

		// 打开Excel
		public void OpenExcel(string path)
		{
			using (FileStream fs = File.Open(path, FileMode.Open, FileAccess.Read))
			{
				IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fs);
				result = reader.AsDataSet();

				fs.Close();
			}
		}

		// 返回excel的数据
		public DataSet SendData()
		{
			return result;
		}
	}

	class Program
	{
		static int MaxTeacher = 100;		// 最大教师数，若大于此数请修改

		// 在excel中寻找教师名字并返回行数
		static int NameFind(DataSet data, Teachers[] tc, int j)			
		{
			DataTable table = data.Tables[0];
			DataRow row = table.Rows[j];
			for (int i = 0; i < tc.Length; i++)
			{
				if (tc[i].name == row[0].ToString())
					return i;
			}

			return -1;
		}

		// 导入教师姓名及岗位数据并按照教师数重新定义教师结构体数组
		static void IName(DataSet data, ref Teachers[] tc)
		{
			Console.Write("正在导入教师岗位...");
			DataTable table = data.Tables[0];
			MaxTeacher = table.Rows.Count;
			Console.Write("共" + MaxTeacher + "份数据...");
			tc = new Teachers[MaxTeacher];
			DataRow row;
			
			for (int i = 0;  i < MaxTeacher; i++) 
			{
				row = table.Rows[i];
				tc[i].name = row[0].ToString();
				tc[i].level = Convert.ToInt32(row[1]);
			}
			Console.WriteLine("完成");
		}

		// 导入教师数据，数据为数字
		static void IData(DataSet data, ref double[] dt, int j, int i)
		{
			DataTable table = data.Tables[0];
			DataRow row = table.Rows[j];
			dt[i] = Convert.ToDouble(row[1]);
		}

		// 重载，导入教师数据，数据为字符串
		static void IData(DataSet data, ref string ut, int j)
		{
			DataTable table = data.Tables[0];
			DataRow row;
			row = table.Rows[j];

			ut = Convert.ToString(row[1]);
		}

		// 根据要求计算分数
		static void ScoresCalculate1(int lv, double[] dt, ref double[] sc, ref string[] cp)
		{
			sc[(int)Scores.s6] += dt[(int)Datas.s6];

			switch (lv)
			{
				case 2:
				case 3:
				case 4:
					sc[(int)Scores.s3] += (dt[(int)Datas.s3] / 120 >= 1.5) ? 1.5 : dt[(int)Datas.s3] / 120;

					for (int i = (int)Datas.as4_1; i <= (int)Datas.as4_11; i++)
					{
						sc[(int)Scores.s4] += dt[i];
					}

					for (int i = (int)Datas.s5_1_1; i <= (int)Datas.s5_1_3; i++)
					{
						sc[(int)Scores.s5] += dt[i];
					}
					sc[(int)Scores.s5] /= 4;
					sc[(int)Scores.s5] += dt[(int)Datas.s5_2] / 2 + dt[(int)Datas.s5_3] / 2;
					sc[(int)Scores.s5] = (sc[(int)Scores.s5] >= 1.5) ? 1.5 : sc[(int)Scores.s5];
					sc[(int)Scores.s5] += dt[(int)Datas.as5_4] + dt[(int)Datas.s5_5] + dt[(int)Datas.s5_6];

					break;
				case 5:
				case 6:
				case 7:
					sc[(int)Scores.s3] += (dt[(int)Datas.s3] / 90 >= 1.5) ? 1.5 : dt[(int)Datas.s3] / 90;

					for (int i = (int)Datas.bs4_1; i <= (int)Datas.bs4_12; i++)
					{
						sc[(int)Scores.s4] += dt[i];
					}

					for (int i = (int)Datas.s5_1_1; i <= (int)Datas.s5_1_3; i++)
					{
						sc[(int)Scores.s5] += dt[i];
					}
					sc[(int)Scores.s5] /= 4;
					sc[(int)Scores.s5] += dt[(int)Datas.s5_2] + dt[(int)Datas.s5_3];
					sc[(int)Scores.s5] = (sc[(int)Scores.s5] >= 1.5) ? 1.5 : sc[(int)Scores.s5];
					sc[(int)Scores.s5] += dt[(int)Datas.bs5_4] + dt[(int)Datas.s5_5] + dt[(int)Datas.s5_6];

					break;
				case 8:
				case 9:
				case 10:
					for (int i = (int)Datas.s1_1_1; i <= (int)Datas.s1_4; i++)
					{
						sc[(int)Scores.s1] += dt[i];
					}
					sc[(int)Scores.s1] /= (32 * 3);

					
					sc[(int)Scores.s2] += dt[(int)Datas.s2_1];
					

					sc[(int)Scores.s3] += (dt[(int)Datas.s3] / 60 >= 1.5) ? 1.5 : dt[(int)Datas.s3] / 60;

					for (int i = (int)Datas.cs4_4; i <= (int)Datas.cs4_9; i++)
					{
						sc[(int)Scores.s4] += dt[i];
					}

					for (int i = (int)Datas.s5_1_1; i <= (int)Datas.s5_1_3; i++)
					{
						sc[(int)Scores.s5] += dt[i];
					}
					sc[(int)Scores.s5] /= 3;
					sc[(int)Scores.s5] += dt[(int)Datas.s5_2] + dt[(int)Datas.s5_3];
					sc[(int)Scores.s5] = (sc[(int)Scores.s5] >= 1.5) ? 1.5 : sc[(int)Scores.s5];
					sc[(int)Scores.s5] += dt[(int)Datas.cs5_4] + dt[(int)Datas.s5_5] + dt[(int)Datas.s5_6];

					for (int i = (int)Scores.s1; i <= (int)Scores.s6; i++)
					{
						sc[(int)Scores.final] += sc[i];
					}

					sc[(int)Scores.final] = Math.Round(sc[(int)Scores.final] /= 4, 2);
					break;
				default:
					break;
			}

			switch (lv)
			{
				case 2:
				case 3:
				case 4:
				case 5:
				case 6:
				case 7:
					for (int i = (int)Datas.s1_1_1; i <= (int)Datas.s1_4; i++)
					{
						sc[(int)Scores.s1] += dt[i];
					}
					sc[(int)Scores.s1] /= (64 * 3);

					for (int i = (int)Datas.s2_1; i <= (int)Datas.s2_2; i++)
					{
						sc[(int)Scores.s2] += dt[i];
					}

					for (int i = (int)Scores.s1; i <= (int)Scores.s6; i++)
					{
						sc[(int)Scores.final] += sc[i];
					}

					sc[(int)Scores.final] = Math.Round(sc[(int)Scores.final] /= 6, 2);
					break;
				default:
					break;
			}

			switch (lv)
			{
				case 2:
				case 5:
				case 8:
					if (sc[(int)Scores.final] > 1.2)
						cp[1] = "达标";
					else
						cp[1] = "未达标";
					break;
				case 3:
				case 6:
				case 9:
					if (sc[(int)Scores.final] > 1.1)
						cp[1] = "达标";
					else
						cp[1] = "未达标"; ;
					break;
				case 4:
				case 7:
				case 10:
					if (sc[(int)Scores.final] > 1)
						cp[1] = "达标";
					else
						cp[1] = "未达标"; ;
					break;
				default: 
					break;
			}

			if (lv <= 7)
			{
				if (sc[(int)Scores.s4] >= 2 || sc[(int)Scores.s5] >= 3 || (sc[(int)Scores.s4] >= 1 && sc[(int)Scores.s5] >= 1))
					cp[0] = "完成";
				else if (sc[(int)Scores.s4] < 1 && sc[(int)Scores.s5] >= 1)
					cp[0] = "S4未完成";
				else if (sc[(int)Scores.s4] >= 1 && sc[(int)Scores.s5] < 1)
					cp[0] = "S5未完成";
				else if (sc[(int)Scores.s4] < 1 && sc[(int)Scores.s5] < 1)
					cp[0] = "S4S5均未完成";
			}
			else
				cp[0] = "完成";
		}

		// 根据要求导出实验岗分数
		static void ScoresCalculate2(int lv, double[] dt, ref double[] sc, ref string[] cp)
		{
			switch (lv)
			{
				case 5:
				case 6:
				case 7:
					for (int i = (int)Datas.s1_1_1; i <= (int)Datas.s1_4; i++)
					{
						sc[(int)Scores.s1] += dt[i];
					}
					sc[(int)Scores.s1] /= (64 * 3);

					for (int i = (int)Datas.ds2_1; i <= (int)Datas.ds2_11; i++)
					{
						sc[(int)Scores.s2] += dt[i];
					}

					for (int i = (int)Datas.s5_1_1; i <= (int)Datas.s5_1_3; i++)
					{
						sc[(int)Scores.s5] += dt[i];
					}
					sc[(int)Scores.s5] /= 4;
					sc[(int)Scores.s5] += dt[(int)Datas.s5_2] + dt[(int)Datas.ds2_12] + dt[(int)Datas.bs5_4];
					sc[(int)Scores.s5] = (sc[(int)Scores.s5] >= 1.5) ? 1.5 : sc[(int)Scores.s5];
					sc[(int)Scores.s5] += dt[(int)Datas.s5_3] + dt[(int)Datas.s5_5] + dt[(int)Datas.s5_6];
					sc[(int)Scores.s2] += sc[(int)Scores.s5];

					break;
				case 8:
				case 9:
				case 10:
				case 11:
				case 12:
					for (int i = (int)Datas.s1_1_1; i <= (int)Datas.s1_4; i++)
					{
						sc[(int)Scores.s1] += dt[i];
					}
					sc[(int)Scores.s1] /= (32 * 3);

					for (int i = (int)Datas.es2_1; i <= (int)Datas.es2_7; i++)
					{
						sc[(int)Scores.s2] += dt[i];
					}

					for (int i = (int)Datas.s5_1_1; i <= (int)Datas.s5_1_3; i++)
					{
						sc[(int)Scores.s5] += dt[i];
					}
					sc[(int)Scores.s5] /= 3;
					sc[(int)Scores.s5] += dt[(int)Datas.s5_2] + dt[(int)Datas.s5_3] + dt[(int)Datas.cs5_4] + dt[(int)Datas.s5_5] + dt[(int)Datas.s5_6];
					// sc[(int)Scores.s5] = (sc[(int)Scores.s5] >= 1.5) ? 1.5 : sc[(int)Scores.s5];
					sc[(int)Scores.s2] += sc[(int)Scores.s5];

					break;
				default:
					break;
			}

			sc[(int)Scores.s3] += dt[(int)Datas.ds3];
			sc[(int)Scores.s4] += dt[(int)Datas.s6];
			sc[(int)Scores.s5] = 0;

			for (int i = (int)Scores.s1; i <= (int)Scores.s4; i++)
			{
				sc[(int)Scores.final] += sc[i];
			}

			sc[(int)Scores.final] = Math.Round(sc[(int)Scores.final] /= 4, 2);

			switch (lv)
			{
				case 2:
				case 5:
				case 8:
				case 11:
					if (sc[(int)Scores.final] > 1.2)
						cp[1] = "达标";
					else
						cp[1] = "未达标";
					break;
				case 3:
				case 6:
				case 9:
				case 12:
					if (sc[(int)Scores.final] > 1.1)
						cp[1] = "达标";
					else
						cp[1] = "未达标"; ;
					break;
				case 4:
				case 7:
				case 10:
					if (sc[(int)Scores.final] > 1)
						cp[1] = "达标";
					else
						cp[1] = "未达标"; ;
					break;
				default:
					break;
			}

			if (lv <= 7)
			{
				if (sc[(int)Scores.s2] >= 1)
					cp[0] = "完成";
				else 
					cp[0] = "S2未完成";
			} else
				cp[0] = "完成";
		}

		// 输出CSV文件
		static void ODataCSV1(StreamWriter sw, string nm, int lv, string ut, double sc, string[] cp)
		{
			sw.Write(lv.ToString() + ',' + nm + ',' + ut + ',' + sc + ',' + cp[1] + ',');
			if (cp[0] == "完成")
				sw.WriteLine(",,");
			else if (cp[0] == "S4未完成")
				sw.WriteLine(",未完成,");
			else if (cp[0] == "S5未完成")
				sw.WriteLine(",,未完成");
			else if (cp[0] == "S4S5均未完成")
				sw.WriteLine(",未完成,未完成");
			else if (cp[0] == "S2未完成")
				sw.WriteLine("未完成,,");
		}

		// 导出TXT文件
		static void ODataTXT1(StreamWriter sw, string nm, int lv, string ut, double[] dt, double[] sc, string[] cp)
		{
			sw = new StreamWriter("F:\\Desktop\\Teachers Scores\\TXT\\" + nm + ".txt", false, Encoding.GetEncoding("gb2312"));

			sw.WriteLine("============================================");
			sw.WriteLine("姓名: " + nm);
			sw.WriteLine("岗位等级: " + lv);
			sw.WriteLine("所在单位: " + ut);
			sw.WriteLine("分数: " + sc[(int)Scores.final]);
			sw.WriteLine("============================================");
			sw.WriteLine("必须完成项: " + cp[0]);
			sw.WriteLine("是否达标: " + cp[1]);
			sw.WriteLine("============================================");
			sw.WriteLine("单项得分: ");
			if (lv <= 7)
			{
				sw.WriteLine("S1 得分: " + Math.Round(sc[(int)Scores.s1], 2));
				sw.WriteLine("S2 得分: " + Math.Round(sc[(int)Scores.s2], 2));
				sw.WriteLine("S3 得分: " + Math.Round(sc[(int)Scores.s3], 2));
				sw.WriteLine("S4 得分: " + Math.Round(sc[(int)Scores.s4], 2));
				sw.WriteLine("S5 得分: " + Math.Round(sc[(int)Scores.s5], 2));
				sw.WriteLine("S6 得分: " + Math.Round(sc[(int)Scores.s6], 2));
			} else
			{
				sw.WriteLine("S1 得分: " + Math.Round(sc[(int)Scores.s1], 2));
				sw.WriteLine("S2&S3 得分: " + Math.Round(sc[(int)Scores.s2] + sc[(int)Scores.s3], 2));
				sw.WriteLine("S4&S5 得分: " + Math.Round(sc[(int)Scores.s4] + sc[(int)Scores.s5], 2));
				sw.WriteLine("S6 得分: " + Math.Round(sc[(int)Scores.s6], 2));
			}
			sw.WriteLine("============================================");
			sw.WriteLine("明细: ");
			switch (lv)
			{
				case 2:
				case 3:
				case 4:
					sw.WriteLine("1. 本科生课时（不含毕业设计）: " + dt[(int)Datas.s1_1_1]);
					sw.WriteLine("2. 本科生毕业设计课时: " + dt[(int)Datas.s1_1_2]);
					sw.WriteLine("3. 班主任课时: " + dt[(int)Datas.s1_1_3]);
					sw.WriteLine("4. 研究生课时: " + dt[(int)Datas.s1_2]);
					sw.WriteLine("5. 本科生导师课时: " + dt[(int)Datas.s1_3]);
					sw.WriteLine("6. 指导科创课时: " + dt[(int)Datas.s1_4]);
					sw.WriteLine("7. 国家级新增立项项目: " + dt[(int)Datas.s2_1]);
					sw.WriteLine("8. 省部级新增立项项目: " + dt[(int)Datas.s2_2]);
					sw.WriteLine("9. 三年科研到款: " + dt[(int)Datas.s3]);
					sw.WriteLine("10. 评教前10%次数（含研究生课程）: " + dt[(int)Datas.as4_1]);
					sw.WriteLine("11. 校级及以上精品课程（前1）: " + dt[(int)Datas.as4_2]);
					sw.WriteLine("12. 校级及以上教材（前1）: " + dt[(int)Datas.as4_3]);
					sw.WriteLine("13. 校级及以上教改项目立项（前1）: " + dt[(int)Datas.as4_4]);
					sw.WriteLine("14. 省级及以上教学比赛奖: " + dt[(int)Datas.as4_5]);
					sw.WriteLine("15. 校级及以上教学团队: " + dt[(int)Datas.as4_6]);
					sw.WriteLine("16. 省级及以上教学平台: " + dt[(int)Datas.as4_7]);
					sw.WriteLine("17. 省级及以上一流专业建设点: " + dt[(int)Datas.as4_8]);
					sw.WriteLine("18. 省学科一级学生优博或优硕: " + dt[(int)Datas.as4_9]);
					sw.WriteLine("19. 指导学生参加创新竞赛或学科竞赛: " + dt[(int)Datas.as4_10]);
					sw.WriteLine("20. 教学成果奖: " + dt[(int)Datas.as4_11]);
					sw.WriteLine("21. SCI期刊论文: " + dt[(int)Datas.s5_1_1]);
					sw.WriteLine("22. 专利成果转化（前3）: " + dt[(int)Datas.s5_2]);
					sw.WriteLine("23. 国防科技成果鉴定的成果: " + dt[(int)Datas.s5_3]);
					sw.WriteLine("24. 科研成果奖: " + dt[(int)Datas.as5_4]);
					sw.WriteLine("25. 新增主持国家重大项目: " + dt[(int)Datas.s5_5]);
					sw.WriteLine("26. 省部级以上科研平台（前1）: " + dt[(int)Datas.s5_6]);
					sw.WriteLine("27. 集体工作达到要求: " + dt[(int)Datas.s6]);
					break;
				case 5:
				case 6:
				case 7:
					sw.WriteLine("1. 本科生课时（不含毕业设计）: " + dt[(int)Datas.s1_1_1]);
					sw.WriteLine("2. 本科生毕业设计课时: " + dt[(int)Datas.s1_1_2]);
					sw.WriteLine("3. 班主任课时: " + dt[(int)Datas.s1_1_3]);
					sw.WriteLine("4. 研究生课时: " + dt[(int)Datas.s1_2]);
					sw.WriteLine("5. 本科生导师课时: " + dt[(int)Datas.s1_3]);
					sw.WriteLine("6. 指导科创课时: " + dt[(int)Datas.s1_4]);
					sw.WriteLine("7. 国家级新增立项项目: " + dt[(int)Datas.s2_1]);
					sw.WriteLine("8. 省部级新增立项项目: " + dt[(int)Datas.s2_2]);
					sw.WriteLine("9. 三年科研到款: " + dt[(int)Datas.s3]);
					sw.WriteLine("10. 评教前10%次数（含研究生课程）: " + dt[(int)Datas.bs4_1]);
					sw.WriteLine("11. 校级及以上精品课程（前2）: " + dt[(int)Datas.bs4_2]);
					sw.WriteLine("12. 校级及以上教材（前2）: " + dt[(int)Datas.bs4_3]);
					sw.WriteLine("13. 校级及以上教改项目立项（前1）: " + dt[(int)Datas.bs4_4]);
					sw.WriteLine("14. 校级教学比赛奖: " + dt[(int)Datas.bs4_5]);
					sw.WriteLine("15. 校级及以上教学团队（前2）: " + dt[(int)Datas.bs4_6]);
					sw.WriteLine("16. 省级及以上教学平台（前2）: " + dt[(int)Datas.bs4_7]);
					sw.WriteLine("17. 省级及以上一流专业建设点（前3）: " + dt[(int)Datas.bs4_8]);
					sw.WriteLine("18. 省学科一级学生优博或优硕: " + dt[(int)Datas.bs4_9]);
					sw.WriteLine("19. 省级本科毕设: " + dt[(int)Datas.bs4_10]);
					sw.WriteLine("20. 指导学生参加创新竞赛或学科竞赛: " + dt[(int)Datas.bs4_11]);
					sw.WriteLine("21. 教学成果奖: " + dt[(int)Datas.bs4_12]);
					sw.WriteLine("22. SCI期刊论文: " + dt[(int)Datas.s5_1_1]);
					sw.WriteLine("23. EI期刊论文: " + dt[(int)Datas.s5_1_2]);
					sw.WriteLine("24. 重要核心期刊论文 : " + dt[(int)Datas.s5_1_3]);
					sw.WriteLine("25. 专利成果转化（前3）: " + dt[(int)Datas.s5_2]);
					sw.WriteLine("26. 国防科技成果鉴定的成果: " + dt[(int)Datas.s5_3]);
					sw.WriteLine("27. 科研成果奖: " + dt[(int)Datas.bs5_4]);
					sw.WriteLine("28. 新增主持国家重大项目: " + dt[(int)Datas.s5_5]);
					sw.WriteLine("29. 省部级以上科研平台（前3）: " + dt[(int)Datas.s5_6]);
					sw.WriteLine("30. 集体工作达到要求: " + dt[(int)Datas.s6]);
					break;
				case 8:
				case 9:
				case 10:
					sw.WriteLine("1. 本科生课时（不含毕业设计）: " + dt[(int)Datas.s1_1_1]);
					sw.WriteLine("2. 本科生毕业设计课时: " + dt[(int)Datas.s1_1_2]);
					sw.WriteLine("3. 班主任课时: " + dt[(int)Datas.s1_1_3]);
					sw.WriteLine("4. 研究生课时: " + dt[(int)Datas.s1_2]);
					sw.WriteLine("5. 本科生导师课时: " + dt[(int)Datas.s1_3]);
					sw.WriteLine("6. 指导科创课时: " + dt[(int)Datas.s1_4]);
					sw.WriteLine("7. 国家级新增立项项目: " + dt[(int)Datas.s2_1]);
					sw.WriteLine("8. 三年科研到款: " + dt[(int)Datas.s3]);
					sw.WriteLine("9. SCI期刊论文: " + dt[(int)Datas.s5_1_1]);
					sw.WriteLine("10. EI期刊论文: " + dt[(int)Datas.s5_1_2]);
					sw.WriteLine("11. 重要核心期刊论文 : " + dt[(int)Datas.s5_1_3]);					
					sw.WriteLine("12. 专利成果转化（前3）: " + dt[(int)Datas.s5_2]);					
					sw.WriteLine("13. 国防科技成果鉴定的成果: " + dt[(int)Datas.s5_3]);
					sw.WriteLine("14. 评教前10%次数（含研究生课程）: " + dt[(int)Datas.cs4_4]);
					sw.WriteLine("15. 校级及以上精品课程（前3）: " + dt[(int)Datas.cs4_5_1]);
					sw.WriteLine("16. 校级及以上教材（前3）: " + dt[(int)Datas.cs4_5_2]);
					sw.WriteLine("17. 校级及以上教改项目立项（前1）: " + dt[(int)Datas.cs4_5_3]);					
					sw.WriteLine("18. 校级及以上教学团队（前3）: " + dt[(int)Datas.cs4_5_4]);
					sw.WriteLine("19. 省级及以上教学平台（前3）: " + dt[(int)Datas.cs4_5_5]);
					sw.WriteLine("20. 省级及以上一流专业建设点（前3）: " + dt[(int)Datas.cs4_5_6]);
					sw.WriteLine("21. 校级教学比赛奖: " + dt[(int)Datas.cs4_6]);
					sw.WriteLine("22. 省学科一级学生优博或优硕: " + dt[(int)Datas.cs4_7]);
					sw.WriteLine("23. 省级本科毕设: " + dt[(int)Datas.cs4_8]);
					sw.WriteLine("24. 指导学生参加创新竞赛或学科竞赛: " + dt[(int)Datas.cs4_9]);
					sw.WriteLine("25. 教学成果奖: " + dt[(int)Datas.cs4_10]);
					sw.WriteLine("26. 科研成果奖: " + dt[(int)Datas.cs5_4]);
					sw.WriteLine("27. 新增主持国家重大项目: " + dt[(int)Datas.s5_5]);
					sw.WriteLine("28. 省部级以上科研平台（前3）: " + dt[(int)Datas.s5_6]);
					sw.WriteLine("29. 集体工作达到要求: " + dt[(int)Datas.s6]);
					break;
				default:
					break;
			}
			sw.WriteLine("============================================");
			sw.WriteLine("计算规则: ");
			switch (lv)
			{
				case 2:
				case 3:
				case 4:
					sw.WriteLine("第1-6项之和除以192（64*3）为S1的得分");
					sw.WriteLine("第7、8项之和为S2得分");
					sw.WriteLine("第9项除以120为S3得分，其最大为1.5");
					sw.WriteLine("第10-20项之和为S4得分");
					sw.WriteLine("第21项的值除以4，加上第22、23项和除以2，加上24-26项和为S5得分，其中21-23项得分和最大为1.5");
					sw.WriteLine("第27项为S6得分");
					sw.WriteLine("6个单项总和除以6为最终分（四舍五入）");
					sw.WriteLine("S4、S5得分均大于1或S4得分大于2或S5得分大于3，必须完成项完成");
					break;
				case 5:
				case 6:
				case 7:
					sw.WriteLine("第1-6项之和除以192（64*3）为S1的得分");
					sw.WriteLine("第7、8项之和为S2得分");
					sw.WriteLine("第9项除以90为S3得分，其最大为1.5");
					sw.WriteLine("第10-21项之和为S4得分");
					sw.WriteLine("第22-24项之和除以4，加上第25-29项和为S5得分，其中22-26项得分和最大为1.5");
					sw.WriteLine("第30项为S6得分");
					sw.WriteLine("6个单项总和除以6为最终分（四舍五入）");
					sw.WriteLine("S4、S5得分均大于1或S4得分大于2或S5得分大于3，必须完成项完成");
					break;
				case 8:
				case 9:
				case 10:
					sw.WriteLine("第1-6项之和除以96（32*3）为S1的得分");
					sw.WriteLine("第8项除以60，加上第7项之和为S2&S3得分，其中第8项得分最大为1.5");
					sw.WriteLine("第9-11项之和除以3，加上第12-28项之和为S4&S5得分，其中9-13项得分和最大为1.5");
					sw.WriteLine("第29项为S6得分");
					sw.WriteLine("4个单项总和除以4为最终分（四舍五入）");
					break;
				default:
					break;
			}
			switch (lv)
			{
				case 2:
				case 5:
				case 8:
					sw.WriteLine("最终得分大于1.2即达标");
					break;
				case 3:
				case 6:
				case 9:
					sw.WriteLine("最终得分大于1.1即达标");
					break;
				case 4:
				case 7:
				case 10:
					sw.WriteLine("最终得分大于1.0即达标");
					break;
				default:
					break;
			}
			sw.WriteLine("============================================");
			sw.Close();
		}

		static void ODataTXT2(StreamWriter sw, string nm, int lv, string ut, double[] dt, double[] sc, string[] cp)
		{
			sw = new StreamWriter("F:\\Desktop\\Teachers Scores\\TXT\\" + nm + ".txt", false, Encoding.GetEncoding("gb2312"));

			sw.WriteLine("============================================");
			sw.WriteLine("姓名: " + nm);
			sw.WriteLine("岗位等级: " + lv);
			sw.WriteLine("所在单位: " + ut);
			sw.WriteLine("分数: " + sc[(int)Scores.final]);
			sw.WriteLine("============================================");
			sw.WriteLine("必须完成项: " + cp[0]);
			sw.WriteLine("是否达标: " + cp[1]);
			sw.WriteLine("============================================");
			sw.WriteLine("单项得分: ");
			sw.WriteLine("S1 得分: " + Math.Round(sc[(int)Scores.s1], 2));
			sw.WriteLine("S2 得分: " + Math.Round(sc[(int)Scores.s2], 2));
			sw.WriteLine("S3 得分: " + Math.Round(sc[(int)Scores.s3], 2));
			sw.WriteLine("S4 得分: " + Math.Round(sc[(int)Scores.s4], 2));
			sw.WriteLine("============================================");
			sw.WriteLine("明细: ");
			switch (lv)
			{
				case 5:
				case 6:
				case 7:
					sw.WriteLine("1. 本科生课时（不含毕业设计）: " + dt[(int)Datas.s1_1_1]);
					sw.WriteLine("2. 本科生毕业设计课时: " + dt[(int)Datas.s1_1_2]);
					sw.WriteLine("3. 班主任课时: " + dt[(int)Datas.s1_1_3]);
					sw.WriteLine("4. 研究生课时: " + dt[(int)Datas.s1_2]);
					sw.WriteLine("5. 本科生导师课时: " + dt[(int)Datas.s1_3]);
					sw.WriteLine("6. 指导科创课时: " + dt[(int)Datas.s1_4]);
					sw.WriteLine("7. 评教前10%次数（含研究生课程）: " + dt[(int)Datas.ds2_1]);
					sw.WriteLine("8. 校级及以上精品课程（前2）: " + dt[(int)Datas.ds2_2]);
					sw.WriteLine("9. 校级及以上教材（前2）: " + dt[(int)Datas.ds2_3]);
					sw.WriteLine("10. 校级及以上教改项目立项（前2）: " + dt[(int)Datas.ds2_4]);
					sw.WriteLine("11. 校级教学比赛奖: " + dt[(int)Datas.ds2_5]);
					sw.WriteLine("12. 校级及以上教学团队（前2）: " + dt[(int)Datas.ds2_6]);
					sw.WriteLine("13. 省级及以上教学平台（前2）: " + dt[(int)Datas.ds2_7]);
					sw.WriteLine("14. 省级及以上一流专业建设点（前3）: " + dt[(int)Datas.ds2_8]);
					sw.WriteLine("15. 省学科一级学生优博或优硕: " + dt[(int)Datas.ds2_9]);
					sw.WriteLine("16. 省级本科毕设: " + dt[(int)Datas.ds2_10]);
					sw.WriteLine("17. 指导学生参加创新竞赛或学科竞赛: " + dt[(int)Datas.ds2_11]);
					sw.WriteLine("18. 教学成果奖: " + dt[(int)Datas.ds2_12]);					
					sw.WriteLine("19. 科研成果奖: " + dt[(int)Datas.bs5_4]);
					sw.WriteLine("20. SCI期刊论文: " + dt[(int)Datas.s5_1_1]);
					sw.WriteLine("21. EI期刊论文: " + dt[(int)Datas.s5_1_2]);
					sw.WriteLine("22. 重要核心期刊论文 : " + dt[(int)Datas.s5_1_3]);
					sw.WriteLine("23. 专利成果转化（前3）: " + dt[(int)Datas.s5_2]);
					sw.WriteLine("24. 国防科技成果鉴定的成果: " + dt[(int)Datas.s5_3]);
					sw.WriteLine("25. 省部级以上科研平台（前3）: " + dt[(int)Datas.s5_6]);
					sw.WriteLine("26. 新增主持国家重大项目: " + dt[(int)Datas.s5_5]);
					sw.WriteLine("27. 实验室工作: " + dt[(int)Datas.ds3]);
					sw.WriteLine("28. 集体工作达到要求: " + dt[(int)Datas.s6]);
					break;
				case 8:
				case 9:
				case 10:
				case 11:
				case 12:
					sw.WriteLine("1. 本科生课时（不含毕业设计）: " + dt[(int)Datas.s1_1_1]);
					sw.WriteLine("2. 本科生毕业设计课时: " + dt[(int)Datas.s1_1_2]);
					sw.WriteLine("3. 班主任课时: " + dt[(int)Datas.s1_1_3]);
					sw.WriteLine("4. 研究生课时: " + dt[(int)Datas.s1_2]);
					sw.WriteLine("5. 本科生导师课时: " + dt[(int)Datas.s1_3]);
					sw.WriteLine("6. 指导科创课时: " + dt[(int)Datas.s1_4]);
					sw.WriteLine("7. 评教前10%次数（含研究生课程）: " + dt[(int)Datas.es2_1]);
					sw.WriteLine("8. 校级及以上精品课程（前3）: " + dt[(int)Datas.es2_2_1]);
					sw.WriteLine("9. 校级及以上教材（前3）: " + dt[(int)Datas.es2_2_2]);
					sw.WriteLine("10. 校级及以上教改项目立项（前1）: " + dt[(int)Datas.es2_2_3]);
					sw.WriteLine("11. 校级及以上教学团队（前3）: " + dt[(int)Datas.es2_2_4]);
					sw.WriteLine("12. 省级及以上教学平台（前3）: " + dt[(int)Datas.es2_2_5]);
					sw.WriteLine("13. 省级及以上一流专业建设点（前3）: " + dt[(int)Datas.es2_2_6]);
					sw.WriteLine("14. 校级教学比赛奖: " + dt[(int)Datas.es2_3]);
					sw.WriteLine("15. 省学科一级学生优博或优硕: " + dt[(int)Datas.es2_4]);
					sw.WriteLine("16. 省级本科毕设: " + dt[(int)Datas.es2_5]);
					sw.WriteLine("17. 指导学生参加创新竞赛或学科竞赛: " + dt[(int)Datas.es2_6]);
					sw.WriteLine("18. 教学成果奖: " + dt[(int)Datas.es2_7]);
					sw.WriteLine("19. 科研成果奖: " + dt[(int)Datas.cs5_4]);
					sw.WriteLine("20. SCI期刊论文: " + dt[(int)Datas.s5_1_1]);
					sw.WriteLine("21. EI期刊论文: " + dt[(int)Datas.s5_1_2]);
					sw.WriteLine("22. 重要核心期刊论文 : " + dt[(int)Datas.s5_1_3]);
					sw.WriteLine("23. 专利成果转化（前3）: " + dt[(int)Datas.s5_2]);
					sw.WriteLine("24. 国防科技成果鉴定的成果: " + dt[(int)Datas.s5_3]);
					sw.WriteLine("25. 省部级以上科研平台（前3）: " + dt[(int)Datas.s5_6]);					
					sw.WriteLine("26. 新增主持国家重大项目: " + dt[(int)Datas.s5_5]);
					sw.WriteLine("27. 实验室工作: " + dt[(int)Datas.ds3]);
					sw.WriteLine("28. 集体工作达到要求: " + dt[(int)Datas.s6]);
					break;
				default:
					break;
			}
			sw.WriteLine("============================================");
			sw.WriteLine("计算规则: ");
			switch (lv)
			{
				case 5:
				case 6:
				case 7:
					sw.WriteLine("第1-6项之和除以192（64*3）为S1的得分");
					sw.WriteLine("第20-22项之和除以4，加上第7-26其余项之和为S2得分，其中18-23项得分和最大为1.5");
					sw.WriteLine("第28项为S3得分");
					sw.WriteLine("第29项为S4得分");
					sw.WriteLine("4个单项总和除以5为最终分（四舍五入）");
					sw.WriteLine("S2得分均大于1，必须完成项完成");
					break;
				case 8:
				case 9:
				case 10:
				case 11:
				case 12:
					sw.WriteLine("第1-6项之和除以96（32*3）为S1的得分");
					sw.WriteLine("第20-22项之和除以3，加上第7-26其余项之和为S2得分");
					sw.WriteLine("第28项为S3得分");
					sw.WriteLine("第29项为S4得分");
					sw.WriteLine("4个单项总和除以5为最终分（四舍五入）");
					break;
				default:
					break;
			}
			switch (lv)
			{
				case 2:
				case 5:
				case 8:
				case 11:
					sw.WriteLine("最终得分大于1.2即达标");
					break;
				case 3:
				case 6:
				case 9:
				case 12:
					sw.WriteLine("最终得分大于1.1即达标");
					break;
				case 4:
				case 7:
				case 10:
					sw.WriteLine("最终得分大于1.0即达标");
					break;
				default:
					break;
			}
			sw.WriteLine("============================================");
			sw.Close();
		}

		static void Main(string[] args)
		{
			// 修改编码防止csv转xlsx出现乱码
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
			Encoding encoding1 = Encoding.GetEncoding(1252);
			Encoding encoding2 = Encoding.GetEncoding("GB2312");

			Teachers[] teachers = new Teachers[MaxTeacher];

			Console.WriteLine("============================================");

			{  
				Excel excel = new Excel();
				excel.OpenExcel("F:\\Desktop\\Teachers Scores\\Data\\" + "S0" + ".xlsx");		// 根据情况修改路径
				IName(excel.SendData(), ref teachers);
			}

			Console.Write("正在创建结构体数组...");
			for (int j = 0 ; j < teachers.Length; j++)
			{
				teachers[j].unit = "未知";
				teachers[j].datas = new double[(int)Datas.max];
				teachers[j].scores = new double[(int)Scores.max];
				teachers[j].comp = new string[2];
			}
			Console.WriteLine("完成");

			Console.Write("正在导入教师单位...");
			foreach (var tc in teachers)
			{
				Excel excel = new Excel();
				excel.OpenExcel("F:\\Desktop\\Teachers Scores\\Data\\S0-1.xlsx");               // 根据情况修改路径
				DataTable table = excel.SendData().Tables[0];
				for (int j = 0; j < table.Rows.Count; j++)
				{
					int k = NameFind(excel.SendData(), teachers, j);
					if (k >= 0)
						IData(excel.SendData(), ref teachers[k].unit, j);
					else
						continue;
				}
			}
			Console.WriteLine("完成");


			Console.WriteLine("============================================");

			string[] files = { "S1-1-1", "S1-1-2", "S1-1-3", "S1-2", "S1-3", "S1-4", "S2-1", "S2-2", "S3", "aS4-1", "aS4-2", "aS4-3", "aS4-4", "aS4-5", "aS4-6", "aS4-7", "aS4-8", "aS4-9", "aS4-10", "aS4-11", "bS4-1", "bS4-2", "bS4-3", "bS4-4", "bS4-5", "bS4-6", "bS4-7", "bS4-8", "bS4-9", "bS4-10", "bS4-11", "bS4-12", "cS4-4", "cS4-5-1", "cS4-5-2", "cS4-5-3", "cS4-5-4", "cS4-5-5", "cS4-5-6", "cS4-6", "cS4-7", "cS4-8", "cS4-9", "cS4-10", "dS2-1", "dS2-2", "dS2-3", "dS2-4", "dS2-5", "dS2-6", "dS2-7", "dS2-8", "dS2-9", "dS2-10", "dS2-11", "dS2-12", "eS2-1", "eS2-2-1", "eS2-2-2", "eS2-2-3", "eS2-2-4", "eS2-2-5", "eS2-2-6", "eS2-3", "eS2-4", "eS2-5", "eS2-6", "eS2-7", "dS3", "S5-1-1", "S5-1-2", "S5-1-3", "S5-2", "S5-3", "aS5-4", "bS5-4", "cS5-4", "S5-5", "S5-6", "S6" };
			int i = 0;
			foreach (var s in files)
			{
				Excel excel = new Excel();
				Console.Write("正在导入" + s + "...");
				excel.OpenExcel("F:\\Desktop\\Teachers Scores\\Data\\" + s + ".xlsx");          // 根据情况修改路径
				DataTable table = excel.SendData().Tables[0];
				Console.Write("共" + table.Rows.Count.ToString() + "份数据...");
				for (int j = 0; j < table.Rows.Count; j++)
				{
					int k = NameFind(excel.SendData(), teachers, j);
					if (k >= 0)
						IData(excel.SendData(), ref teachers[k].datas, j, i);
					else
						continue;
				}
				i++;
				Console.WriteLine("完成");
			}

			Console.WriteLine("============================================");

			Console.Write("正在计算...");
			for (int j = 0;j < teachers.Length; j++)
            {
				if (teachers[j].name == "xxx")			// 实验岗教师单独挑出计算
					ScoresCalculate2(teachers[j].level, teachers[j].datas, ref teachers[j].scores, ref teachers[j].comp);
				else
					ScoresCalculate1(teachers[j].level, teachers[j].datas, ref teachers[j].scores, ref teachers[j].comp);
			}
			Console.WriteLine("完成");

			Console.WriteLine("============================================");
			Console.WriteLine("输入C显示分数，按其他键退出");


			StreamWriter sw = new StreamWriter("F:\\Desktop\\Teachers Scores\\CSV\\Page1.csv", false, Encoding.GetEncoding("gb2312"));          // 根据情况修改路径
			sw.WriteLine("岗位等级,姓名,所在单位,最终分,是否达标,S2必须完成项,S4必须完成项,S5必须完成项");
			foreach (var tc in teachers) 
			{
				ODataCSV1(sw, tc.name, tc.level, tc.unit, tc.scores[(int)Scores.final], tc.comp);
			}
			sw.Close();

			foreach (var tc in teachers)
			{
				if (tc.name == "xxx")                   // 实验岗教师单独挑出输出
					ODataTXT2(sw, tc.name, tc.level, tc.unit, tc.datas, tc.scores, tc.comp);
				else
					ODataTXT1(sw, tc.name, tc.level, tc.unit, tc.datas, tc.scores, tc.comp);
			}

			if (Console.ReadKey(true).Key == ConsoleKey.C)
			{
				Console.WriteLine("============================================");
				foreach (var tc in teachers)
				{
					Console.Write(tc.name + '\t' + tc.level.ToString() + '\t');

					for (int j = (int)Scores.s1; j < (int)Scores.final; j++)
						Console.Write(Math.Round(tc.scores[j], 2).ToString() + '\t');
					Console.WriteLine(Math.Round(tc.scores[(int)Scores.final], 2).ToString() + '\t' + tc.unit);
				}
				Console.ReadKey(true);
			} 
			Environment.Exit(0);
		}
	}
}
