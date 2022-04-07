using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;

namespace PPT에서_가사_뽑아내는_매크로
{
    class Contribute {
        public String title;
        public String person;
        public String description;
    }

    class Program
    {
        static void PrintContribute()
        {
            List<Contribute> contributes = new List<Contribute>(10);
            contributes.Add(new Contribute()
            {
                title = "Programmer",
                person = "백성수",
                description = "프로그램 설계 및 제작"
            });
            contributes.Add(new Contribute()
            {
                title = "Icon Designer",
                person = "백지원",
                description = "아이콘 제작"
            });

            foreach (Contribute c in contributes)
                Console.Write("{0} : {1}\n{2}\n\n",c.title,c.person,c.description);
        }

        /// <summary>
        /// 잘못된 개행문자를 윈도우 표준에 맞게 바꿔줍니다.
        /// </summary>
        /// <param name="original">
        /// 원본 문자열
        /// </param>
        /// <returns>
        /// 고쳐진 문자열
        /// </returns>
        static string makeCorrectNewline(string original)
        {
            StringBuilder str = new StringBuilder(original);
            for (int i = 0, j = 0; i < original.Length; i++, j++)
            {
                if (original[i] == '\r')
                {
                    if ((i + 1 == original.Length) || (original[i + 1] != '\n'))
                    {
                        str = str.Insert(j + 1, '\n');
                        j++;
                    }
                    else
                    {
                        i++;
                        j++;
                    }
                }
                else if (original[i] == '\n')
                {
                    str = str.Insert(j, '\r');
                    j++;
                }
                else if (original[i] == '\v')
                {
                    str = str.Replace("\v", "\r\n", j, 1);
                    j++;
                }
            }
            return str.ToString();
        }

        [STAThread]
        static void Main(string[] args)
        {
            Console.Write("\n\t특정 폴더에 담긴 ppt파일들을 불러와\n\ttxt파일로 추출합니다.\n\n\t사용방법 : ppt파일이 들어있는 폴더를 선택해주세요.\n\n\n\n\n\n\n");
            PrintContribute();

            Application app = new Application();

            String path;

            Microsoft.Office.Core.FileDialog fd =
            app.FileDialog[Microsoft.Office.Core.MsoFileDialogType.msoFileDialogFolderPicker];
            fd.InitialFileName = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\";
            fd.Show();

            if (fd.SelectedItems.Count == 0)
                return;

            path = fd.SelectedItems.Item(1);

            DirectoryInfo id = new DirectoryInfo(path);

            StringBuilder allLiric = new StringBuilder("");

            Console.Write("\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");
            foreach (FileInfo file in id.GetFiles())
            {
                if (((file.Extension.ToLower().CompareTo(".ppt") == 0)
                    || (file.Extension.ToLower().CompareTo(".pptx") == 0)
                    || (file.Extension.ToLower().CompareTo(".pptm") == 0))
                    && (file.Name[0] != '~'))
                {
                    Console.WriteLine("파일 {0}을 여는중...",file.Name);
                    Presentation ppt = app.Presentations.Open(file.FullName,WithWindow:Microsoft.Office.Core.MsoTriState.msoFalse);

                    allLiric.Append(file.Name);
                    allLiric.Append("\r\n");
                    foreach (Slide s in ppt.Slides)
                    {
                        Console.WriteLine("{0} 처리중 : {1}번째 슬라이드",file.Name,s.SlideIndex);
                        foreach (Shape sh in s.Shapes)
                            if (sh.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoFalse)
                            {
                                allLiric.Append(makeCorrectNewline(sh.TextFrame.TextRange.Text));
                                allLiric.Append("\r\n");
                            }
                    }

                    allLiric.Append("∂\r\n");

                    ppt.Close();
                    Console.WriteLine("");
                }
            }

            if (allLiric.Length != 0)
            {
                StreamWriter outputFile = new StreamWriter(path + "\\출력.txt", false);
                outputFile.Write(allLiric);
                outputFile.Close();
            }
            Process.Start("explorer.exe", path);
        }
    }
}
