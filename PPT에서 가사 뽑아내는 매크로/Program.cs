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

        [STAThread]
        static void Main(string[] args)
        {
            Console.Write("\n\t특정 폴더에 담긴 ppt파일들을 불러와\n\ttxt파일로 추출합니다.\n\n\n\n\n\n\n\n\n");
            PrintContribute();

            Application app = new Application();

            String path;

            Microsoft.Office.Core.FileDialog fd =
            app.FileDialog[Microsoft.Office.Core.MsoFileDialogType.msoFileDialogFolderPicker];
            fd.InitialFileName = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            fd.Show();

            path = fd.SelectedItems.Item(1);

            DirectoryInfo id = new DirectoryInfo(path);

            StringBuilder allLiric = new StringBuilder("");

            foreach (FileInfo file in id.GetFiles())
            {
                if ((file.Extension.ToLower().CompareTo(".ppt") == 0)
                    || (file.Extension.ToLower().CompareTo(".pptx") == 0)
                    || (file.Extension.ToLower().CompareTo(".pptm") == 0))
                {
                    Presentation ppt = app.Presentations.Open(file.FullName);

                    allLiric.Append(file.Name);
                    foreach (Slide s in ppt.Slides)
                        foreach (Shape sh in s.Shapes)
                            if (sh.HasTextFrame != Microsoft.Office.Core.MsoTriState.msoFalse)
                            {
                                allLiric.Append(sh.TextFrame.TextRange.Text);
                                allLiric.Append("\n");
                            }

                    allLiric.Append("∂\n");

                    ppt.Close();
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
