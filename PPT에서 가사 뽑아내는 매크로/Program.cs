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
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application app = new Application();

            String path;

            Microsoft.Office.Core.FileDialog fd =
            app.FileDialog[Microsoft.Office.Core.MsoFileDialogType.msoFileDialogFolderPicker];
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
