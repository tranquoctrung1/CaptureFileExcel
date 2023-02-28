using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaptureExcelFile.Actions
{
    public class DeleteFolderAction
    {
        public void DeleteFolder(string path, string folder)
        {
            if (Directory.Exists(path + "\\" + folder))
            {
                Directory.Delete(path + "\\" + folder, true);
            }
        }
    }
}
