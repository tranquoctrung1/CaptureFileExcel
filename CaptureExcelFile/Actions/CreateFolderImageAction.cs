using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaptureExcelFile.Actions
{
    public class CreateFolderImageAction
    {
        public void CreateFolderImage(string pathToSaveFileImage, string folder)
        {
            
            if(!Directory.Exists(pathToSaveFileImage + "\\" + folder)) {
                Directory.CreateDirectory(pathToSaveFileImage + "\\" + folder);
            }
        }
    }
}
