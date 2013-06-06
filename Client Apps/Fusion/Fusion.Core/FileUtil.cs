
namespace Fusion.Core
{
    using System.IO;

    public class FileUtil
    {
        /// <summary>
        /// Performs a File.Move, but if destination filename is already used will add numeric extension until found
        /// </summary>
        /// <param name="source"> Source for the. </param>
        /// <param name="dest">   Destination for the. </param>
        static public void SafeMove(string source, string dest) {


            string targetFile = dest;

            int counter = 0;

            for(;;) {
                
                if (!File.Exists(targetFile)) {
                    // Possible race condition here for high concurrency systems could be avoided by creating the file here and checking for success, but
                    // move function below would need to overwrite (be a copy/delete?)
                                
                    break;
                }
            
                targetFile = dest + (++counter).ToString();
            }


            File.Move(source, targetFile);            
        }
    }
}
