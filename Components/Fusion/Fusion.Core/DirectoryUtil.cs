// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DirectoryUtil.cs" company="Advanced Health and Care Limited">
//   Copyright © 2012 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the directory utility class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Fusion.Core
{
    using System.IO;

    public class DirectoryUtil
    {
        /// <summary>
        /// Make sure that a given directory (path) exists
        /// </summary>
        /// <param name="path"> Full pathname of the directory. </param>
        public static void EnforceDirectory(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
    }
}
