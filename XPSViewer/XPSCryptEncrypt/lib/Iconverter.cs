using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace XPSCryptEncrypt.Lib
{
    public interface Iconverter
    {
         bool Convert(string sourcePath, string targetPath);
    }
}