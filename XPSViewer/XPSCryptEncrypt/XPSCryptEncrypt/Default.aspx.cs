using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using XPSCryptEncrypt.Lib;
using System.IO;

namespace XPSCryptEncrypt
{
    public partial class _Default : System.Web.UI.Page
    {
        string key = "";
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        protected void btnCrpt_Click(object sender, EventArgs e)
        {
        }

        protected void btnEncrpt_Click(object sender, EventArgs e)
        {
        }

        protected void button_Click(object sender, EventArgs e)
        {
            key = Helper.GenerateKey();

            string orginal = @"C:\Users\liaxiaop\Desktop\test_org.txt";
            string encryptFile = @"C:\Users\liaxiaop\Desktop\test_ecrpt.txt";
            string decryptFile = @"C:\Users\liaxiaop\Desktop\test_crpt.txt";
            Helper.EncryptFile(orginal, encryptFile, key);
            Helper.DecryptFile(encryptFile, decryptFile, key);
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            string source = @"C:\Users\liaxiaop\Desktop\test.pdf";
            string target = @"C:\Users\liaxiaop\Desktop\test.xps";

            ConverterProxy converter = new ConverterProxy(Path.GetExtension(source).Replace(".",""));
            converter.Convert(source, target);
        }
    }
}
