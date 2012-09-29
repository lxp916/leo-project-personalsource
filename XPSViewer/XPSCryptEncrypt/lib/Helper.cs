using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Text;
using System.Security.Cryptography;
using System.IO;
using System.IO.Packaging;
using System.Windows.Xps.Packaging;
using System.Windows.Xps.Serialization;
using System.Threading;
using System.Windows.Documents;
using System.Windows.Xps;
using System.Windows.Markup;

namespace XPSCryptEncrypt.Lib
{
    public class Helper
    {
        //  Call this function to remove the key from memory after use for security
        [System.Runtime.InteropServices.DllImport("KERNEL32.DLL", EntryPoint = "RtlZeroMemory")]
        public static extern bool ZeroMemory(IntPtr Destination, int Length);

        // Function to Generate a 64 bits Key.
        public static string GenerateKey()
        {
            // Create an instance of Symetric Algorithm. Key and IV is generated automatically.
            DESCryptoServiceProvider desCrypto = (DESCryptoServiceProvider)DESCryptoServiceProvider.Create();

            // Use the Automatically generated key for Encryption. 
            return ASCIIEncoding.ASCII.GetString(desCrypto.Key);
        }

        public static void EncryptFile(string sInputFilename, string sOutputFilename, string sKey)
        {

            DESCryptoServiceProvider DES = new DESCryptoServiceProvider();
            DES.Key = ASCIIEncoding.ASCII.GetBytes(sKey);
            DES.IV = ASCIIEncoding.ASCII.GetBytes(sKey);
            ICryptoTransform desencryptor = DES.CreateEncryptor();

            using (FileStream fsInput = new FileStream(sInputFilename, FileMode.Open, FileAccess.Read))
            {
                using (FileStream fsEncrypted = new FileStream(sOutputFilename, FileMode.Create, FileAccess.Write))
                {
                    using (CryptoStream cryptostream = new CryptoStream(fsEncrypted, desencryptor, CryptoStreamMode.Write))
                    {
                        byte[] bytearrayinput = new byte[fsInput.Length];
                        fsInput.Read(bytearrayinput, 0, bytearrayinput.Length);
                        fsInput.Seek(0, SeekOrigin.Begin);
                        cryptostream.Write(bytearrayinput, 0, bytearrayinput.Length);
                    }
                }
            }
        }
        public static void DecryptFile(string sInputFilename, string sOutputFilename, string sKey)
        {
            DESCryptoServiceProvider DES = new DESCryptoServiceProvider();
            //A 64 bit key and IV is required for this provider.
            //Set secret key For DES algorithm.
            DES.Key = ASCIIEncoding.ASCII.GetBytes(sKey);
            //Set initialization vector.
            DES.IV = ASCIIEncoding.ASCII.GetBytes(sKey);
            //Create a DES decryptor from the DES instance.
            ICryptoTransform desdecrypt = DES.CreateDecryptor();

            //Create a file stream to read the encrypted file back.
            using (FileStream fsInput = new FileStream(sInputFilename, FileMode.Open, FileAccess.Read))
            {
                //Create crypto stream set to read and do a 
                //DES decryption transform on incoming bytes.
                using (CryptoStream cryptostreamDecr = new CryptoStream(fsInput, desdecrypt, CryptoStreamMode.Read))
                {
                    //Print the contents of the decrypted file.
                    //using (FileStream fsDecrypted = new FileStream(sOutputFilename, FileMode.Create, FileAccess.Write))
                    using (StreamWriter fsDecrypted = new StreamWriter(sOutputFilename))
                    {
                        fsDecrypted.Write(new StreamReader(cryptostreamDecr).ReadToEnd());
                    }
                }
            }
        }
        //public static byte[] ToXpsDocument(IEnumerable<FixedPage> pages)
        //{
        //    // XPS DOCUMENTS MUST BE CREATED ON STA THREADS!!!
        //    // Note, this is test code, so I don't care about disposing my memory streams
        //    // You'll have to pay more attention to their lifespan.  You might have to 
        //    // serialize the xps document and remove the package from the package store 
        //    // before disposing the stream in order to prevent throwing exceptions
        //    byte[] retval = null;
        //    Thread t = new Thread(new ThreadStart(() =>
        //    {
        //        // A memory stream backs our document
        //        MemoryStream ms = new MemoryStream(2048);
        //        // a package contains all parts of the document
        //        Package p = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite);
        //        // the package store manages packages
        //        Uri u = new Uri("pack://TemporaryPackageUri.xps");
        //        PackageStore.AddPackage(u, p);
        //        // the document uses our package for storage
        //        XpsDocument doc = new XpsDocument(p, CompressionOption.NotCompressed, u.AbsoluteUri);
        //        // An xps document is one or more FixedDocuments containing FixedPages
        //        FixedDocument fDoc = new FixedDocument();
        //        PageContent pc;
        //        foreach (var fp in pages)
        //        {
        //            // this part of the framework is weak and hopefully will be fixed in 4.0
        //            pc = new PageContent();
        //            ((IAddChild)pc).AddChild(fp);
        //            fDoc.Pages.Add(pc);
        //        }
        //        // we use the writer to write the fixed document to the xps document
        //        XpsDocumentWriter writer;
        //        writer = XpsDocument.CreateXpsDocumentWriter(doc);
        //        // The paginator controls page breaks during the writing process
        //        // its important since xps document content does not flow 
        //        writer.Write(fDoc.DocumentPaginator);
        //        // 
        //        p.Flush();

        //        // this part serializes the doc to a stream so we can get the bytes
        //        ms = new MemoryStream();
        //        var xpswriter = new XpsSerializerFactory().CreateSerializerWriter(ms);
        //        xpswriter.Write(doc.GetFixedDocumentSequence());

        //        retval = ms.ToArray();
        //    }));
        //    // Instantiating WPF controls on a MTA thread throws exceptions
        //    t.SetApartmentState(ApartmentState.STA);
        //    // adjust as needed
        //    t.Priority = ThreadPriority.AboveNormal;
        //    t.IsBackground = false;
        //    t.Start();
        //    //~five seconds to finish or we bail
        //    int milli = 0;
        //    //while (buffer == null && milli++ < 5000)
        //    //    Thread.Sleep(1);
        //    //Ditch the thread
        //    if (t.IsAlive)
        //        t.Abort();
        //    // If we time out, we return null.
        //    return retval;
        //}
        private static MemoryStream Encrypt_Read(ICryptoTransform ict, byte[] data)
        {
            using (var ms = new MemoryStream(data))
            using (var cstream = new CryptoStream(ms, ict, CryptoStreamMode.Read))
            using (var destMs = new MemoryStream())
            {
                byte[] buffer = new byte[100];
                int readLen;

                while ((readLen = cstream.Read(buffer, 0, 100)) > 0)
                    destMs.Write(buffer, 0, readLen);
                return destMs;
            }
        }
        private static MemoryStream Encrypt_Write(ICryptoTransform ict, byte[] data)
        {
            using (var ms = new MemoryStream())
            using (var cstream = new CryptoStream(ms, ict, CryptoStreamMode.Write))
            {
                cstream.Write(data, 0, data.Length);
                return ms;
            }
        }
    }
}