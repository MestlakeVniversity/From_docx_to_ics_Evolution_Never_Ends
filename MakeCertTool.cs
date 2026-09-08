using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;

class MakeCertTool
{
    static int Main(string[] args)
    {
        try
        {
            string pfxPath = args.Length > 0 ? args[0] : "KejianToIcs.pfx";
            string cerPath = args.Length > 1 ? args[1] : "KejianToIcs.cer";
            string pwd = args.Length > 2 ? args[2] : "KejianToIcs";

            using (var rsa = RSA.Create(2048))
            {
                var req = new CertificateRequest(
                    "CN=KejianToIcs, O=KejianToIcs, C=CN",
                    rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

                req.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, false));
                req.CertificateExtensions.Add(new X509BasicConstraintsExtension(false, false, 0, false));

                var eku = new OidCollection();
                eku.Add(new Oid("1.3.6.1.5.5.7.3.3")); // Code Signing
                req.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(eku, false));

                var cert = req.CreateSelfSigned(DateTimeOffset.UtcNow.AddDays(-1), DateTimeOffset.UtcNow.AddYears(5));

                byte[] pfx = cert.Export(X509ContentType.Pfx, pwd);
                File.WriteAllBytes(pfxPath, pfx);

                byte[] cer = cert.Export(X509ContentType.Cert);
                File.WriteAllBytes(cerPath, cer);
            }
            return 0;
        }
        catch (Exception ex)
        {
            try { File.WriteAllText("MakeCertTool.err", ex.ToString(), new UTF8Encoding(false)); }
            catch { }
            return 1;
        }
    }
}
