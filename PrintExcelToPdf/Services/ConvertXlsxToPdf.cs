using System.Diagnostics;

namespace PrintExcelToPdf.Services
{
    public class ConvertXlsxToPdf
    {
        public MemoryStream Convert(MemoryStream xlsxStream)
        {
            // Créer un fichier temporaire pour enregistrer le fichier XLSX
            var tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
            var outputDirectory = Path.Combine(Path.GetTempPath(), "converted_pdfs");

            // Créer le répertoire de sortie pour le fichier PDF s'il n'existe pas
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            // Enregistrer le fichier XLSX dans un fichier temporaire
            using (var fileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
            {
                xlsxStream.CopyTo(fileStream);
            }

            // Spécifier le nom de fichier PDF de sortie
            var outputPdfPath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(tempFilePath) + ".pdf");

            // Détecter la plateforme et ajuster le chemin vers l'exécutable LibreOffice
            string libreOfficePath = GetLibreOfficePath();

            // Construire la commande pour exécuter LibreOffice en mode headless
            var arguments = $"--headless --convert-to pdf --outdir \"{outputDirectory}\" \"{tempFilePath}\"";

            // Exécution de la commande LibreOffice
            var processStartInfo = new ProcessStartInfo
            {
                FileName = libreOfficePath,
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            try
            {
                using (var process = Process.Start(processStartInfo))
                {
                    process.WaitForExit();
                }

                // Vérifier si le fichier PDF a été généré
                if (System.IO.File.Exists(outputPdfPath))
                {
                    // Lire le fichier PDF généré dans un MemoryStream
                    var pdfMemoryStream = new MemoryStream();
                    using (var fileStream = new FileStream(outputPdfPath, FileMode.Open, FileAccess.Read))
                    {
                        fileStream.CopyTo(pdfMemoryStream);
                    }

                    // Retourner le MemoryStream contenant le fichier PDF
                    pdfMemoryStream.Seek(0, SeekOrigin.Begin);
                    return pdfMemoryStream;
                }
                else
                {
                    throw new Exception("Erreur lors de la conversion du fichier.");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Erreur interne : {ex.Message}");
            }
            finally
            {
                // Nettoyage des fichiers temporaires
                if (System.IO.File.Exists(tempFilePath))
                {
                    System.IO.File.Delete(tempFilePath);
                }
            }
        }

        private string GetLibreOfficePath()
        {
            // Détecter la plateforme et ajuster le chemin vers LibreOffice
            if (OperatingSystem.IsWindows())
            {
                // Chemin pour Windows
                return @"C:\Program Files\LibreOffice\program\soffice.exe"; // Assurez-vous que le chemin est correct
            }
            else if (OperatingSystem.IsLinux())
            {
                // Sur Linux, le binaire est généralement dans le PATH
                return "libreoffice";
            }
            else if (OperatingSystem.IsMacOS())
            {
                // Sur macOS, le binaire est généralement dans /Applications/LibreOffice.app/Contents/MacOS/soffice
                return "/Applications/LibreOffice.app/Contents/MacOS/soffice";
            }
            else
            {
                throw new PlatformNotSupportedException("Le système d'exploitation n'est pas supporté.");
            }
        }
    }
}
