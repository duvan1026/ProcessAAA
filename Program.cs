using System;
//using System.Linq;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.AccessControl;
using System.Security.Principal;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using TesseractOCR.Library.src;

namespace TripleAAA
{
    internal class Program
    {
        #region Constantes

        const string tessdataPath = @"C:\Program Files (x86)\Tesseract-OCR\tessdata";
        const string language = "spa";
        const double dpi = 96.0;                                  // Resolución estándar de pantalla

        const string inputFile = @"C:\Users\duvan.castro\Desktop\TestPDFText\CAJA.30"; // Reemplaza con la ruta de tu carpeta
        const string outputFile = @"C:\Users\duvan.castro\Desktop\TestPDFText\CAJA.30\OutputData";


        //const string nameDirectoryDestination = "Data.Process";
        const string inputFormat = "*.tif";
        const string outputFormat = ".pdf";
        const string outputInformationFormat = ".txt";
        const string filterSuffix = ".#";
        const string delimiter = "-";
        const string informationName = "informacion";

        private static string nameFolderDestination = null;     // Nombre carpeta de destino
        private static string boxFolderName = null;             // Nombre de la carpeta principal (Caja)
        private static string bookFolderName = null;            // Nombre de la carpeta secundaria (Libro)
        private static string expedienteFolderName = null;      // Nombre de la carpeta expediente (Expediente)
        private static string imageFolderName = null;           // Nombre de la carpeta Imagen (Image)



        #endregion

        static void Main(string[] args)
        {
            // Iniciar el cronómetro para medir el tiempo de ejecución
            Stopwatch stopwatch = Stopwatch.StartNew();

            // TODO: Agregar validacion si la carpeta ya fue procesada ".#", si ya lo fue proceder a enviar mensaje indicando que esta carpeta ya fue procesada y no puede realizarse el proceso
            boxFolderName = Path.GetFileName(inputFile);  // Nombre Carpeta de Caja
            nameFolderDestination = Path.GetFileName(outputFile);

            // Crear el directorio de destino para el archivo de salida
            string outputFileDestination = Path.Combine(outputFile, boxFolderName);
            CreateDirectoryWithWriteAccess(outputFileDestination);

            // Filtrar y recorrer los Book solo que cumplen con las condiciónes
            DirectoryInfo inputBoxDirectory = new DirectoryInfo(inputFile);          
            foreach (var currentBookFolder in inputBoxDirectory.GetDirectories()
                                            .Where(dir => !dir.Name.EndsWith(filterSuffix) && dir.Name != nameFolderDestination))
            {

                bookFolderName = Path.GetFileName(currentBookFolder.FullName);  // Nombre del libro

                // Filtrar y recorrer los Expedientes cumplen con la condición
                DirectoryInfo inputBookDirectory = new DirectoryInfo(currentBookFolder.FullName);                
                foreach (var currentExpedienteFolder in inputBookDirectory.GetDirectories().Where(dir => !dir.Name.EndsWith(filterSuffix)))
                {
                    expedienteFolderName = Path.GetFileName(currentExpedienteFolder.FullName);  // Nombre del expediente 

                    ProcessTiffFiles(currentExpedienteFolder.FullName, outputFileDestination);
                }
            }

            // Detener el cronómetro y calcular el tiempo transcurrido
            stopwatch.Stop();
            TimeSpan elapsedTime = stopwatch.Elapsed;
            Console.WriteLine($"Proceso completado en {elapsedTime.TotalMinutes} minutos {elapsedTime.Seconds} segundos");
            Console.WriteLine("Proceso completado. El PDF de texto se ha guardado.");
            Console.ReadLine();
        }

        /// <summary>
        /// Procesa archivos TIFF en un directorio de entrada y los convierte en un documento PDF de destino.
        /// </summary>
        /// <param name="directoryInput">El directorio de entrada que contiene archivos TIFF a procesar.</param>
        /// <param name="destinationRoute">El directorio de destino donde se almacenará el documento PDF resultante.</param>
        static void ProcessTiffFiles(string directoryInput, string destinationRoute)
        {
            // Obtener archivos TIFF en el directorio actual
            string[] tiffFiles = GetTiffFilesInDirectory(directoryInput)
                        .OrderBy(fileName => fileName)
                        .ToArray();

            if (tiffFiles == null || tiffFiles.Length == 0) return;

            string outputnamePDFTotal = GetOutputPdfName(destinationRoute);
            string outputInformationPath = GetTXTOutputPath(destinationRoute);

            using (PdfDocument outputDocument = CreatePdfDocument())
            {

                foreach (string tiffFile in tiffFiles)
                {
                    // TODO: OPCION 1:
                    // Recorre todos los files y guarda uno por 1
                    imageFolderName = Path.GetFileNameWithoutExtension(tiffFile);  // Nombre Carpeta de Caja
                    string outputPath = GetPdfOutputPath(destinationRoute);
                    PDFService.ConvertTiffToPdf(tiffFile, outputPath);
                    // Crear o abrir el archivo para escritura
                    
                    //string imageName = Path.GetFileName(outputPath);    // Obtener el nombre del archivo
                    //DateTime fechaLog = DateTime.Now;                   // Obtener la fecha actual
                    //string tamañoFormateado = GetFormattedSize(outputPath);
                    // Escribir información en el archivo de texto
                    WriteInformationToFile(outputPath, outputInformationPath);

                    //long tamañoEnBytes = ObtenerTamañoArchivo(outputPath);
                    //string tamañoFormateado = FormatearTamañoArchivo(tamañoEnBytes);// Formatear el tamaño del archivo
                    //using (StreamWriter sw = File.AppendText(outputInformationPath))
                    //{                        
                    //    sw.WriteLine($"{imageName}\t{fechaLog}\t{tamañoFormateado}"); // Escribir la información en el archivo
                    //}

                    // TODO: OPCION 2:
                    // Recorre todos los files y guarda uno por 1 en la ruta Temp para construir el PDF con todos
                    string tempFolderPath = Path.GetTempPath();
                    string tempPath = GetPdfOutputPath(tempFolderPath);
                    PDFService.ConvertTiffToPdf(tiffFile, tempPath);
                    AddPagesFromTiffToPdf(tempPath, outputDocument);
                    File.Delete(tempPath);                                // Eliminar la imagen en la ruta temp
                }
                SavePdfDocument(outputDocument, outputnamePDFTotal);
                WriteInformationToFile(outputnamePDFTotal, outputInformationPath);

                // TODO: Cambiar nombre d ela carpeta  aña cual se extrajo la informacion
            }

        }

        // Escribir información en el archivo de texto
        static void WriteInformationToFile(string outputPath, string outputInformationPath)
        {
            string imageName = Path.GetFileName(outputPath);    // Obtener el nombre del archivo
            DateTime fechaLog = DateTime.Now;                   // Obtener la fecha actual
            string tamañoFormateado = GetFormattedSize(outputPath);

            using (StreamWriter sw = File.AppendText(outputInformationPath))
            {
                sw.WriteLine($"{imageName}\t{fechaLog}\t{tamañoFormateado}");
            }
        }

        // Formatear el tamaño del archivo en KB o MB
        static string GetFormattedSize(string outputPath)
        {
            long tamañoEnBytes = ObtenerTamañoArchivo(outputPath);

            const long kilobyte = 1024;
            const long megabyte = kilobyte * 1024;

            double tamañoFormateado;

            if (tamañoEnBytes >= megabyte)
            {
                tamañoFormateado = Math.Round((double)tamañoEnBytes / megabyte, 1);
                return $"{tamañoFormateado} MB";
            }
            else
            {
                tamañoFormateado = Math.Round((double)tamañoEnBytes / kilobyte, 1);
                return $"{tamañoFormateado} KB";
            }
        }



        /// <summary>
        /// Agrega las páginas de un archivo TIFF a un documento PDF de destino.
        /// </summary>
        /// <param name="tiffFilePath">La ruta del archivo TIFF del cual se extraerán las páginas.</param>
        /// <param name="targetDocument">El documento PDF de destino al cual se agregarán las páginas del archivo TIFF.</param>
        static void AddPagesFromTiffToPdf(string tiffFilePath, PdfDocument targetDocument)
        {
            using (PdfDocument inputDocument = PdfReader.Open(tiffFilePath, PdfDocumentOpenMode.Import))
            {
                int pageCounter = inputDocument.PageCount;
                for (int pageIndex = 0; pageIndex < pageCounter; pageIndex++)
                {
                    PdfPage page = inputDocument.Pages[pageIndex];
                    targetDocument.AddPage(page);
                }
            }
        }

        static long ObtenerTamañoArchivo(string rutaArchivo)
        {
            // Crear un objeto FileInfo
            FileInfo fileInfo = new FileInfo(rutaArchivo);

            // Obtener el tamaño del archivo en bytes
            return fileInfo.Length;
        }

        static string FormatearTamañoArchivo(long tamañoEnBytes)
        {
            const long kilobyte = 1024;
            const long megabyte = kilobyte * 1024;

            double tamañoFormateado;

            if (tamañoEnBytes >= megabyte)
            {
                // Convertir a megabytes con un decimal
                tamañoFormateado = Math.Round((double)tamañoEnBytes / megabyte, 1);
                return $"{tamañoFormateado} MB";
            }
            else
            {
                // Convertir a kilobytes con un decimal
                tamañoFormateado = Math.Round((double)tamañoEnBytes / kilobyte, 1);
                return $"{tamañoFormateado} KB";
            }
        }

        /// <summary>
        /// Guarda un documento PDF.
        /// </summary>
        /// <param name="document">El objeto PdfDocument que se va a guardar y cerrar.</param>
        /// <param name="outputPath">La ruta de salida donde se guardará el documento PDF.</param>
        static void SavePdfDocument(PdfDocument document, string outputPath)
        {
            document.Save(outputPath);
        }

        /// <summary>
        /// Crea un nuevo documento PDF.
        /// </summary>
        /// <returns>Un objeto PdfDocument que se utilizará como documento PDF de destino.</returns>
        static PdfDocument CreatePdfDocument()
        {
            return new PdfDocument();
        }

        static string GetTXTOutputPath(string destinationRoute)
        {
            return Path.Combine(destinationRoute, informationName + outputInformationFormat);
        }

        /// <summary>
        /// Obtiene la ruta de salida del archivo PDF basada en la ruta de destino y el archivo TIFF de entrada.
        /// </summary>
        /// <returns>La ruta del archivo PDF de salida.</returns>
        static string GetPdfOutputPath(string destinationRoute)
        {
            return Path.Combine(destinationRoute, bookFolderName + delimiter + 
                                                  expedienteFolderName + delimiter +
                                                  imageFolderName + outputFormat);
        }

        /// <summary>
        /// Obtiene el nombre del archivo PDF de salida basado en el directorio de destino.
        /// </summary>
        /// <param name="destinationRoute">El directorio de destino para el archivo PDF.</param>
        /// <returns>El nombre del archivo PDF de salida.</returns>
        static string GetOutputPdfName(string destinationRoute)
        {
            return Path.Combine(destinationRoute, bookFolderName + delimiter +
                                                  expedienteFolderName + outputFormat);
        }

        /// <summary>
        /// Obtiene archivos TIFF en un directorio de entrada.
        /// </summary>
        /// <param name="directoryInput">El directorio de entrada que contiene archivos TIFF.</param>
        /// <returns>Un arreglo de rutas de archivo TIFF encontrados en el directorio de entrada.</returns>
        static string[] GetTiffFilesInDirectory(string directoryInput)
        {
            return Directory.GetFiles(directoryInput, inputFormat);
        }

        /// <summary>
        /// Verifica si un directorio existe y lo crea si no existe. Luego, asigna permisos de escritura al directorio.
        /// </summary>
        /// <param name="directoryPath">La ruta del directorio a verificar y crear si es necesario.</param>
        public static void CreateDirectoryWithWriteAccess(string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
                AssignWritePrivilegesToDirectory(directoryPath);
            }
        }

        /// <summary>
        /// Agrega permisos de escritura a una carpeta.
        /// </summary>
        /// <param name="folderPath">La ruta de la carpeta a la que se le asignarán permisos de escritura.</param>
        static void AssignWritePrivilegesToDirectory(String folderPath)
        {
            DirectoryInfo newFileInfo = new DirectoryInfo(folderPath);               // Obtener información de la carpeta.
            DirectorySecurity newFileSecurity = newFileInfo.GetAccessControl();      // Obtener el control de acceso actual de la carpeta.
            FileSystemAccessRule writeRule = new FileSystemAccessRule(               // Crear una regla de acceso para permitir la escritura a todos los usuarios.
                new SecurityIdentifier(WellKnownSidType.BuiltinUsersSid, null),
                FileSystemRights.Write,
                InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                PropagationFlags.None,
                AccessControlType.Allow
            );
            newFileSecurity.AddAccessRule(writeRule);                                // Agregar la regla de acceso al control de acceso.
            newFileInfo.SetAccessControl(newFileSecurity);                           // Establecer el nuevo control de acceso en la carpeta.
        }
    }
}
