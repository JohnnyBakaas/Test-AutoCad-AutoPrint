using System.Runtime.InteropServices;

namespace AutoCADBatchPlotter
{
    class Program
    {
        static async Task Main(string[] args)
        {
            string folderPath = @"C:\Users\johnn\OneDrive\Skrivebord\MSG";

            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine("The provided folder path does not exist.");
                return;
            }

            string[] dwgFiles = Directory.GetFiles(folderPath, "*.dwg");

            if (dwgFiles.Length == 0)
            {
                Console.WriteLine("No DWG files found in the specified folder.");
                return;
            }

            int maxRetries = 5;
            int retryDelayMilliseconds = 2000;
            bool success = false;

            for (int retry = 0; retry < maxRetries && !success; retry++)
            {
                try
                {
                    Type acadType = Type.GetTypeFromProgID("AutoCAD.Application.24");
                    dynamic acadApp = Activator.CreateInstance(acadType, true);
                    acadApp.Visible = false;

                    // Close the default drawing (Drawing1)
                    acadApp.ActiveDocument.Close(false);

                    foreach (string dwgFile in dwgFiles)
                    {
                        bool fileProcessed = false;
                        int fileRetry = 0;
                        while (!fileProcessed && fileRetry < maxRetries)
                        {
                            try
                            {
                                Console.WriteLine($"Opening {dwgFile}...");
                                dynamic doc = acadApp.Documents.Open(dwgFile, false);

                                string outputFileName = Path.ChangeExtension(dwgFile, "pdf");
                                Console.WriteLine($"Publishing {outputFileName}...");

                                // Publish the drawing using default settings
                                doc.Plot.PlotToFile(outputFileName, "DWG To PDF.pc3");

                                Console.WriteLine($"Closing {dwgFile}...");
                                doc.Close(false);
                                fileProcessed = true;
                            }
                            catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x8001010A))
                            {
                                fileRetry++;
                                Console.WriteLine($"Error processing {dwgFile} (attempt {fileRetry} of {maxRetries}): {ex.Message}. Retrying...");
                                await Task.Delay(retryDelayMilliseconds);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error processing {dwgFile}: {ex.Message}");
                                break;
                            }
                        }
                    }

                    acadApp.Quit();
                    Console.WriteLine("Finished publishing all drawings.");
                    success = true;
                }
                catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x8001010A))
                {
                    Console.WriteLine($"Error initializing AutoCAD (attempt {retry + 1} of {maxRetries}): {ex.Message}. Retrying...");
                    await Task.Delay(retryDelayMilliseconds);
                }
                catch (COMException ex)
                {
                    Console.WriteLine($"Error initializing AutoCAD: {ex.Message}");
                    break;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected error: {ex.Message}");
                    break;
                }
            }

            if (!success)
            {
                Console.WriteLine("Failed to initialize AutoCAD after all retries.");
            }
        }
    }
}
