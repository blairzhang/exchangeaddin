using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Diagnostics;
using System.Reflection;


namespace Radvision.Scopia.ExchangeMeetingAddIn
{
    /// <summary>
    /// Provide log level of messages
    /// </summary>
    public enum LogLevel
    {
        Fatal,
        Exception,
        Error,
        Info,
        Verbose,
        Debug,
        DeepDebug,
        Trace
    }

    /// <summary>
    /// Provide log logical module names
    /// </summary>
    public enum LogModule
    {
        Core,
        DB,
        ICM,
        Network
    }

    /// <summary>
    /// Provide RvLogger mechanism for add-in
    /// </summary>
    /// 




    internal sealed class RvLogger
    {
        /// <summary>
        /// Log file text writer
        /// </summary>
        private static TextWriter _Writer;

        private static bool hasOpened = false;

        private static object theLock = new object();

        /// <summary>
        /// Provide default log file name. 
        /// if filed not exist in the registry will be generate automatically [Log_[DATE]_[TIME].log].
        /// </summary>

        public static string DefaultFileName
        {
            get
            {
                string logDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Log";
                string logFile = logDir + @"\scopia.log";
                return logFile;
            }
        }

        /// <summary>
        /// Provide file operation mode status
        /// </summary>
        public static bool AppendMode
        {
            get
            {
                return true;
            }
        }


        #region ...Public Write methods section...

        /// <summary>
        /// Writes out a formatted string, using the same semantics as String ..::.Format . 
        /// </summary>
        /// <param name="module">The owner module.</param>
        /// <param name="level">Log level.</param>
        /// <param name="format">The formatting string.</param>
        /// <param name="parameters">The object array to write into the formatted string.</param>
        public static void Write(LogModule module, LogLevel level, string format, params object[] parameters)
        {
            if (RvLogger.IsPrintable(level))
                try
                {
                    RvLogger.Write(module, level, String.Format(format, parameters));
                }
                catch (Exception ex)
                {
                    RvLogger.Write(module, level, ex.Message);
                }
        }

        public static void DebugWrite(string text)
        {
            RvLogger.Write(LogModule.ICM, LogLevel.Debug, text);
        }

        public static void InfoWrite(string text)
        {
            RvLogger.Write(LogModule.ICM, LogLevel.Info, text);
        }

        public static void FatalWrite(string text)
        {
            RvLogger.Write(LogModule.ICM, LogLevel.Fatal, text);
        }

        public static void ErrorWrite(string text)
        {
            RvLogger.Write(LogModule.ICM, LogLevel.Error, text);
        }

        /// <summary>
        /// Writes a string to the log.
        /// </summary>
        /// <param name="module">The owner module.</param>
        /// <param name="level">Log level.</param>
        /// <param name="text">The string to write.</param>
        private static void Write(LogModule module, LogLevel level, string text)
        {
            lock (RvLogger.theLock)
            {
                if (RvLogger._Writer != null && RvLogger.IsPrintable(level))
                {
                    using (StringReader reader = new StringReader(text))
                    {
                        string line;
                        while ((line = reader.ReadLine()) != null)
                            RvLogger._Writer.WriteLine("{0} | {1,9} | {2,11} | {3}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss.ff"), level.ToString(), module.ToString(), line);
                    }

                    RvLogger._Writer.Flush();
                }
            }
        }

        #endregion

        #region ...Internal service methods section...

        /// <summary>
        /// Checking if message should be printed to the log file
        /// </summary>
        /// <param name="level">Message log level</param>
        /// <returns>Return true if message should be printed to the log, otherwise false</returns>
        private static bool IsPrintable(LogLevel level)
        {
            return true;
        }


        #endregion

        #region ...Close & Open log file methods section...

        /// <summary>
        /// Create a file for writing UTF-8 encoded log.
        /// </summary>
        public static void Open()
        {
            RvLogger.Open(DefaultFileName, false);
        }

        /// <summary>
        /// Create a file for writing UTF-8 encoded log.
        /// </summary>
        /// <param name="filename">The file to be opened for writing.</param>
        public static void Open(string filename, bool isNew)
        {
            if (filename == null)
                throw new ArgumentNullException("filename");

            if (hasOpened && File.Exists(filename) && !isNew)
            {
                return;
            }
            if (!Directory.Exists(Path.GetDirectoryName(filename)) && !"scopia.log".Equals(filename))
                throw new DirectoryNotFoundException(filename);

            Exception newLogException = null; 
            lock (RvLogger.theLock)
            {
                if (File.Exists(filename) && AppendMode)
                {
                    if (isNew)
                    {
                        try
                        {
                            RvLogger.Close();
                            RvLogger.RenameFiles();
                        }
                        catch (Exception ex) {
                            newLogException = ex;
                        }
                        if (File.Exists(filename))
                            RvLogger._Writer = File.AppendText(filename);
                        else
                            RvLogger._Writer = File.CreateText(filename);
                    }
                    else
                        RvLogger._Writer = File.AppendText(filename);
                }
                else
                {
                    RvLogger._Writer = File.CreateText(filename);
                }
            }

            if (newLogException != null) {
                RvLogger.InfoWrite(newLogException.Message);
                RvLogger.InfoWrite(newLogException.StackTrace);
            }

            hasOpened = true;
        }

        /// <summary>
        /// Create a file for writing UTF-8 encoded log.
        /// </summary>
        /// <param name="filename">The file to be opened for writing.</param>
        public static void RenameFiles()
        {
            string logDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Log";
            DirectoryInfo dir = new DirectoryInfo(logDir);

            FileInfo[] fileInfos = dir.GetFiles("scopia*.log", SearchOption.TopDirectoryOnly);
            for (int i = 0, len = fileInfos.Length; i < len - 1; i++)
            {
                for (int j = 0; j < len - 1 - i; j++)
                {
                    if (fileInfos[j].LastWriteTimeUtc.CompareTo(fileInfos[j + 1].LastWriteTimeUtc) > 0)
                    {
                        FileInfo temp = fileInfos[j];
                        fileInfos[j] = fileInfos[j + 1];
                        fileInfos[j + 1] = temp;
                    }

                }
            }
            for (int i = 0, len = fileInfos.Length; i < len; i++)
            {
                try
                {
                    string destFilePath = logDir + @"\" + "scopia_" + (len - i) + ".log";
                    File.Move(fileInfos[i].FullName, destFilePath);
                    if (i >= 50)
                        fileInfos[i].Delete();
                } catch (Exception ex){
                    RvLogger.InfoWrite(ex.Message);
                    RvLogger.InfoWrite(ex.StackTrace);
                }
            }
        }

        /// <summary>
        /// Closes the current log writer and releases any system resources associated with the log writer.
        /// </summary>
        public static void Close()
        {
            if (RvLogger._Writer != null)
            {
                RvLogger._Writer.Flush();
                RvLogger._Writer.Close();
                RvLogger._Writer = null;
            }
        }

        /// <summary>
        /// Clears all buffers for the current writer and causes any buffered data to be written to the underlying device.
        /// </summary>
        public static void Flush()
        {
            if (RvLogger._Writer != null)
                RvLogger._Writer.Flush();
        }

        public static void OpenLog()
        {
            try
            {
                RvLogger.Open();
            }
            catch (Exception ex)
            {
                try
                {
                    string temp = System.Environment.GetEnvironmentVariable("TEMP");
                    DirectoryInfo info = new DirectoryInfo(temp);
                    string tempPath = info.FullName;
                    StreamWriter streamWriter = File.CreateText(tempPath + "/RvScopiaMeetingAddIn.log");
                    streamWriter.WriteLine(ex.Message);
                    streamWriter.WriteLine(ex.StackTrace);
                    streamWriter.Close();

                    RvLogger.Open(tempPath + "/RvScopiaMeetingAddIn.log", false);
                }
                catch (Exception)
                {

                }
            }
        }

        #endregion
    }
}
