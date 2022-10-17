using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Check_NOAA_Outlook
{
    public class Logging : IDisposable
    {
        public Logging(string LogFile)
        {
            m_LogFile = LogFile;
            m_StopWatch = new System.Diagnostics.Stopwatch();

            System.Threading.Thread t = new(FlushTimerLoop)
            {
                IsBackground = true
            };
            t.Start();
        }

        private FileStream m_FS;
        private readonly System.Diagnostics.Stopwatch m_StopWatch;
        private bool m_StopFlush;
        private bool m_CanNotWriteToFile;
        private bool disposed = false;
        private readonly object lock_Main = new();

        ~Logging()
        {
            if (!this.disposed) Dispose();
        }

        public void Dispose()
        {
            Dispose(true);
            // This object will be cleaned up by the Dispose method.
            // Therefore, you should call GC.SupressFinalize to
            // take this object off the finalization queue
            // and prevent finalization code for this object
            // from executing a second time.
            GC.SuppressFinalize(this);
            this.disposed = true;
        }

        protected virtual void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!this.disposed)
            {
                // If disposing equals true, dispose all managed
                // and unmanaged resources.
                if (disposing)
                {
                    // Dispose managed resources.

                }

                // Call the appropriate methods to clean up
                // unmanaged resources here.
                // If disposing is false,
                // only the following code is executed.

                // Note disposing has been done.
                disposed = true;
            }
            m_StopFlush = true;
            CloseFile();
        }

        private bool OpenFile()
        {
            if (m_FS != null)
            {
                lock (lock_Main)
                {
                    if (m_StopWatch.IsRunning) m_StopWatch.Reset();
                    m_FS.Close();
                }
            }

            try
            {
                m_FS = new FileStream(m_LogFile, FileMode.Append, FileAccess.Write, FileShare.Read);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to write to {0}, {1}", m_LogFile, ex.Message);
                m_CanNotWriteToFile = true;
                return false;
            }

            return true;
        }

        public string LogFile
        {
            get { return m_LogFile; }
            set
            {
                bool FileChanged = !string.Equals(m_LogFile, value, StringComparison.CurrentCultureIgnoreCase);
                m_LogFile = value;
                if (FileChanged)
                {
                    CloseFile();
                    m_CanNotWriteToFile = false;
                }
            }
        }

        private void CloseFile()
        {
            lock (lock_Main)
            {
                if (m_StopWatch.IsRunning) m_StopWatch.Reset();
            }

            if (m_FS != null && m_FS.CanWrite)
            {
                //m_FS.Flush(true);
                m_FS.Close();
            }
        }

        private void FlushTimerLoop()
        {
            while (!m_StopFlush)
            {
                if (!m_StopWatch.IsRunning)
                {
                    System.Threading.Thread.Sleep(5);
                    continue;
                }

                if (m_StopWatch.ElapsedMilliseconds >= 1000 && !m_StopFlush)
                {
                    lock (lock_Main)
                    {
                        if (m_FS.CanWrite) m_FS.Flush();
                        m_StopWatch.Stop();
                    }
                }
            } // next loop
        }

        public bool CheckIfRotationNeeded(int MaxRollovers, int MaxSizeMB, out string Message)
        {
            Message = null;

            if (MaxSizeMB <= 0 || MaxRollovers < 1) return false;

            if (!File.Exists(m_LogFile)) return false;

            FileInfo FI = new(m_LogFile);
            if (FI.Length < 1024) return false;
            double FileSizeMB = (double)FI.Length / 1024.0 / 1024.0;
            if (FileSizeMB < (double)MaxSizeMB)
            {
                Message = string.Format("File size {0:#,0.##} was under the limit of {1:#,0}", FileSizeMB, MaxSizeMB);
                return false;
            }

            // has met or exceeded max size, do rotation
            //TODO: truncate current file not supported yet

            bool DidDelete = false;
            int Rotations = 0;

            for (int i = MaxRollovers - 2; i >= -1; i--)
            {
                string Filename = i == -1 ?
                    m_LogFile : string.Format("{0}.{1}", m_LogFile, i);
                bool FileExists = System.IO.File.Exists(Filename);

                if (FileExists)
                {
                    if (i == MaxRollovers - 2)
                    {
                        // last log in rotation, delete it
                        try { System.IO.File.Delete(Filename); DidDelete = true; } catch { }
                    }
                    else
                    {
                        // rename to next number in series
                        string NewFilename = string.Format("{0}.{1}", m_LogFile, (i + 1));
                        try { System.IO.File.Move(Filename, NewFilename); Rotations++; } catch { }
                    }
                }
            }

            Message = string.Format("Rotated {0} file(s)", Rotations);
            if (DidDelete) Message += ", deleted oldest file";

            return true;
        }

        private string m_LogFile;
        public WhichLogType m_LogLevel = WhichLogType.DEBUG;

        public bool IsInfoEnabled
        {
            get { return (m_LogLevel >= WhichLogType.INFO); }
        }
        public bool IsWarnEnabled
        {
            get { return (m_LogLevel >= WhichLogType.WARN); }
        }

        public bool IsFatalEnabled
        {
            get { return (m_LogLevel >= WhichLogType.FATAL); }
        }
        public bool IsDebugEnabled
        {
            get { return (m_LogLevel >= WhichLogType.DEBUG); }
        }
        public bool IsErrorEnabled
        {
            get { return (m_LogLevel >= WhichLogType.ERROR); }
        }

        public void Info(string Text)
        {
            WriteLog(WhichLogType.INFO, Text);
        }
        public void InfoFormat(string Text, params object[] Args)
        {
            WriteLog(WhichLogType.INFO, string.Format(Text, Args));
        }

        public void Info(string Text, Exception ex)
        {
            WriteLog(WhichLogType.INFO, Text, ex);
        }

        public void Error(string Text)
        {
            WriteLog(WhichLogType.ERROR, Text);
        }

        public void Error(string Text, Exception ex)
        {
            WriteLog(WhichLogType.ERROR, Text, ex);
        }

        public void ErrorFormat(string Text, params object[] Args)
        {
            WriteLog(WhichLogType.ERROR, string.Format(Text, Args));
        }
        public void FatalFormat(string Text, params object[] Args)
        {
            WriteLog(WhichLogType.FATAL, string.Format(Text, Args));
        }

        public void Fatal(string Text)
        {
            WriteLog(WhichLogType.FATAL, Text);
        }

        public void Fatal(string Text, Exception ex)
        {
            WriteLog(WhichLogType.FATAL, Text, ex);
        }

        public void Warn(string Text)
        {
            WriteLog(WhichLogType.WARN, Text);
        }

        public void Warn(string Text, Exception ex)
        {
            WriteLog(WhichLogType.WARN, Text, ex);
        }

        public void WarnFormat(string Text, params object[] Args)
        {
            WriteLog(WhichLogType.WARN, string.Format(Text, Args));
        }

        public void Debug(string Text)
        {
            WriteLog(WhichLogType.DEBUG, Text);
        }

        public void DebugFormat(string Text, params object[] Args)
        {
            WriteLog(WhichLogType.DEBUG, string.Format(Text, Args));
        }

        public void Debug(string Text, Exception ex)
        {
            WriteLog(WhichLogType.DEBUG, Text, ex);
        }

        public void WriteLog(WhichLogType LogType, string Text, Exception ex = null)
        {
            if (m_CanNotWriteToFile) return;
            if (m_LogLevel < LogType) return;

            string ExtraInfo = string.Format("{0:yyyy-MM-dd HH:mm:ss.fff} [{1}] {2} - ",
                DateTime.Now, Environment.CurrentManagedThreadId, LogType);

            // open the file on the first write
            if (m_FS == null && !OpenFile()) return;

            lock (lock_Main)
            {
                try
                {
                    m_FS.Write(Encoding.Default.GetBytes(ExtraInfo + Text + Environment.NewLine));
                    //m_FS.Flush();
                    if (ex != null) m_FS.Write(Encoding.Default.GetBytes(ex.ToString() + Environment.NewLine));
                }
                catch (Exception ex2)
                {
                    Console.WriteLine("Unable to write to {0}, {1}", m_LogFile, ex2.Message);
                    m_CanNotWriteToFile = true;
                    return;
                }
                if (!m_StopWatch.IsRunning) m_StopWatch.Restart();
            }
        }
    }

    public enum WhichLogType
    {
        NONE = 0,
        FATAL = 1,
        ERROR = 2,
        WARN = 3,
        INFO = 4,
        DEBUG = 5
    }
}
