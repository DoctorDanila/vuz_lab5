using System;
using System.IO;

namespace LabWork
{
    public class Logger
    {
        private StreamWriter writer;

        public Logger(string filePath, bool append)
        {
            try
            {
                writer = new StreamWriter(filePath, append);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при открытии файла логов: {ex.Message}");
                writer = new StreamWriter("log.txt", append); 
            }
        }

        public void Log(string message)
        {
            writer.WriteLine($"{DateTime.Now}: {message}");
        }

        public void Close()
        {
            writer.Close();
        }
    }
}