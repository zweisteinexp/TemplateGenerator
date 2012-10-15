using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace tplgnrtr
{
    class Program
    {
        static int Main(string[] args)
        {
            string sourcePath;
            string sqlPath;
            string asPath;

            try
            {
                XmlDocument config = new XmlDocument();
                config.Load("path.xml");

                sourcePath = config.SelectSingleNode("paths/source").InnerText;
                sqlPath = config.SelectSingleNode("paths/sql").InnerText;
                asPath = config.SelectSingleNode("paths/as").InnerText;
            }
            catch
            {
                Console.WriteLine("Cannot load config file.");
                return 1;
            }

            Generator g = new Generator(sourcePath, sqlPath, asPath);
            g.go();

            return 0;
        }
    }
}
