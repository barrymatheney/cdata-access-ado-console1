using System;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            CompactAccessDB("", "");
        }



        /// <summary>
        /// MBD compact method (c) 2004 Alexander Youmashev
        /// !!IMPORTANT!!
        /// !make sure there's no open connections
        ///    to your db before calling this method!
        /// !!IMPORTANT!!
        /// </summary>
        /// <param name="connectionString">connection string to your db</param>
        /// <param name="mdwfilename">FULL name
        ///     of an MDB file you want to compress.</param>
        public static void CompactAccessDB(string connectionString, string mdwfilename)
        {
            object[] oParams;

            //create an instance of a Jet Replication Object
            object objJRO =
              Activator.CreateInstance(Type.GetTypeFromProgID("JRO.JetEngine"));

            //filling Parameters array
            //change "Jet OLEDB:Engine Type=5" to an appropriate value
            // or leave it as is if you db is JET4X format (access 2000,2002)
            //(yes, jetengine5 is for JET4X, no misprint here)
            oParams = new object[] {
        connectionString,
        "Provider=Microsoft.Jet.OLEDB.4.0;Data" +
        " Source=C:\\tempdb.mdb;Jet OLEDB:Engine Type=5"};

            //invoke a CompactDatabase method of a JRO object
            //pass Parameters array
            objJRO.GetType().InvokeMember("CompactDatabase",
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                objJRO,
                oParams);

            //database is compacted now
            //to a new file C:\\tempdb.mdw
            //let's copy it over an old one and delete it

            System.IO.File.Delete(mdwfilename);
            System.IO.File.Move("C:\\tempdb.mdb", mdwfilename);

            //clean up (just in case)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objJRO);
            objJRO = null;
        }



    }
}
