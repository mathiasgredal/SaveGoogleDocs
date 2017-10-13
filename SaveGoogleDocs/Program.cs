using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Google.Apis.Util.Store;
using Google.Apis.Services;
using SaveGoogleDocs;
using Newtonsoft.Json;
using System.Dynamic;

namespace SaveGoogleDocs
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            if (!GoogleDriveV2Library.IsInternet())
            {
                Console.WriteLine("Error no internet connection");
                return;
            }

            // Authenticate with google drive
            Credentials credentials;
            string pathToCredentials = "/Users/mathiasgredal/projects/SaveGoogleDocs/SaveGoogleDocs/Credentials.json";
            //string pathToCredentials = "/Users/mathiasgredal/projects/SaveGoogleDocs/SaveGoogleDocs/My Credentials.json";

            // Open the text file using a stream reader.
            using (System.IO.StreamReader sr = new System.IO.StreamReader(pathToCredentials))
            {
                // Read the stream to a string, and write the string to the console.
                string json = sr.ReadToEnd();
                JsonTextReader reader = new JsonTextReader(new System.IO.StringReader(json));
                credentials = JsonConvert.DeserializeObject<Credentials>(json);
                Console.WriteLine(credentials.ClientId);
                Console.WriteLine(credentials.ClientSecret);

            }

            var service = GoogleDriveV2Library.Authentication(credentials.ClientId, credentials.ClientSecret);

            var skoleFolder = GoogleDriveV2Library.GetFiles(service, "title=\"Skole\"")[0];

            GoogleDriveV2Library.CopyFolder(skoleFolder.Id, "Skole - Backup", service);

            //GoogleDriveV2Library.CopyFileIntoFolder("1-lK6Skh2PQfNSl3836qWZOZaL8CpF6EpcsHoIVIkVi0", "0B8NEg58TrL9RMm04MFhmX1lnT1k", true, serviceV2);
            //GoogleDriveV2Library.MoveFile("0B8NEg58TrL9ROUIxQW9WN2RlOWs", "0B8NEg58TrL9RMm04MFhmX1lnT1k", serviceV2);
            //Console.WriteLine("Total Children is: " + GoogleDriveV2Library.NumChildsInFolder("0B8NEg58TrL9RMm04MFhmX1lnT1k", serviceV2));
            Console.WriteLine("Succesful");
        }

    }
}

