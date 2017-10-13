using System;
using System.Collections.Generic;
using System.Linq;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v2.Data;
using System.Threading;
using Google.Apis.Util.Store;
using Google.Apis.Services;
using Google.Apis.Drive.v2;
using ShellProgress;


namespace SaveGoogleDocs
{
    public static class GoogleDriveV2Library
    {
        public static DriveService Authentication(
            string clientId = "1024267779669-p1a352ck3sq1v12opb2grgoo2dlb729d.apps.googleusercontent.com",
            string clientSecret = "xD_Y5yMijTP2Zutw6E_AqxD0",
            string FileDataStoreIdentifier = "MathiasGredal.GoogleDrive.Auth.Store",
            string ApplicationName = "Save Google Docs")
        {
            //Scopes for use with the Google Drive API
            string[] scopes = new string[] { DriveService.Scope.Drive, DriveService.Scope.DriveFile };

            UserCredential credential = GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets
            {
                ClientId = clientId,
                ClientSecret = clientSecret
            },
            scopes,
            Environment.UserName,
            CancellationToken.None,
            new FileDataStore(FileDataStoreIdentifier)).Result;

            DriveService service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }

        /// <summary>
        /// Query files in google drive
        /// </summary>
        /// <returns>The files.</returns>
        /// <param name="service">The Driveservice.</param>
        /// <param name="search">Your query for help see https://developers.google.com/drive/v3/web/search-parameters.</param>
        public static IList<File> GetFiles(DriveService service, string search)
        {

            IList<File> Files = new List<File>();

            try
            {
                //List all of the files and directories for the current user.  
                // Documentation: https://developers.google.com/drive/v2/reference/files/list
                FilesResource.ListRequest list = service.Files.List();
                list.MaxResults = 1000;
                if (search != null)
                {
                    list.Q = search;
                }
                FileList filesFeed = list.Execute();

                //// Loop through until we arrive at an empty page
                while (filesFeed.Items != null)
                {
                    // Adding each item  to the list.
                    foreach (File item in filesFeed.Items)
                    {
                        Files.Add(item);
                    }

                    // We will know we are on the last page when the next page token is
                    // null.
                    // If this is the case, break.
                    if (filesFeed.NextPageToken == null)
                    {
                        break;
                    }

                    // Prepare the next page of results
                    list.PageToken = filesFeed.NextPageToken;

                    // Execute and process the next page request
                    filesFeed = list.Execute();
                }
            }
            catch (Exception ex)
            {
                // In the event there is an error with the request.
                Console.WriteLine(ex.Message);
            }
            return Files;
        }

        /// <summary>
        /// Copies the file.(It doesnt write the file it just gives you a copy)
        /// </summary>
        /// <returns>The copied file.</returns>
        /// <param name="service">The Driveservice</param>
        /// <param name="originFileId">Id of the file you want to copy</param>
        /// <param name="copyTitle">The title of the copied file</param>
        public static File CopyFile(DriveService service, String originFileId, String copyTitle)
        {
            File copiedFile = new File();
            copiedFile.Title = copyTitle;
            try
            {
                return service.Files.Copy(copiedFile, originFileId).Execute();
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
            }
            return null;
        }

        public static void CopyFolder(string folderId, string newFolderName, DriveService service)
        {
            // Create new folder
            Console.WriteLine("Copying Folder");
            File SkoleBackup;
            if (GetFiles(service, "title='" + newFolderName + "'").Count > -1)
            {
                // Skole-Backup doesn't exist we need to create it
                SkoleBackup = CreateDirectory(service, newFolderName, "0AMNEg58TrL9RUk9PVA");
            }
            else
            {
                // Skole-Backup does exist we need to create it
                SkoleBackup = GetFiles(service, "title='" + newFolderName + "'")[0];
            }

            // Copy folder structure
            Console.WriteLine("Finding All Folders");
            string StartDirectory = folderId;
            string WorkingDirectory = SkoleBackup.Id;

            List<File> StartFolders;
            {
                // Find all folders
                List<File> nestedUnsearched = NestedFolders(service, folderId);
                List<File> nestedSearched = new List<File>();
                do
                {

                    nestedSearched.Add(nestedUnsearched[0]);
                    nestedUnsearched.AddRange(NestedFolders(service, nestedUnsearched[0].Id));
                    nestedUnsearched.RemoveAt(0);
                } while (nestedUnsearched.Count > 0);
                nestedSearched.Remove(service.Files.Get(SkoleBackup.Id).Execute());
                StartFolders = nestedSearched;
            }
            Dictionary<string, string> MovedFiles = new Dictionary<string, string>();
            MovedFiles.Add(folderId, SkoleBackup.Id);
            List<File> WorkingNested = new List<File>();
            {
                Console.WriteLine("Creating Folders");
                do
                {
                    List<File> ToBeRemoved = new List<File>();
                    for (var i = 0; i < StartFolders.Count; i++)
                    {
                        //Console.WriteLine("Folder Found");
                        if (MovedFiles.ContainsKey(StartFolders[i].Parents[0].Id))
                        {
                            var newFolder = CreateDirectory(service, StartFolders[i].Title, MovedFiles[StartFolders[i].Parents[0].Id]);
                            MovedFiles.Add(StartFolders[i].Id, newFolder.Id);
                            ToBeRemoved.Add(StartFolders[i]);
                        }
                        Console.Write("\r{0}", i + 1);
                        Console.Write("/" + StartFolders.Count);
                    }
                    foreach (var file in ToBeRemoved)
                    {
                        StartFolders.Remove(file);
                    }
                } while (StartFolders.Count > 0);
                //This is to make sure next print isn't fucked
                Console.WriteLine("");
            }

            // Write copy all content and convert docs to pdf
            {
                int maxTicks = 0;

                foreach (var folder in MovedFiles)
                {
                    maxTicks += NumChildsInFolder(folder.Key, service);
                }

                Console.WriteLine("Copying Files");
                ProgressBar progress = new ProgressBar(maxTicks);
                int ticks = 0;

                foreach (var folder in MovedFiles)
                {
                    foreach (var file in AllChildsInFolder(folder.Key, service))
                    {
                        var fileInfo = CopyFileIntoFolder(file, folder.Value, service);
                        ClearCurrentConsoleLine();
                        string message = "Current File: " + fileInfo.Title + "(" + fileInfo.Id + ")";
                        Console.CursorLeft = 0;
                        Console.Write(message);
                        progress.Update(ticks);
                        ticks++;
                    }
                }
                progress.Complete();
            }
        }

        /// <summary>
        /// Creates a new folder
        /// </summary>
        /// <returns>The new folder.</returns>
        /// <param name="_service">The Driveservice.</param>
        /// <param name="_title">Title.</param>
        /// <param name="_parent">Parent folder.</param>
        /// <param name="_description">Description.</param>
        public static File CreateDirectory(
            DriveService _service,
            string _title,
            string _parent = "0AMNEg58TrL9RUk9PVA",
            string _description = "")
        {

            File NewDirectory = null;

            // Create metaData for a new Directory
            File body = new File();
            body.Title = _title;
            body.MimeType = "application/vnd.google-apps.folder";
            body.Parents = new List<ParentReference>() { new ParentReference() { Id = _parent } };
            try
            {
                FilesResource.InsertRequest request = _service.Files.Insert(body);
                NewDirectory = request.Execute();
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
            }

            return NewDirectory;
        }

        /// <summary>
        /// Finds all the nested folders
        /// </summary>
        /// <returns>Returns a list of the nested folders.</returns>
        /// <param name="service">The Driveservice.</param>
        /// <param name="_parentID">The folder from which to search from</param>
        public static List<File> NestedFolders(DriveService service, string _parentID)
        {
            List<File> results = new List<File>();
            string pageToken = "";
            do
            {
            FINDNESTEDFOLDERS:
                try
                {
                    var query = "'" + _parentID + "'" + " in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false";

                    var request = service.Files.List();
                    request.Fields = "items(id,parents)";
                    request.Q = query;
                    request.PageToken = pageToken;
                    var result = request.Execute();
                    results.AddRange(result.Items);
                    pageToken = result.NextPageToken;
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("ERROR WHEN FINDING NESTED FOLDERS");
                    Console.ResetColor();
                    Console.WriteLine("ERROR MESSAGE: ");
                    Console.WriteLine(ex.Message);
                    Console.WriteLine("");
                    Console.WriteLine("Retrying...");
                    goto FINDNESTEDFOLDERS;
                }

            } while (pageToken != null);
            return results;

        }

        /// <summary>
        /// Copies the file into folder, and turns any google document into pdf
        /// </summary>
        /// <param name="fileId">File identifier.</param>
        /// <param name="folderId">Folder identifier.</param>
        /// <param name="service">Service.</param>
        public static File CopyFileIntoFolder(string fileId, string folderId, DriveService service)
        {
            // copy the file
            var file = service.Files.Get(fileId).Execute();
            // if it isn't a google type we don't need to convert it and we can just move it directly

            try
            {
                if (!IsGoogleType(file.MimeType))
                {
                    var copy = CopyFile(service, fileId, service.Files.Get(fileId).Execute().Title);
                    var setRequest = service.Files.Update(copy, copy.Id);
                    List<string> parentIds = new List<string>();
                    foreach (var parent in copy.Parents)
                    {
                        parentIds.Add(parent.Id);
                    }
                    var previousParents = String.Join(",", parentIds);
                    setRequest.RemoveParents = previousParents;
                    setRequest.AddParents = folderId;
                    setRequest.Execute();
                }
                else
                {
                    var pdf = service.Files.Export(file.Id, "application/pdf");
                    var stream = System.IO.File.Open(Environment.CurrentDirectory + "/temp.pdf", System.IO.FileMode.OpenOrCreate);
                    pdf.Download(stream);
                    stream.Close();
                    UploadFile(Environment.CurrentDirectory + "/temp.pdf", file.Title, "application/pdf", folderId, service);

                }
            }
            catch (Exception ex)
            {
                /*
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("ERROR WHEN COPYING FILE");
                Console.ResetColor();
                Console.WriteLine("ERROR MESSAGE: ");
                Console.WriteLine(ex.Message);
                Console.WriteLine("");
                Console.WriteLine("The error was with " + file.Title + "(" + file.Title + ")");
                Console.WriteLine("Retrying...");
                */

                CopyFileIntoFolder(fileId, folderId, service);
            }

            return file;

        }

        public static void UploadFile(string path, string title, string mimeType, string parentId, DriveService service)
        {
            // Upload file
            var fileMetadata = new File()
            {
                Title = title
            };

            FilesResource.InsertMediaUpload request;

            using (var stream = System.IO.File.Open(path, System.IO.FileMode.Open))
            {
                request = service.Files.Insert(fileMetadata, stream, mimeType);
                //request.Fields = "id";
                request.Upload();
            }
            var file = request.ResponseBody;

            MoveFile(file.Id, parentId, service);
        }

        public static void MoveFile(string fileId, string parentId, DriveService service)
        {
            // Retrieve the existing parents to remove
            var getRequest = service.Files.Get(fileId);
            //getRequest.Fields = "parents";
            var file = getRequest.Execute();
            List<string> parentIds = new List<string>();
            foreach (var parent in file.Parents)
            {
                parentIds.Add(parent.Id);
            }
            var previousParents = String.Join(",", parentIds);

            // Move the file to the new folder
            var updateRequest = service.Files.Update(new File(), fileId);
            updateRequest.AddParents = parentId;
            if (previousParents.Count() > 0)
                updateRequest.RemoveParents = previousParents;
            file = updateRequest.Execute();
        }

        public static bool IsGoogleType(string mimeType)
        {
            switch (mimeType)
            {
                case "application/vnd.google-apps.document":
                    return true;
                case "application/vnd.google-apps.drawing":
                    return true;
                case "application/vnd.google-apps.presentation":
                    return true;
                case "application/vnd.google-apps.script":
                    return true;
                case "application/vnd.google-apps.spreadsheet":
                    return true;
                default:
                    return false;
            }
        }

        public static bool IsInternet()
        {

            try
            {
                using (var client = new System.Net.WebClient())
                using (var stream = client.OpenRead("http://www.google.com"))
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        public static List<string> AllChildsInFolder(string folderId, DriveService service)
        {
            ChildrenResource.ListRequest request = service.Children.List(folderId);

            List<string> childIds = new List<string>();

            do
            {
                try
                {
                    ChildList children = request.Execute();

                    foreach (ChildReference child in children.Items)
                    {
                        var getRequest = service.Files.Get(child.Id);
                        getRequest.Fields = "mimeType , id, labels";
                        var childInfo = getRequest.Execute();
                        string mimetype = childInfo.MimeType;
                        bool isTrashed = (bool)childInfo.Labels.Trashed;
                        if (mimetype != "application/vnd.google-apps.folder" && !isTrashed)
                            childIds.Add(child.Id);
                    }
                    request.PageToken = children.NextPageToken;
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occurred: " + e.Message);
                    request.PageToken = null;
                }
            } while (!String.IsNullOrEmpty(request.PageToken));

            return childIds;

        }

        public static int NumChildsInFolder(string folderId, DriveService service)
        {

            ChildrenResource.ListRequest request = service.Children.List(folderId);
            request.Fields = "items(id)";
            request.Q = "mimeType != 'application/vnd.google-apps.folder' and trashed = false";

            int totalChildren = 0;

            do
            {
                try
                {
                    ChildList children = request.Execute();

                    totalChildren += children.Items.Count;

                    request.PageToken = children.NextPageToken;
                }
                catch (Exception e)
                {
                    Console.WriteLine("An error occurred: " + e.Message);
                    request.PageToken = null;
                }
            } while (!String.IsNullOrEmpty(request.PageToken));

            return totalChildren;
        }

        public static void ClearCurrentConsoleLine()
        {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, currentLineCursor);
        }
    }

    class Credentials
    {
        public string ClientId = "";
        public string ClientSecret = "";
    }
}
