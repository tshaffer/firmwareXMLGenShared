using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Net;
using System.Xml;

namespace FirmwareXmlGenShared
{

    public class FirmwareXmlGen
    {
        public class FirmwareFile
        {
            public string Family { get; set; }
            public string Type { get; set; }
            public string BAVersion { get; set; }
            public string Version { get; set; }
            public string VersionNumber { get; set; }
            public string Link { get; set; }
            public string Sha1 { get; set; }
            public string FileLength { get; set; }
            public bool ForceDownload { get; set; }
            public bool DontDownload { get; set; }

            public FirmwareFile()
            {
            }

            public static FirmwareFile ReadXml(XmlElement element)
            {
                string family = String.Empty;
                string type = String.Empty;
                string baVersion = String.Empty;   //the BAVersion property is not set for all FirmFile objects
                string version = String.Empty;
                string versionNumber = String.Empty;
                string link = String.Empty;
                string sha1 = String.Empty;
                string fileLength = String.Empty;
                bool forceDownload = false;
                bool dontDownload = true;

                XmlNodeList families = element.GetElementsByTagName("family");
                if (families.Count == 1)
                {
                    XmlElement familyXML = families[0] as XmlElement;
                    family = familyXML.InnerText;
                }

                XmlNodeList types = element.GetElementsByTagName("type");
                if (types.Count == 1)
                {
                    XmlElement typeXML = types[0] as XmlElement;
                    type = typeXML.InnerText;
                }

                XmlNodeList baVersions = element.GetElementsByTagName("BAVersion");
                if (baVersions.Count == 1)
                {
                    XmlElement baVersionXML = baVersions[0] as XmlElement;
                    baVersion = baVersionXML.InnerText;
                }
                else if (baVersions.Count == 0)
                {
                    baVersion = String.Empty;
                }

                XmlNodeList versions = element.GetElementsByTagName("version");
                if (versions.Count == 1)
                {
                    XmlElement versionXML = versions[0] as XmlElement;
                    version = versionXML.InnerText;
                }

                XmlNodeList versionNumbers = element.GetElementsByTagName("versionNumber");
                if (versionNumbers.Count == 1)
                {
                    XmlElement versionNumberXML = versionNumbers[0] as XmlElement;
                    versionNumber = versionNumberXML.InnerText;
                }

                XmlNodeList links = element.GetElementsByTagName("link");
                if (links.Count == 1)
                {
                    XmlElement linkXML = links[0] as XmlElement;
                    link = linkXML.InnerText.Trim();
                }

                XmlNodeList sha1s = element.GetElementsByTagName("sha1");
                if (sha1s.Count == 1)
                {
                    XmlElement sha1XML = sha1s[0] as XmlElement;
                    sha1 = sha1XML.InnerText;
                }

                XmlNodeList fileLengths = element.GetElementsByTagName("fileLength");
                if (fileLengths.Count == 1)
                {
                    XmlElement fileLengthXML = fileLengths[0] as XmlElement;
                    fileLength = fileLengthXML.InnerText;
                }

                FirmwareFile firmwareFile = new FirmwareFile
                {
                    Family = family,
                    Type = type,
                    BAVersion = baVersion,
                    Version = version,
                    VersionNumber = versionNumber,
                    Link = link,
                    Sha1 = sha1,
                    FileLength = fileLength,
                    ForceDownload = forceDownload,
                    DontDownload = dontDownload
                };

                return firmwareFile;
            }
        }

        private static XmlTextWriter _writer = null;
        private static string _amazonBaseURL = "http://bsnm.s3.amazonaws.com/public/";

        private static ObservableCollection<FirmwareFile> _firmwareFiles = new ObservableCollection<FirmwareFile>();
        private static ObservableCollection<FirmwareFile> _newFirmwareFiles = new ObservableCollection<FirmwareFile>();
        private static Dictionary<string, FirmwareFile> _minimumCompatibleDictionary = new Dictionary<string, FirmwareFile>();
        private static Dictionary<string, FirmwareFile> _fwFilesOriginals = new Dictionary<string, FirmwareFile>();

        public static ObservableCollection<FirmwareFile> GetFirmwareFiles()
        {
            return _firmwareFiles;
        }

        public static void ClearFirmwareFiles()
        {
            _firmwareFiles.Clear();
        }

        public static bool OpenExistingXML(string existingFilePath)
        {
            try
            {
                if (File.Exists(existingFilePath))
                {
                    StreamReader sr = new StreamReader(existingFilePath);
                    XmlDocument doc = new XmlDocument();
                    doc.Load(sr);
                    return  ParseExistingXML(doc);
                }
            }
            catch (Exception e)
            {
                //MessageBox.Show("Error: " + e.ToString());
            }

            return false;
        }

        private static bool ParseExistingXML(XmlDocument doc)
        {
            XmlElement docElement = doc.DocumentElement;

            if (docElement.Name != "BrightSignFirmware")
            {
                return false;
            }

            XmlNodeList firmwareFilesXML = doc.GetElementsByTagName("FirmwareFile");

            foreach (XmlElement firmwareFileXML in firmwareFilesXML)
            {
                FirmwareFile firmwareFile = FirmwareFile.ReadXml(firmwareFileXML);
                if (firmwareFile != null)
                {
                    _firmwareFiles.Add(firmwareFile);
                }
            }

            return true;
        }

        public static void Create(string inputFile, string outputFolder, string outputFile)
        {
            if (outputFolder == String.Empty)
            {
                //MessageBox.Show("An output folder was not specified, please specify one to continue.");
                return;
            }

            try
            {
                ClearDataOnCreate();

                GetOriginalXMLValues(inputFile);

                _writer = new XmlTextWriter(System.IO.Path.Combine(outputFile) + ".tmp", System.Text.Encoding.UTF8);
                _writer.Formatting = Formatting.Indented;
                _writer.WriteStartDocument();
                _writer.WriteStartElement("BrightSignFirmware");

                bool success = WriteExistingFWObjects(outputFolder);
                if (!success)
                {
                    //MessageBox.Show("XML Document not created");
                    _writer.Close();
                    File.Delete(System.IO.Path.Combine(outputFile) + ".tmp");
                    return;
                }
                WriteNewFWObjects(outputFolder);

                _writer.WriteEndElement();
                _writer.WriteEndDocument();
                _writer.Close();
                _writer = null;

                ClearData();

                if (File.Exists(inputFile))
                {
                    OpenExistingXML(inputFile);
                }
                MoveXMLFile(outputFile);

                //MessageBox.Show("XML Document successfully created");
            }
            catch (Exception ex)
            {
            }
            //catch (InvalidUpgradeType error)
            //{
            //    MessageBox.Show("An invalid upgrade type was experienced, only Production, Beta, or MinimumCompatible are allowed");
            //}
            //catch (InvalidFamilyName error)
            //{
            //    MessageBox.Show("An invalid family name was experienced, only Monaco, Panther, Cheetah, Pandora3, Puma, Tiger, Lynx, Bobcat, Apollo, or Bpollo are allowed");
            //}
            //catch (AlreadyExistentMinCompatibleFamilyBAVersion error)
            //{
            //    MessageBox.Show("An existing Family and BA Version combination already exists");
            //}
            //catch (Exception error)
            //{
            //    MessageBox.Show("A very unexpected error was encountered.");
            //}
        }

        private static void ClearDataOnCreate()
        {
            _newFirmwareFiles.Clear();
            _fwFilesOriginals.Clear();
            _minimumCompatibleDictionary.Clear();
        }

        private static void ClearData()
        {
            _newFirmwareFiles.Clear();
            _fwFilesOriginals.Clear();
            _firmwareFiles.Clear();
            _minimumCompatibleDictionary.Clear();
        }

        private static void GetOriginalXMLValues(string existingXMLFile)
        {
            try
            {
                if (File.Exists(existingXMLFile))
                {
                    StreamReader sr = new StreamReader(existingXMLFile);
                    XmlDocument doc = new XmlDocument();
                    doc.Load(sr);

                    XmlElement docElement = doc.DocumentElement;

                    if (docElement.Name != "BrightSignFirmware")
                    {
                        //throw new InvalidXMLDocTitle();
                    }

                    XmlNodeList firmwareFilesXML = doc.GetElementsByTagName("FirmwareFile");

                    foreach (XmlElement firmwareFileXML in firmwareFilesXML)
                    {
                        FirmwareFile fwFile = FirmwareFile.ReadXml(firmwareFileXML);
                        if (fwFile != null)
                        {
                            string currentKey = fwFile.Type + fwFile.BAVersion + fwFile.Sha1;
                            _fwFilesOriginals.Add(currentKey, fwFile);
                        }
                    }
                }
                return;
            }
            catch (Exception ex)
            {
            }
            //catch (InvalidXMLDocTitle e)
            //{
            //    //MessageBox.Show("The title of the existing XML document does not match BrightSignFirmware, please specify the correct file to continue");
            //}
            //catch (Exception e)
            //{
            //    //MessageBox.Show("Error: " + e.ToString());
            //}
        }

        private static void WriteOutputFile(string existingFilePath, string outputFileName, string outputFolder)
        {
            _writer = new XmlTextWriter(System.IO.Path.Combine(outputFileName) + ".tmp", System.Text.Encoding.UTF8);
            _writer.Formatting = Formatting.Indented;
            _writer.WriteStartDocument();
            _writer.WriteStartElement("BrightSignFirmware");

            bool success = WriteExistingFWObjects(outputFolder);
            if (!success)
            {
                //MessageBox.Show("XML Document not created");
                _writer.Close();
                File.Delete(System.IO.Path.Combine(outputFileName) + ".tmp");
                return;
            }
            WriteNewFWObjects(outputFolder);

            _writer.WriteEndElement();
            _writer.WriteEndDocument();
            _writer.Close();
            _writer = null;

            ClearData();

            if (File.Exists(existingFilePath))
            {
                //_existingXML = existingFilePath;
                OpenExistingXML(existingFilePath);
            }
            MoveXMLFile(outputFileName);

            //MessageBox.Show("XML Document successfully created");
        }

        private static void WriteNewFWObjects(string outputFolder)
        {
            foreach (FirmwareFile fwFile in _newFirmwareFiles)
            {
                string currentMinCompKey = fwFile.Family + fwFile.BAVersion;

                if (_minimumCompatibleDictionary != null && _minimumCompatibleDictionary.ContainsKey(currentMinCompKey))//key already exists
                {
                    //throw new AlreadyExistentMinCompatibleFamilyBAVersion();
                }
                else
                {
                    CheckParameters(fwFile);
                    SelectUgradeFWFileType(fwFile, outputFolder);
                    WriteXMLFragment(fwFile);
                    _minimumCompatibleDictionary.Add(currentMinCompKey, fwFile);
                }
            }
        }


        private static bool WriteExistingFWObjects(string outputFolder)
        {
            FirmwareFile thisOriginal;
            foreach (FirmwareFile fwFile in _firmwareFiles)
            {
				// TODO - BIG HACK TO PROCEED IN DEBUGGING
				fwFile.DontDownload = false;

                CheckParameters(fwFile);
                _fwFilesOriginals.TryGetValue(fwFile.Type + fwFile.BAVersion + fwFile.Sha1, out thisOriginal);

				// TODO - for debugging purposes
				bool fwEqual = Equals(fwFile, thisOriginal);
				if (!fwEqual) {
					Console.WriteLine ("Poop");
				}

                //will execute in 2 conditions:
                //1) force download is checked
                //2) fwFile is not the same as thisOriginal and don't download is not checked
                if ((fwFile.ForceDownload || !Equals(fwFile, thisOriginal)) && !fwFile.DontDownload)
                {
                    SelectUgradeFWFileType(fwFile, outputFolder);
                }

                // don't download is checked but the files are not the same, need new information before writing to xml
				if (!Equals (fwFile, thisOriginal) && fwFile.DontDownload) {
					bool success = UpdateInformationWithoutDownload (fwFile);
					if (!success)
						return false;
				} else {
					Console.WriteLine ("foo");
				}
                WriteXMLFragment(fwFile);

                if (fwFile.Type == "MinimumCompatible")
                {
                    string currentKey = fwFile.Family + fwFile.BAVersion;
                    _minimumCompatibleDictionary.Add(currentKey, fwFile);
                }
            }
            return true;
        }

		public static bool Equals(FirmwareFile fwFile1, FirmwareFile fwFile2)
		{
			if (fwFile1.Family == fwFile2.Family &&
				fwFile1.Type == fwFile2.Type &&
				fwFile1.Version == fwFile2.Version &&
				fwFile1.BAVersion == fwFile2.BAVersion &&
				fwFile1.FileLength == fwFile2.FileLength &&
				fwFile1.ForceDownload == fwFile2.ForceDownload &&
				fwFile1.Link == fwFile2.Link &&
				fwFile1.Sha1 == fwFile2.Sha1 &&
				fwFile1.VersionNumber == fwFile2.VersionNumber)
				return true;
			else
				return false;
		}

        private static void MoveXMLFile(string outputFileName)
        {
            try
            {
                string currentFile = outputFileName + ".tmp";
                string finalFile = outputFileName;
                File.Copy(currentFile, finalFile, true);
                File.Delete(currentFile);

            }
            catch (Exception e)
            {
                //MessageBox.Show("Error: " + e.ToString());
            }
        }

        private static bool CheckParameters(FirmwareFile fwFile)
        {
            if (!(fwFile.Type == "Production" ||
                fwFile.Type == "Beta" ||
                fwFile.Type == "MinimumCompatible"))
            {
                //throw new InvalidUpgradeType();
                return false;
            }
            if (!(fwFile.Family == "Monaco" ||
                fwFile.Family == "Panther" ||
                fwFile.Family == "Cheetah" ||
                fwFile.Family == "Tiger" ||
                fwFile.Family == "Bobcat" ||
                fwFile.Family == "Lynx" ||
                fwFile.Family == "Puma" ||
                fwFile.Family == "Apollo" ||
                fwFile.Family == "Bpollo" ||
                fwFile.Family == "Pandora3"))
            {
                //throw new InvalidFamilyName();
                return false;
            }
            return true;
        }

        private static void SelectUgradeFWFileType(FirmwareFile fwFile, string outputFolder)
        {
            switch (fwFile.Family)
            {
                //case "Monaco":
                //case "Pandora3":
                //case "Apollo":
                //case "Bpollo":
                //    EditROKFWFileInfo(fwFile);
                //    break;
                case "Panther":
                case "Cheetah":
                case "Puma":
                case "Tiger":
                case "Bobcat":
                case "Lynx":
                    EditBSFWFFWFileInfo(fwFile, outputFolder);
                    break;
            }
            return;
        }

        private static void WriteXMLFragment(FirmwareFile fwFile)
        {
            _writer.WriteStartElement("FirmwareFile");

            _writer.WriteElementString("family", fwFile.Family);
            _writer.WriteElementString("type", fwFile.Type);

            if (fwFile.BAVersion != String.Empty)
            {
                _writer.WriteElementString("BAVersion", fwFile.BAVersion);
            }

            _writer.WriteElementString("version", fwFile.Version);

            if (fwFile.VersionNumber != String.Empty)
            {
                _writer.WriteElementString("versionNumber", fwFile.VersionNumber);
            }

            _writer.WriteElementString("link", fwFile.Link);
            _writer.WriteElementString("sha1", fwFile.Sha1);
            _writer.WriteElementString("fileLength", fwFile.FileLength);

            _writer.WriteEndElement(); // FirmwareFile

            return;
        }

        private static bool UpdateInformationWithoutDownload(FirmwareFile fwFile)
        {
            string familyAndVersion = fwFile.Family + fwFile.Version;
            string originalFamilyAndVersion = "";
            Dictionary<string, FirmwareFile>.ValueCollection values = _fwFilesOriginals.Values;
            foreach (FirmwareFile originalFile in values)
            {
                originalFamilyAndVersion = originalFile.Family + originalFile.Version;
                if (familyAndVersion == originalFamilyAndVersion)
                {
                    fwFile.VersionNumber = originalFile.VersionNumber;
                    fwFile.Link = originalFile.Link;
                    fwFile.Sha1 = originalFile.Sha1;
                    fwFile.FileLength = originalFile.FileLength;
                    return true;
                }
            }
            //MessageBox.Show("Error: " + fwFile.Family + " " + fwFile.Version + " must be downloaded.");
            return false;
        }

        //private static void EditROKFWFileInfo(FirmwareFile fwFile, string outputFolder)
        //{

        //    string family = fwFile.Family;
        //    try
        //    {
        //        int i;
        //        string baseFirmwareAccessURL = "https://builds.brightsign.biz/brightsign-releases/";

        //        string[] splitBAVersion = fwFile.BAVersion.Split(new Char[] { '.', '.', '.' });
        //        if ((fwFile.BAVersion == String.Empty || splitBAVersion.Length != 2) && fwFile.Type == "MinimumCompatible") //4 is used because a BA version number follows X.X convention
        //        {
        //            //throw new InvalidBAVersionException();
        //        }

        //        string[] splitFWVersion = fwFile.Version.Split(new Char[] { '.', '.' });
        //        if (splitFWVersion.Length != 3) //3 is used because a FW version number follows X.X.X convention
        //        {
        //            //throw new InvalidFWVersionException();
        //        }

        //        string[] branchArray = new string[2];
        //        branchArray[0] = splitFWVersion[0];
        //        branchArray[1] = splitFWVersion[1];
        //        string branch = String.Join(".", branchArray);
        //        int[] firmwareNumbersSplit = new int[3];
        //        for (i = 0; i < 3; i++) { firmwareNumbersSplit[i] = Convert.ToInt32(splitFWVersion[i]); }
        //        int firmwareNumberInt = firmwareNumbersSplit[0] * 65536 + firmwareNumbersSplit[1] * 256 + firmwareNumbersSplit[2];
        //        fwFile.VersionNumber = firmwareNumberInt.ToString();

        //        family = String.Concat(family.ToLower(), "-");
        //        string expectedFileNameBase = String.Concat(family, fwFile.Version, "-update");

        //        string[] firmwarePathArray = new string[2];
        //        firmwarePathArray[0] = branch;
        //        firmwarePathArray[1] = String.Concat("brightsign-", fwFile.Version, "/");
        //        string uRLFirmwareAccessPath = String.Join("/", firmwarePathArray);

        //        string zipFileName = expectedFileNameBase + ".zip";
        //        string zipFilePath = System.IO.Path.Combine(outputFolder, zipFileName);
        //        string zipDownloadFilePath = System.IO.Path.Combine(outputFolder, zipFileName);

        //        string rokFileName = expectedFileNameBase + ".rok";
        //        string rokFilePath = System.IO.Path.Combine(outputFolder, "update.rok");
        //        fwFile.Link = _amazonBaseURL + rokFileName;

        //        string accessFirmwareURL = baseFirmwareAccessURL + uRLFirmwareAccessPath + zipFileName;

        //        Downloader(accessFirmwareURL, zipDownloadFilePath, outputFolder);
        //        ExtractZip(zipDownloadFilePath, outputFolder);

        //        string namedROKFilePath = System.IO.Path.Combine(outputFolder, rokFileName);
        //        string extractedTextFilePath = System.IO.Path.Combine(outputFolder, expectedFileNameBase + ".sha1.txt");

        //        File.Copy(rokFilePath, namedROKFilePath, true);
        //        File.Delete(rokFilePath);
        //        File.Delete(zipFilePath);
        //        File.Delete(extractedTextFilePath);

        //        fwFile.Sha1 = GetSHA1Hash(namedROKFilePath);
        //        FileInfo fi = new FileInfo(namedROKFilePath);
        //        int fileLength = (int)fi.Length;
        //        fwFile.FileLength = fileLength.ToString();

        //        return;
        //    }
        //    catch (Exception ex)
        //    {
        //    }
        //    //catch (FileNotFoundException error)
        //    //{
        //    //    MessageBox.Show("Error: " + error.ToString());
        //    //}
        //    //catch (DirectoryNotFoundException error)
        //    //{
        //    //    MessageBox.Show("Error: " + error.ToString());
        //    //}
        //    //catch (WebException error)
        //    //{
        //    //    MessageBox.Show("Error: " + error.ToString());
        //    //}
        //    //catch (InvalidBAVersionException error)
        //    //{
        //    //    MessageBox.Show("Error: " + error.ToString());
        //    //}
        //    //catch (InvalidFWVersionException error)
        //    //{
        //    //    MessageBox.Show("Error: " + error.ToString());
        //    //}
        //    //catch (Exception error)
        //    //{
        //    //    MessageBox.Show("Error: " + error.ToString());
        //    //}
        //}

        private static void EditBSFWFFWFileInfo(FirmwareFile fwFile, string outputFolder)
        {
            string family = fwFile.Family;
            try
            {
                int i;
                family = String.Concat(family.ToLower(), "-");
                string baseFirmwareAccessURL = "https://builds.brightsign.biz/brightsign-releases/";

                string[] splitBAVersion = fwFile.BAVersion.Split(new Char[] { '.', '.', '.' });
                if ((fwFile.BAVersion == String.Empty || splitBAVersion.Length != 2) && fwFile.Type == "MinimumCompatible") //4 is used because a BA version number follows X.X.X.X convention
                {
                    //throw new InvalidBAVersionException();
                }

                string[] splitFWVersion = fwFile.Version.Split(new Char[] { '.', '.' });
                //if (splitFWVersion.Length != 3) //3 is used because a FW version number follows X.X.X convention
                //{
                //    throw new InvalidFWVersionException();
                //}

                string[] branchArray = new string[2];
                branchArray[0] = splitFWVersion[0];
                branchArray[1] = splitFWVersion[1];
                string branch = String.Join(".", branchArray);
                int[] firmwareNumbersSplit = new int[3];
                for (i = 0; i < 3; i++) { firmwareNumbersSplit[i] = Convert.ToInt32(splitFWVersion[i]); }
                int firmwareNumberInt = firmwareNumbersSplit[0] * 65536 + firmwareNumbersSplit[1] * 256 + firmwareNumbersSplit[2];
                fwFile.VersionNumber = firmwareNumberInt.ToString();

                string expectedFileName = String.Concat(family, fwFile.Version, "-update.bsfw");

                string[] firmwarePathArray = new string[2];
                firmwarePathArray[0] = branch;
                firmwarePathArray[1] = String.Concat("brightsign-", fwFile.Version, "/");
                string uRLFirmwareAccessPathPartial = String.Concat(baseFirmwareAccessURL, branch, "/", fwFile.Version);
                string accessFirmwareURL = String.Concat(uRLFirmwareAccessPathPartial, "/", expectedFileName);

                fwFile.Link = _amazonBaseURL + expectedFileName;
                string bsfwDownloadFilePath = System.IO.Path.Combine(outputFolder, expectedFileName);

                Downloader(accessFirmwareURL, bsfwDownloadFilePath, outputFolder);

                fwFile.Sha1 = GetSHA1Hash(bsfwDownloadFilePath);
                FileInfo fi = new FileInfo(bsfwDownloadFilePath);
                fwFile.FileLength = fi.Length.ToString();

                return;
            }
            catch (Exception ex)
            {
            }
            //catch (WebException error)
            //{
            //    MessageBox.Show("Error: " + error.ToString());
            //}
            //catch (InvalidBAVersionException error)
            //{
            //    MessageBox.Show("Error: " + error.ToString());
            //}
            //catch (InvalidFWVersionException error)
            //{
            //    MessageBox.Show("Error: " + error.ToString());
            //}
            //catch (Exception error)
            //{
            //    MessageBox.Show("Error: " + error.ToString());
            //}
        }

        private static void Downloader(string firmwareURL, string filePath, string browsedDirectory)
        {
            string userName = "jshaffer";
            string password = "jowzos34";

            WebClient client = new WebClient();
            NetworkCredential credentials = new NetworkCredential(userName, password);

            client.Credentials = credentials;
            client.DownloadFile(firmwareURL, filePath);

            return;
        }

        public static string GetSHA1Hash(string pathName)
        {
            string strResult = "";
            string strHashData = "";

            byte[] arrbytHashValue;
            System.IO.FileStream oFileStream = null;

            System.Security.Cryptography.SHA1CryptoServiceProvider oSHA1Hasher =
                       new System.Security.Cryptography.SHA1CryptoServiceProvider();

            try
            {
                oFileStream = GetFileStream(pathName);
                arrbytHashValue = oSHA1Hasher.ComputeHash(oFileStream);
                oFileStream.Close();

                strHashData = System.BitConverter.ToString(arrbytHashValue);
                strHashData = strHashData.Replace("-", "");
                strResult = strHashData;
            }
            catch (Exception ex)
            {
                //Trace.WriteLine("Exception in FirmwareXMLGen in GetSHA1Hash(string pathName)");
                //Trace.WriteLine("Exception is: " + ex.ToString());
            }

            return (strResult.ToLower());
        }

        private static System.IO.FileStream GetFileStream(string pathName)
        {
            return (new System.IO.FileStream(pathName, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite));
        }


    }
}
