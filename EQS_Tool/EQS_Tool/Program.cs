using System;
using System.IO;
using System.Data;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Content;
using System.Text;
using UglyToad.PdfPig.Geometry;
using System.Linq;
using UglyToad.PdfPig.Graphics.Operations.SpecialGraphicsState;
using Sylvan.Data;
using Sylvan.Data.Excel;
using System.Linq;
using System.Collections.Frozen;
using static UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor.ContentOrderTextExtractor;
using System.Text.RegularExpressions;
using System.Numerics;
using System.Configuration;
using System.Collections.Specialized;
using System.Xml;

namespace EQS_Tool
{ 
    class Program
    {
        static void Main(string[] args)
        {
            // get needed paths and settings from config file
            // likley to change over time so easy editing
            var appSettings = ConfigurationManager.AppSettings; 
            
            float scoreThreshold = float.Parse(appSettings["ScoreThreshold"]);

            bool autoMoveFiles = false;
            
            if(appSettings["AutoMovefiles"] == "true")
            {
                autoMoveFiles = true;
            }

            var tgpEHFilePath = appSettings["TGPEHFilePath"];
            var tgpFWTFilePath = appSettings["TGPFWTFilePath"];
            var tgpRRFilePath = appSettings["TGPRRFilePath"];
            var tgpUNSATFilePath = appSettings["TGPUNSATFilePath"];
            var gpEICRFilePath = appSettings["GPEICRFilePath"];
            var gpUNSATFilePath = appSettings["GPUNSATFilePath"];
            var gpEQSLogPath = appSettings["GPEQSLogPath"];

            // get base .exe directory - CWD once deployed????
            // these may need reviewing before deployment
            var pathCur = Directory.GetCurrentDirectory();
            var basePath = pathCur.Split(new string[] { "\\bin" }, StringSplitOptions.None)[0];
            var baseDir = basePath + "/" + "Certificates";          

            // get excel file or exit
            // needed for address/uprn matching
            // possible direct DB connection????
            string excelAddressListPath = basePath + "/PROPERTIES.xlsx";
            if (!File.Exists(excelAddressListPath))
            {
                Console.WriteLine("Missing dependency address file DB, exiting....");
                Environment.Exit(-1);
            }

            // load excel data to table object to process and pass between functions
            DataTable addressTable = new DataTable();
            using (var excelReader = ExcelDataReader.Create(excelAddressListPath))
            {
                addressTable.Load(excelReader);
            }

            // department sub folders, needed for logic
            string[] depts = { "EH", "FWT", "RR" };

            // needed directory structure
            // could be simplified with a loop
            string[] folders =
            {
                "Certificates",
                "Certificates\\EH",
                "Certificates\\EH\\ERROR",
                "Certificates\\EH\\UPRN_ERROR",
                "Certificates\\EH\\ADDRESS_CHECK_FAILED",
                "Certificates\\EH\\PROCESSED",
                "Certificates\\FWT",
                "Certificates\\FWT\\ERROR",
                "Certificates\\FWT\\UPRN_ERROR",
                "Certificates\\FWT\\ADDRESS_CHECK_FAILED",
                "Certificates\\FWT\\PROCESSED",
                "Certificates\\RR",                
                "Certificates\\RR\\ERROR",
                "Certificates\\RR\\UPRN_ERROR",   
                "Certificates\\RR\\ADDRESS_CHECK_FAILED",
                "Certificates\\RR\\PROCESSED"
            };


            // check and create directories if missing or exit
            foreach (string folder in folders)
            {
                try
                {
                    string folderPath = basePath + "/" + folder;

                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.WriteLine("Error creating directory structure, quitting.");
                    Environment.Exit(-1);
                }
            }

            // time date to save accu file
            string timestamp = MakeValidFileName(DateTime.Now.ToString());
            //accu file name
            string accuservFileName = $"accuserv_{timestamp}.txt";

            // process only pdf files in root of directory per dept sub folder
            foreach (string dept in depts)
            {
                string deptFolderPath = baseDir + "/" + dept;
                string[] files = Directory.GetFiles(deptFolderPath, "*.pdf", SearchOption.TopDirectoryOnly);

                foreach (string file in files)
                {
                    try
                    {
                        PdfDocument doc = PdfDocument.Open(file);
                        Page page = doc.GetPage(1);

                        // needed data from pdf files
                        string jobRef = "";
                        string uprn = "";
                        string certificateNumber = "";
                        string date = "";
                        string addressLine1 = "";
                        string addressLine2 = "";
                        string postcode = "";
                        string engineer = "";
                        string supervisor = "";

                        // only from eicr
                        string result = "";

                        // pdf data location rects
                        // populate per certificate type
                        // deeper testing to validate correct bounds
                        float[] jobRefRect = new float[4];
                        float[] uprnRect = new float[4];
                        float[] certificateNumberRect = new float[4];
                        float[] dateRect = new float[4];
                        float[] addressLine1Rect = new float[4];
                        float[] addressLine2Rect = new float[4];
                        float[] postcodeRect = new float[4];
                        float[] engineerRect = new float[4];
                        float[] supervisorRect = new float[4];
                        float[] resultRect = new float[4];

                        string certificateType = "";
                        var words = page.GetWords();
                        var pageText = string.Join(" ", words);

                        // get certificate type or move to error folder if invalid
                        // set bounding box coords per certificate type
                        if (pageText.Contains("ELECTRICAL INSTALLATION CONDITION REPORT"))
                        {
                            certificateType = "EICR";
                         
                            jobRefRect[0] = 407;
                            jobRefRect[1] = 450;
                            jobRefRect[2] = 460;
                            jobRefRect[3] = 458;
                            //Text: 5436545/1, Coordinates: (408, 451.4959989999999) - (445.8079999999998, 457.3202177499999)

                            uprnRect[0] = 582;
                            uprnRect[1] = 436;
                            uprnRect[2] = 650;
                            uprnRect[3] = 443;
                            //Text: B630540004, Coordinates: (582, 436.4959989999999) - (627.3679999999998, 442.2459989999999)

                            certificateNumberRect[0] = 610;
                            certificateNumberRect[1] = 546;
                            certificateNumberRect[2] = 647;
                            certificateNumberRect[3] = 553;
                            //Text: 28586529, Coordinates: (611, 546.495999) - (646.5839999999998, 552.245999)

                            dateRect[0] = 189;
                            dateRect[1] = 309;
                            dateRect[2] = 231;
                            dateRect[3] = 317;
                            //Text: 08/12/2023, Coordinates: (190, 310.495999) - (230.03200000000004, 316.32021775)

                            addressLine1Rect[0] = 581;
                            addressLine1Rect[1] = 422;
                            addressLine1Rect[2] = 780;
                            addressLine1Rect[3] = 430;
                            //add 1
                            //Text: 32, Coordinates: (582, 423.4959989999999) - (590.896, 429.2459989999999)
                            //: to, Coordinates: (593.12, 423.4959989999999) - (599.792, 429.0936552499999)
                            //: 79, Coordinates: (602.0160000000001, 423.4959989999999) - (610.912, 429.2459989999999)
                            //: Guinness, Coordinates: (613.1360000000001, 423.4959989999999) - (646.928, 429.3202177499999)
                            //: House,, Coordinates: (649.152, 423.4959989999999) - (674.496, 429.2225614999999)
                            //: Little, Coordinates: (676.72, 423.4959989999999) - (693.616, 429.2225614999999)
                            //: Hardings,, Coordinates: (695.84, 423.4959989999999) - (730.0719999999999, 429.2225614999999)

                            addressLine2Rect[0] = 551;
                            addressLine2Rect[1] = 408;
                            addressLine2Rect[2] = 780;
                            addressLine2Rect[3] = 416;
                            //add2
                            //Text: Welwyn, Coordinates: (552, 409.4959989999999) - (579.9999999999999, 415.2225614999999)
                            //Text: Garden, Coordinates: (582.2239999999999, 409.4959989999999) - (608.9039999999999, 415.3202177499999)
                            //Text: City,, Coordinates: (611.1279999999999, 409.4959989999999) - (627.1279999999999, 415.3202177499999)
                            //Text: Hertfordshire, Coordinates: (629.352, 409.4959989999999) - (675.5839999999998, 415.3202177499999)

                            postcodeRect[0] = 587;
                            postcodeRect[1] = 396;
                            postcodeRect[2] = 645;
                            postcodeRect[3] = 404;
                            //pc
                            //Text: AL7, Coordinates: (588, 397.4959989999999) - (602.232, 403.2225614999999)
                            //Text: 2EN, Coordinates: (604.456, 397.4959989999999) - (620.016, 403.2459989999999)
                            
                            engineerRect[0] = 218;
                            engineerRect[1] = 122;
                            engineerRect[2] = 279;
                            engineerRect[3] = 129;
                            //Text: BRUNO, Coordinates: (218, 122.49599899999998) - (246.88800000000003, 128.32412399999998)
                            //Text: VITELLI, Coordinates: (249.11200000000002, 122.49599899999998) - (278.01599999999996, 128.22256149999998)
                            
                            supervisorRect[0] = 218;
                            supervisorRect[1] = 46;
                            supervisorRect[2] = 303;
                            supervisorRect[3] = 53;
                            //Text: STEPHEN, Coordinates: (218, 46.49599899999987) - (255.78400000000008, 52.32021774999987)
                            //Text: PICKERING, Coordinates: (258.0080000000001, 46.49599899999987) - (302.01600000000013, 52.32021774999987)

                            //resultRect[0] = 661;
                            // resultRect[1] = 211;
                            //resultRect[2] = 750;
                            //resultRect[3] = 220;
                            resultRect[0] = 597;
                            resultRect[1] = 211;
                            resultRect[2] = 750;
                            resultRect[3] = 220;
                            //??????????Text: XXXXXXXXXXXXXX, Coordinates: (662, 212.49599899999998) - (736.7040000000002, 218.22256149999998)
                            //????Text: Satisfactory/Unsatisfactory**, Coordinates: (597.406, 211.0674) - (749.1729999999999, 219.2074)
                            //sat = Text: XXXXXXXXXXXXXX, Coordinates: (662, 212.49599899999998) - (736.7040000000002, 218.22256149999998)
                            //unsat = Text: XXXXXXXXXXX, Coordinates: (598, 212.49599899999998) - (656.6960000000001, 218.22256149999998)

                        }
                        else if (pageText.Contains("ELECTRICAL INSTALLATION CERTIFICATE"))
                        {
                            certificateType = "EIC";

                            jobRefRect[0] = 407;
                            jobRefRect[1] = 450;
                            jobRefRect[2] = 446;
                            jobRefRect[3] = 458;
                            //Text: 5666922/1, Coordinates: (408, 451.4959989999999) - (445.8079999999998, 457.3202177499999)

                            uprnRect[0] = 677;
                            uprnRect[1] = 438;
                            uprnRect[2] = 750;
                            uprnRect[3] = 445;
                            //Text: 220057, Coordinates: (678, 438.4959989999999) - (704.6879999999999, 444.2459989999999)

                            certificateNumberRect[0] = 610;
                            certificateNumberRect[1] = 546;
                            certificateNumberRect[2] = 647;
                            certificateNumberRect[3] = 553;
                            //Text: 29074008, Coordinates: (611, 546.495999) - (646.5839999999998, 552.245999)

                            dateRect[0] = 377;
                            dateRect[1] = 133;
                            dateRect[2] = 450;
                            dateRect[3] = 140;
                            //Text: 13/03/2024, Coordinates: (378, 133.49599899999998) - (418.0319999999998, 139.32021774999998)

                            addressLine1Rect[0] = 577;
                            addressLine1Rect[1] = 424;
                            addressLine1Rect[2] = 780;
                            addressLine1Rect[3] = 431;
                            //Text: 14, Coordinates: (578, 424.4959989999999) - (586.896, 430.2459989999999)
                            //Text: Co - Operative, Coordinates: (589.12, 424.4959989999999) - (636.6879999999999, 430.3241239999999)
                            //Text: Street,, Coordinates: (638.9119999999999, 424.4959989999999) - (662.48, 430.3202177499999)
                            //Text: Cudworth,, Coordinates: (664.7040000000001, 424.4959989999999) - (701.16, 430.3202177499999)
                            //Text: Barnsley,, Coordinates: (703.384, 424.4959989999999) - (736.728, 430.2225614999999)

                            addressLine2Rect[0] = 552;
                            addressLine2Rect[1] = 411;
                            addressLine2Rect[2] = 780;// 676;
                            addressLine2Rect[3] = 418;
                            //Text: South, Coordinates: (553, 411.4959989999999) - (573.904, 417.3202177499999)
                            //Text: Yorkshire, Coordinates: (576.128, 411.4959989999999) - (609.9119999999999, 417.2225614999999)

                            postcodeRect[0] = 587;
                            postcodeRect[1] = 396;
                            postcodeRect[2] = 645;
                            postcodeRect[3] = 404;
                            //Text: S72, Coordinates: (588, 397.4959989999999) - (602.232, 403.3202177499999)
                            //Text: 8DJ, Coordinates: (604.456, 397.4959989999999) - (618.68, 403.2459989999999)

                            engineerRect[0] = 96;
                            engineerRect[1] = 107;
                            engineerRect[2] = 180;
                            engineerRect[3] = 114;
                            //Text: GRAHAM, Coordinates: (97, 107.49599899999998) - (132.112, 113.32021774999998)
                            //Text: SHELDON, Coordinates: (134.33599999999998, 107.49599899999998) - (173.00800000000004, 113.32412399999998)

                            supervisorRect[0] = 96;
                            supervisorRect[1] = 43;
                            supervisorRect[2] = 180;
                            supervisorRect[3] = 50;
                            //Text: PAUL, Coordinates: (97, 43.49599899999987) - (117.89599999999999, 49.22256149999987)
                            //Text: WOODHOUSE, Coordinates: (120.11999999999999, 43.49599899999987) - (174.34400000000002, 49.32412399999987)
                        }
                        else if (pageText.Contains("MINOR ELECTRICAL INSTALLATION WORKS CERTIFICATE"))
                        {
                            certificateType = "MW";

                            jobRefRect[0] = 403;
                            jobRefRect[1] = 445;
                            jobRefRect[2] = 442;
                            jobRefRect[3] = 452;
                            //Text: 5144409/1, Coordinates: (404, 445.4959989999999) - (441.8079999999998, 451.3202177499999)

                            uprnRect[0] = 571;
                            uprnRect[1] = 431;
                            uprnRect[2] = 599;
                            uprnRect[3] = 438;
                            //Text: 204048, Coordinates: (572, 431.4959989999999) - (598.6879999999999, 437.2459989999999)

                            certificateNumberRect[0] = 610;
                            certificateNumberRect[1] = 545;
                            certificateNumberRect[2] = 647;
                            certificateNumberRect[3] = 552;
                            //Text: 28959731, Coordinates: (611, 545.495999) - (646.5839999999998, 551.245999)

                            dateRect[0] = 95;
                            dateRect[1] = 335;
                            dateRect[2] = 137;
                            dateRect[3] = 342;
                            //Text: 09/02/2024, Coordinates: (96, 335.4959989999999) - (136.03200000000004, 341.3202177499999)

                            addressLine1Rect[0] = 577;
                            addressLine1Rect[1] = 418;
                            addressLine1Rect[2] = 780;
                            addressLine1Rect[3] = 425;
                            //Text: 75, Coordinates: (578, 418.4959989999999) - (586.896, 424.1483427499999)
                            //Text: Castle, Coordinates: (589.12, 418.4959989999999) - (611.7919999999999, 424.3202177499999)
                            //Text: Walk,, Coordinates: (614.016, 418.4959989999999) - (634.016, 424.2225614999999)
                            //Text: Sheffield, Coordinates: (636.24, 418.4959989999999) - (667.3679999999999, 424.3202177499999)

                            addressLine2Rect[0] = 554;
                            addressLine2Rect[1] = 405;
                            addressLine2Rect[2] = 780;
                            addressLine2Rect[3] = 412;
                            //Text: Carr, Coordinates: (555, 405.4959989999999) - (570.5519999999999, 411.3202177499999)
                            //Text: Vale,, Coordinates: (572.776, 405.4959989999999) - (591.0079999999999, 411.2225614999999)
                            //Text: Bolsover,, Coordinates: (593.232, 405.4959989999999) - (626.5759999999999, 411.2225614999999)
                            //Text: Derbyshire, Coordinates: (628.8, 405.4959989999999) - (667.4719999999998, 411.2225614999999)

                            postcodeRect[0] = 583;
                            postcodeRect[1] = 391;
                            postcodeRect[2] = 620;
                            postcodeRect[3] = 398;
                            //Text: S2, Coordinates: (584, 391.4959989999999) - (593.784, 397.3202177499999)
                            //Text: 5JB, Coordinates: (596.008, 391.4959989999999) - (609.792, 397.2225614999999)

                            engineerRect[0] = 471;
                            engineerRect[1] = 137;
                            engineerRect[2] = 550;
                            engineerRect[3] = 144;
                            //Text: ROBIN, Coordinates: (472, 137.49599899999998) - (497.336, 143.32412399999998)
                            //Text: HOLLAND, Coordinates: (499.56, 137.49599899999998) - (537.3439999999999, 143.32412399999998)

                            supervisorRect[0] = 467;
                            supervisorRect[1] = 77;
                            supervisorRect[2] = 550;
                            supervisorRect[3] = 84;
                            //Text: JAMES, Coordinates: (468, 77.49599899999998) - (494.672, 83.32021774999998)
                            //Text: BARRETT, Coordinates: (496.896, 77.49599899999998) - (534.2320000000001, 83.22256149999998)
                        }
                        else if (pageText.Contains("DOMESTIC VISUAL CONDITION REPORT"))
                        {
                            certificateType = "VIS";

                        }
                        else if (pageText.Contains("INSTALLATION AND COMMISSIONING OF A FIRE DETECTION"))
                        {
                            certificateType = "DFHN";
                        }
                        else
                        {
                            File.Move(file, deptFolderPath + "/ERROR/" + Path.GetFileName(file));
                            continue;
                        }

                        // get text from pdf based on bounding rect
                        foreach (var word in words)
                        {
                            if (word.BoundingBox.Left >= jobRefRect[0] && word.BoundingBox.Bottom >= jobRefRect[1] &&
                               word.BoundingBox.Right <= jobRefRect[2] && word.BoundingBox.Top <= jobRefRect[3])
                            {
                                jobRef += word;
                            }
                            else if (word.BoundingBox.Left >= uprnRect[0] && word.BoundingBox.Bottom >= uprnRect[1] &&
                               word.BoundingBox.Right <= uprnRect[2] && word.BoundingBox.Top <= uprnRect[3])
                            {
                                uprn += word;
                            }
                            else if (word.BoundingBox.Left >= certificateNumberRect[0] && word.BoundingBox.Bottom >= certificateNumberRect[1] &&
                               word.BoundingBox.Right <= certificateNumberRect[2] && word.BoundingBox.Top <= certificateNumberRect[3])
                            {
                                certificateNumber += word;
                            }
                            else if (word.BoundingBox.Left >= dateRect[0] && word.BoundingBox.Bottom >= dateRect[1] &&
                               word.BoundingBox.Right <= dateRect[2] && word.BoundingBox.Top <= dateRect[3])
                            {
                                date += word;
                            }
                            else if (word.BoundingBox.Left >= addressLine1Rect[0] && word.BoundingBox.Bottom >= addressLine1Rect[1] &&
                               word.BoundingBox.Right <= addressLine1Rect[2] && word.BoundingBox.Top <= addressLine1Rect[3])
                            {
                                addressLine1 += word + " ";
                            }
                            else if (word.BoundingBox.Left >= addressLine2Rect[0] && word.BoundingBox.Bottom >= addressLine2Rect[1] &&
                               word.BoundingBox.Right <= addressLine2Rect[2] && word.BoundingBox.Top <= addressLine2Rect[3])
                            {
                                addressLine2 += word + " ";
                            }
                            else if (word.BoundingBox.Left >= postcodeRect[0] && word.BoundingBox.Bottom >= postcodeRect[1] &&
                               word.BoundingBox.Right <= postcodeRect[2] && word.BoundingBox.Top <= postcodeRect[3])
                            {
                                postcode += word + " ";
                            }
                            else if (word.BoundingBox.Left >= engineerRect[0] && word.BoundingBox.Bottom >= engineerRect[1] &&
                               word.BoundingBox.Right <= engineerRect[2] && word.BoundingBox.Top <= engineerRect[3])
                            {
                                engineer += word + " ";
                            }
                            else if (word.BoundingBox.Left >= supervisorRect[0] && word.BoundingBox.Bottom >= supervisorRect[1] &&
                               word.BoundingBox.Right <= supervisorRect[2] && word.BoundingBox.Top <= supervisorRect[3])
                            {
                                supervisor += word + " ";
                            }
                            else if (word.BoundingBox.Left >= resultRect[0] && word.BoundingBox.Bottom >= resultRect[1] &&
                               word.BoundingBox.Right <= resultRect[2] && word.BoundingBox.Top <= resultRect[3])
                            {
                                result += word;
                            }
                        }

                        // get address from db that matches uprn in certificate
                        // returns only the first item - may be problemtic for communals????
                        // if return UPRN_ERROR move to UPRN ERROR Folder
                        string addressDB = GetAddressFromDataTable(addressTable, uprn.ToUpper());

                        if (addressDB == "UPRN_ERROR") 
                        {
                            File.Move(file, deptFolderPath + $"/UPRN_ERROR/{Path.GetFileName(file)}");
                            continue;
                        }

                        // this is the amount of X's present on an unsatisfactory certificate
                        if (!result.Contains("XXXXXXXXXXXXXX"))
                        {
                            result = "UNSAT";
                        }
                        else
                        {
                            result = "SAT";
                        }

                        // clean and format date
                        // is DD/MM/YYYY
                        // need DDMMYY
                        //date = date.Replace("/", "");
                        string[] dateArray = date.Split("/");
                        date = dateArray[0] + dateArray[1] + dateArray[2].Substring(2);

                        //helper function to get text coords
                        //PrintCoordsToConsole(page);

                        // get a text match score based on address
                        // if meets scorethreshold process
                        // else move to naming error folder
                        float score = AddressFuzzyMatchScore(addressDB, addressLine1, addressLine2, postcode);
                        if (score > scoreThreshold)
                        {                            
                            string namingConvention = "";
                            string fileName = "";

                            uprn = uprn.ToUpper();

                            if (certificateType == "EICR")
                            {
                                if (uprn.Contains("B") || uprn.Contains("C"))
                                {
                                    namingConvention = "CEICR";
                                }
                                else
                                {
                                    namingConvention = "DEICR";
                                }

                                if (result == "UNSAT")
                                {
                                    fileName = $"{uprn}_{namingConvention}_{date}_UNSAT.pdf";
                                }
                                else
                                {
                                    fileName = $"{uprn}_{namingConvention}_{date}.pdf";
                                }

                                if (dept == "FWT")
                                {
                                    // append to log file
                                    File.AppendAllText(baseDir + "/" + accuservFileName,
                                        $"{jobRef} : {date} : {uprn} : {addressLine1} : {result} : {certificateNumber}" + Environment.NewLine);
                                }
                            }
                            else
                            {
                                fileName = $"{uprn}_{certificateType}_{date}.pdf";
                            }

                            if (autoMoveFiles)
                            {                                
                                // copy to tgp & gp locations
                                if (certificateType == "EICR")
                                {
                                    if (result == "UNSAT")
                                    {
                                        File.Copy(file, gpUNSATFilePath + $"/{fileName}");
                                        File.Copy(file, tgpUNSATFilePath + $"/{fileName}");
                                        File.Delete(file);
                                        continue;
                                    }
                                    else
                                    {
                                        File.Copy(file, gpEICRFilePath + $"/{fileName}");
                                    }
                                }

                                if (dept == "EH")
                                {
                                    File.Copy(file, tgpEHFilePath + $"/{fileName}");
                                    File.Delete(file);
                                }
                                else if (dept == "FWT")
                                {
                                    File.Copy(file, tgpFWTFilePath + $"/{fileName}");
                                    File.Delete(file);
                                }
                                else if (dept == "RR")
                                {
                                    File.Copy(file, tgpRRFilePath + $"/{fileName}");
                                    File.Delete(file);
                                }
                                else if (result == "UNSAT")
                                {
                                    File.Copy(file, tgpUNSATFilePath + $"/{fileName}");
                                    File.Delete(file);
                                }  
                            }
                            else
                            {
                                // move locally
                                File.Move(file, deptFolderPath + $"/PROCESSED/{fileName}");
                            }
                            
                        }
                        else
                        {
                            //move address error
                            File.Move(file, deptFolderPath + $"/ADDRESS_CHECK_FAILED/{Path.GetFileName(file)}");
                        }
                    }
                    catch (Exception ex)
                    {
                        //File.Move(file, deptFolderPath + "/ERROR/" + Path.GetFileName(file));
                        Console.WriteLine(ex.Message);
                        // show file with error
                        Console.WriteLine($"Error encounted with: {Path.GetFileName(file)}");
                        File.Move(file, deptFolderPath + $"/ERROR/{Path.GetFileName(file)}");
                    }
                }                
            }

            string GetAddressFromDataTable(DataTable addressTable, string uprn)
            {
                string columnName = "[Property Reference]";
                var address = "";

                // Filter rows based on the matching cell value
                try
                {
                    DataRow[] matchingRows = addressTable.Select($"{columnName} = '{uprn}'");
                    address = matchingRows[0]["Property Address"].ToString();
                }
                catch
                {
                    address = "UPRN_ERROR";
                }

                return address;
            }

            // helper function
            void PrintCoordsToConsole(Page page)
            {
                var words = page.GetWords();

                foreach (var word in words)
                {
                    Console.WriteLine($"Text: {word.Text}, Coordinates: ({word.BoundingBox.BottomLeft.X}, {word.BoundingBox.BottomLeft.Y}) - ({word.BoundingBox.TopRight.X}, {word.BoundingBox.TopRight.Y})");
                }
            }

            float AddressFuzzyMatchScore(string addressDB, string address1, string address2, string postcode)
            {
                // need to build score logic here
                String[] addressDBSplit = addressDB.Split(',');

                string num1 = new String(addressDBSplit.First<string>().Where(Char.IsDigit).ToArray());                
                string num2 = new String(address1.Where(Char.IsDigit).ToArray());                

                string addressDBString = string.Join("", addressDBSplit.Take(addressDBSplit.Count() - 1).ToArray()).ToLower();
                string addressCert = address1.Replace(",", "").ToLower() + address2.Replace(",", "").ToLower();
                //Console.WriteLine(addressCert);

                //Console.WriteLine($"Postcode DB: {addressDBSplit.Last<string>().Trim()}");
                //Console.WriteLine($"Postcode Cert: {postcode.Trim()}");
                var scorePC = LevenshteinDistance(addressDBSplit.Last<string>().Trim(), postcode.Trim());
                //Console.WriteLine($"Score Postcode: {scorePC}");

                //Console.WriteLine($"Number DB: {num1}");
                //Console.WriteLine($"Number Cert: {num2}");
                var scoreNum = LevenshteinDistance(num1, num2);
                //Console.WriteLine($"Score Number: {scoreNum}");
                
                //Console.WriteLine($"Address DB: {addressDBString}");
                //Console.WriteLine($"Address Cert: {addressCert}");
                var scoreADD = LevenshteinDistance(addressDBString, addressCert);
                //Console.WriteLine($"Address Score: {scoreADD}");    
                
                return 1.0f;
            }

            // from copilot
            int LevenshteinDistance(string s1, string s2)
            {
                int len1 = s1.Length;
                int len2 = s2.Length;

                // Initialize a 2D array to store distances
                int[,] dp = new int[len1 + 1, len2 + 1];

                // Initialize the first row and column
                for (int i = 0; i <= len1; i++)
                    dp[i, 0] = i;
                for (int j = 0; j <= len2; j++)
                    dp[0, j] = j;

                // Fill in the rest of the array
                for (int i = 1; i <= len1; i++)
                {
                    for (int j = 1; j <= len2; j++)
                    {
                        int cost = (s1[i - 1] == s2[j - 1]) ? 0 : 1;
                        dp[i, j] = Math.Min(Math.Min(dp[i - 1, j] + 1, dp[i, j - 1] + 1), dp[i - 1, j - 1] + cost);
                    }
                }

                // The final value in the array represents the Levenshtein distance
                return dp[len1, len2];
            }

            string MakeValidFileName(string name)
            {
                string invalidChars = Regex.Escape(new string(System.IO.Path.GetInvalidFileNameChars()));
                string invalidRegStr = string.Format(@"([{0}]*\.+$)|([{0}]+)", invalidChars);

                return Regex.Replace(name, invalidRegStr, "_");
            }

        }

    }
}
