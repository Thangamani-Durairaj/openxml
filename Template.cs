using System;
using System.Data;
using System.Web;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
 

public class Template 
{
    public static object SideHeading { get; private set; }

    public static void RMTemplate(DataSet ds, string strUserCode)
    {
        if (File.Exists(System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Resume/"), strUserCode + ".docx")))
        {
            File.Delete(System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Resume/"), strUserCode + ".docx"));
        }

        string sourceFile = System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Templates/"), "RM.docx");
        string destinationFile = System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Resume/"), strUserCode + ".docx");
        File.Copy(sourceFile, destinationFile, true);

        try
        {
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Resume/"), strUserCode + ".docx"), true))
            {
                HttpContext.Current.Trace.Write(strUserCode.ToString());
                
                Body body = new Body();

                Table table = new Table();

                myDoc.MainDocumentPart.StyleDefinitionsPart.Styles.Append(CreateStyle("H1", "FreightSans Light", "935242", 30, "Normal"));
                myDoc.MainDocumentPart.StyleDefinitionsPart.Styles.Append(CreateStyle("H2", "FreightSans", "935242", 25, "Normal"));

                myDoc.MainDocumentPart.StyleDefinitionsPart.Styles.Append(CreateStyle("H3", "FreightSans Light", "935242", 22, "Bold"));
                myDoc.MainDocumentPart.StyleDefinitionsPart.Styles.Append(CreateStyle("H4", "FreightSans Bold", "89896E", 15, "Bold"));
                myDoc.MainDocumentPart.StyleDefinitionsPart.Styles.Append(CreateStyle("H5", "FreightSans Bold", "Black", 15, "Bold"));

                myDoc.MainDocumentPart.StyleDefinitionsPart.Styles.Append(CreateStyle("P1", "FreightSans Book", "Black", 15, "Bold"));
                myDoc.MainDocumentPart.StyleDefinitionsPart.Styles.Append(CreateStyle("P2", "FreightSans Book", "Black", 15, "Normal"));
                myDoc.MainDocumentPart.StyleDefinitionsPart.Styles.Save();

                ImagePart imagePart = myDoc.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);
                using (System.IO.Stream stream = GetStreamFromUrl(System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Photos/"), strUserCode + ".jpg")))
                {
                    imagePart.FeedData(stream);
                }

                table.Append(RMAddRowC2(AddImageToBody(myDoc.MainDocumentPart.GetIdOfPart(imagePart)), ParaHeading(ds.Tables[0].Rows[0]["Name"].ToString(), "H1")));

                table.Append(RMAddRowC1(ParaHeading("Educational Qualifications", "H3")));

                Table tlQual = new Table();
                tlQual.Append(TabelStyle(BorderValues.Single));
                if (ds.Tables[3].Rows.Count >= 1)
                {
                    for (int i = 0; i <= ds.Tables[3].Rows.Count - 1; i++)
                    {
                        tlQual.Append(AddRow1(ds.Tables[3].Rows[i]["YearPass"].ToString(), ds.Tables[3].Rows[i]["Qualify"].ToString() +","+ ds.Tables[3].Rows[i]["BoardOrUniver"].ToString()));
                    }
                    table.Append(tlQual);
                }

                table.Append(RMAddRowC1(ParaHeading("Core Focus Areas", "H3")));
                table.Append(RMAddRowC1(ParaHeading("Thematic areas of Practice or Research", "H3")));
                table.Append(RMAddRowC1(ParaHeading("Research Projects at IIHS", "H3")));
                table.Append(RMAddRowC1(ParaHeading("Practice Projects at IIHS", "H3")));
                table.Append(RMAddRowC1(ParaHeading("Core Teaching Areas", "H3")));
                table.Append(RMAddRowC1(ParaHeading("Previous Teaching Experience", "H3")));

                Table tlTeach = new Table();
                tlTeach.Append(TabelStyle(BorderValues.Single));
                if (ds.Tables[21].Rows.Count >= 1)
                {
                    for (int i = 0; i <= ds.Tables[21].Rows.Count - 1; i++)
                    {
                        tlTeach.Append(AddRow1(ds.Tables[21].Rows[i]["Institute"].ToString(), ds.Tables[21].Rows[i]["Role"].ToString()));
                    }
                    table.Append(tlTeach);
                }

                table.Append(RMAddRowC1(ParaHeading("Administrative Responsibilities at IIHS", "H3")));
                table.Append(RMAddRowC1(ParaHeading("Employment History (last three outside IIHS)", "H3")));

                Table tlWork = new Table();
                tlWork.Append(TabelStyle(BorderValues.Single));
                if (ds.Tables[8].Rows.Count >= 1)
                {
                    for (int i = 0; i <= ds.Tables[8].Rows.Count - 1; i++)
                    {
                        tlWork.Append(AddRow1(ds.Tables[8].Rows[i]["From"].ToString() + "-" + ds.Tables[8].Rows[i]["To"].ToString(), ds.Tables[8].Rows[i]["Organization"].ToString()));
                    }
                    table.Append(tlWork);
                }

                table.Append(RMAddRowC1(ParaHeading("Publications", "H3")));
                Table tlPub = new Table();
                tlPub.Append(TabelStyle(BorderValues.Single));
                if (ds.Tables[19].Rows.Count >= 1)
                {
                    for (int i = 0; i <= ds.Tables[19].Rows.Count - 1; i++)
                    {
                        tlPub.Append(AddRow1(ds.Tables[19].Rows[i]["Year"].ToString(), ds.Tables[19].Rows[i]["Title1"].ToString()));
                    }
                    table.Append(tlPub);
                }

                table.Append(RMAddRowC1(ParaHeading("Conference Presentations / Lectures / Talks", "H3")));
                Table tlConf = new Table();
                tlConf.Append(TabelStyle(BorderValues.Single));
                if (ds.Tables[1].Rows.Count >= 1)
                {
                    for (int i = 0; i <= ds.Tables[1].Rows.Count - 1; i++)
                    {
                        tlConf.Append(AddRow1(ds.Tables[1].Rows[i]["Year"].ToString(), ds.Tables[1].Rows[i]["Description"].ToString()));
                    }
                    table.Append(tlConf);
                }

                table.Append(RMAddRowC1(ParaHeading("Fellowships, Awards, Memberships of Professional Associations", "H3")));
                Table tlFellow = new Table();
                tlFellow.Append(TabelStyle(BorderValues.Single));
                if (ds.Tables[2].Rows.Count >= 1)
                {
                    for (int i = 0; i <= ds.Tables[2].Rows.Count - 1; i++)
                    {
                        tlFellow.Append(AddRow1(ds.Tables[2].Rows[i]["Type"].ToString() , ds.Tables[2].Rows[i]["ADetails"].ToString()));
                    }
                    table.Append(tlFellow);
                }

                table.Append(RMAddRowC1(ParaHeading("Links", "H3")));
                Table tlLinks = new Table();
                tlLinks.Append(TabelStyle(BorderValues.Single));
                if (ds.Tables[12].Rows.Count == 1)
                {
                    tlLinks.Append(AddRow1(ds.Tables[12].Rows[0]["Linked"].ToString(), ds.Tables[12].Rows[0]["Google"].ToString()));
                    tlLinks.Append(AddRow1(ds.Tables[12].Rows[0]["Twitter"].ToString(), ds.Tables[12].Rows[0]["URL1"].ToString()));
                    tlLinks.Append(AddRow1(ds.Tables[12].Rows[0]["URL2"].ToString(), ds.Tables[12].Rows[0]["URL3"].ToString()));
                    table.Append(tlLinks);
                }

                body.Append(table);
                myDoc.MainDocumentPart.Document.Append(body);
                myDoc.MainDocumentPart.Document.Save();
                myDoc.Close();
            }
        }
        catch(Exception ex)
        {
           // HttpContext.Current.Response.Write(ex.ToString());
        }
    }

    public static void ARTemplate(DataSet ds, string strUserCode)
    {
        if (File.Exists(System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Resume/"), strUserCode + ".docx")))
        {
            File.Delete(System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Resume/"), strUserCode + ".docx"));
        }

        try
        {
            using (WordprocessingDocument myDoc = WordprocessingDocument.Create(System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Resume/"), strUserCode + ".docx"), WordprocessingDocumentType.Document))
            {
                HttpContext.Current.Trace.Write(strUserCode.ToString());

                MainDocumentPart mainPart = myDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = new Body();

                StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylePart.Styles = new Styles();

                //ParagraphProperties UserHeadingParagPro = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties();
                //Justification CenterHeading = new Justification { Val = JustificationValues.Center };
                
                //    UserHeadingParagPro.Append(CenterHeading);
                //MainDocumentPart.Append(UserHeadingParagPro);

                stylePart.Styles.Append(CreateStyle("H1", "FreightSans Light", "406797", 30, "Normal"));
                stylePart.Styles.Append(CreateStyle("H2", "FreightSans", "4F81BD", 25, "Normal"));
                stylePart.Styles.Append(CreateStyle("P1", "FreightSans", "black", 20, "Normal"));

                body.Append(ParaHeading(ds.Tables[0].Rows[0]["Name"].ToString(), "H1"));

                Table t0 = new Table();

                //Table t1 = new Table();

                Table table = new Table();
                table.Append(TabelStyle(BorderValues.BasicThinLines));

               
                

                table.Append(ARAddRC0(ParaHeading("Personal Information:","H2")));

                if (ds.Tables[0].Rows.Count == 1)
                {
                    table.Append(ARAddRow6("Title",              ":", ds.Tables[0].Rows[0]["Name"].ToString()));
                    table.Append(ARAddRow6("Last Name",          ":", ds.Tables[0].Rows[0]["LastName"].ToString()));
                    table.Append(ARAddRow6("First Name",         ":", ds.Tables[0].Rows[0]["FirstName"].ToString()));
                    table.Append(ARAddRow6("Middle Name",        ":", ds.Tables[0].Rows[0]["MiddleName"].ToString()));
                    table.Append(ARAddRow6("Date of Birth",      ":", ds.Tables[0].Rows[0]["DOB"].ToString()));
                    table.Append(ARAddRow6("Country of Birth",   ":", ds.Tables[0].Rows[0]["Country"].ToString()));
                    table.Append(ARAddRow6("Citizenship",        ":", ds.Tables[0].Rows[0]["Domicile"].ToString()));
                }

                if (ds.Tables[11].Rows.Count == 1)
                {
                    table.Append(ARAddRow6("Pan Card", ":", ds.Tables[11].Rows[0]["Pan"].ToString()));
                    table.Append(ARAddRow6("Passport No.",":", ds.Tables[11].Rows[0]["PassportNo"].ToString()));
                }
                table.Append(ARAddRC1(ParaHeading("Contact Details:", "H2")));
                if (ds.Tables[4].Rows.Count == 1)
                {
                    //table.Append(AddRow3(ParaHeading("Contact Details:", "H2")));
                    table.Append(ARAddRow6("Country", ":", ds.Tables[4].Rows[0]["CountryCATitle"].ToString()));
                    table.Append(ARAddRow6("Street Address", ":", ds.Tables[4].Rows[0]["AddressCA1"].ToString()));
                    table.Append(ARAddRow6("City / Town / Locality", ":", ds.Tables[4].Rows[0]["CityCA"].ToString()));
                    table.Append(ARAddRow6("State", ":", ds.Tables[4].Rows[0]["StateCATitle"].ToString()));
                    table.Append(ARAddRow6("Postal Code", ":", ds.Tables[4].Rows[0]["PincodeCA"].ToString()));
                    table.Append(ARAddRow6("Phone", ":", ds.Tables[4].Rows[0]["MobileCA"].ToString()));
                }


                table.Append(ARAddRC2(ParaHeading("User Account Details:", "H2")));
                if (ds.Tables[4].Rows.Count == 1)
                {
                    table.Append(ARAddRow6("E - mail address", ":", ds.Tables[4].Rows[0]["Email"].ToString()));
                    table.Append(ARAddRow6("Alternative e - mail id", ":", ""));
                }
                
                //t0.Append(table);
                //Table tlQual = new Table();
                //tlQual.Append(TabelStyle(BorderValues.Single));
                //table.Append(ARAddRC3(ParaHeading("Education:", "H2")));
                //if (ds.Tables[3].Rows.Count == 1)
                //{
                   
                //    for (int i = 0; i <= ds.Tables[3].Rows.Count - 1; i++)
                //    {
                      
                //        tlQual.Append(ARAddRC4("Degree", "Period", "Status", "Institution", "Country"));
                //        tlQual.Append(ARAddRC4(ds.Tables[3].Rows[i]["Qualify"].ToString(), ds.Tables[3].Rows[i]["YearPass"].ToString() , ds.Tables[3].Rows[i]["Status"].ToString(), ds.Tables[3].Rows[i]["BoardOrUniver"].ToString(), ds.Tables[3].Rows[i]["Country"].ToString()));
                       
                //    }
                //    t0.Append(tlQual);
                //}


                
                t0.Append(table);
                Table tlQual = new Table();
                tlQual.Append(TabelStyle(BorderValues.Single));
                table.Append(ARAddRC3(ParaHeading("Education:", "H2")));
                if (ds.Tables[3].Rows.Count >= 1)
                {
                    tlQual.Append(ARAddRC4("Degree", "Period", "Status", "Institution", "Country"));
                    for (int i = 0; i <= ds.Tables[3].Rows.Count - 1; i++)
                    {
                        tlQual.Append(ARAddRC4(ds.Tables[3].Rows[i]["Qualify"].ToString(), ds.Tables[3].Rows[i]["YearPass"].ToString(), ds.Tables[3].Rows[i]["Status"].ToString(), ds.Tables[3].Rows[i]["BoardOrUniver"].ToString(), ds.Tables[3].Rows[i]["Country"].ToString()));
                    }
                    t0.Append(tlQual);
                }



                //table.Append(AddRC6(ParaHeading("Training Program attended (if any):", "H2")));
                //if (ds.Tables[6].Rows.Count == 1)
                //{

                //}



                //if (ds.Tables[3].Rows.Count >= 1)
                //{
                //    table.Append(AddRC6(ParaHeading("6 . EDUCATION:", "H2")));
                //    //table.Append(AddRc6("","Education Level", "Institution", "Year Of Passing"));
                //    table.Append(AddRC5("Degree", "Period", "Status", "Institution", "Country"));
                //    //("Degree", "Period", "Status", "Institution", "Country"));    tlQual.Append(AddRC5(ds.Tables[3].Rows[i]["Qualify"].ToString(), ds.Tables[3].Rows[i]["YearPass"].ToString() , ds.Tables[3].Rows[i]["Status"].ToString(), ds.Tables[3].Rows[i]["BoardOrUniver"].ToString(), ds.Tables[3].Rows[i]["Country"].ToString()));
                //    for (int i = 0; i <= ds.Tables[3].Rows.Count - 1; i++)
                //    {
                //        table.Append(AddRC5(ds.Tables[3].Rows[i]["Qualify"].ToString(), ds.Tables[3].Rows[i]["YearPass"].ToString(), ds.Tables[3].Rows[i]["Status"].ToString(), ds.Tables[3].Rows[i]["BoardOrUniver"].ToString(), ds.Tables[3].Rows[i]["Country"].ToString()));
                //    }
                //}


                //table.Append(AddRC6(ParaHeading("Language and Proficiency:", "H2")));    
                //table.Append(AddRC6(ParaHeading("Work Experience:", "H2")));
                //table.Append(AddRC6(ParaHeading("Expertise (provide atleast one expertise):", "H2")));


                //if (ds.Tables[3].Rows.Count >= 1)
                //{
                //    table.Append(AddRow4("Education Level", "Institution", "Year Of Passing"));
                //    for (int i = 0; i <= ds.Tables[3].Rows.Count - 1; i++)
                //    {
                //        table.Append(AddRow4(ds.Tables[3].Rows[i]["Level"].ToString(), ds.Tables[3].Rows[i]["BoardOrUniver"].ToString(), ds.Tables[3].Rows[i]["YearPass"].ToString()));
                //    }
                //}



                //table.Append(AddRC6(ParaHeading("Details of Key Tasks undertaken for Projects (if applicable):", "H2")));

                //table.Append(AddRC6(ParaHeading("Publications (if applicable):", "H2")));

                //if (ds.Tables[20].Rows.Count == 1)
                //{
                //    table.Append(AddRow2("Document Title", "Publisher", "Total Or Page Range", "Published Date"));
                //    for (int i = 0; i <= ds.Tables[20].Rows.Count - 1; i++)
                //    {
                //        table.Append(AddRow2(ds.Tables[20].Rows[i]["Title"].ToString(), ds.Tables[20].Rows[i]["Publisher"].ToString(), ds.Tables[20].Rows[i]["Pages"].ToString(), ds.Tables[20].Rows[i]["Year Of Publication"].ToString()));
                //    }
                //}

                //table.Append(AddRC6(ParaHeading("Publications (if applicable)", "H2")));
                //table.Append(AddRow2("Document Title", "Publisher", "Total Or Page Range", "Published Date"));
                //if (ds.Tables[20].Rows.Count == 1)
                //{

                //    for (int i = 0; i <= ds.Tables[20].Rows.Count - 1; i++)
                //    {
                //        table.Append(AddRow2(ds.Tables[20].Rows[i]["Document Title"].ToString(), ds.Tables[20].Rows[i]["Publisher"].ToString(), ds.Tables[20].Rows[i]["Total Or Page Range"].ToString(), ds.Tables[20].Rows[i]["Published Date"].ToString()));
                //    }
                //}
                //table.Append(AddRC6(ParaHeading("Research Interests (if any):", "H2")));
                //table.Append(AddRC6(ParaHeading("Master Planning and Development Control Regulations, Participatory Planning, Role of Consultants in Planning, Impact of Large Scale Development Programmes like JNNURM, NREGA, etc.", "P1")));

                body.Append(t0);
               // body.Append(t1);
           
                mainPart.Document.Append(body);
                    mainPart.Document.Save();
                myDoc.Close();
            }
        }
        catch(Exception ex)
        {
           // HttpContext.Current.Response.Write(ex.ToString());
        }
    }

    public static void ADBTemplate(DataSet ds, string strUserCode)
    {
        if (File.Exists(System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Resume/"), strUserCode + ".docx")))
        {
            File.Delete(System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Resume/"), strUserCode + ".docx"));
        }

        try
        {
            using (WordprocessingDocument myDoc = WordprocessingDocument.Create(System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Resume/"), strUserCode + ".docx"), WordprocessingDocumentType.Document))
            {
                HttpContext.Current.Trace.Write(strUserCode.ToString());
                
                MainDocumentPart mainPart = myDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = new Body();

                StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylePart.Styles = new Styles();

                stylePart.Styles.Append(CreateStyle("H1", "Cambria ", "406797", 30, "Normal"));
                stylePart.Styles.Append(CreateStyle("H2", "Arial Narrow", "4F81BD", 20, "Normal"));
                stylePart.Styles.Append(CreateStyle("P1", "FreightSans", "black", 20, "Normal"));

                body.Append(ParaHeading(ds.Tables[0].Rows[0]["Name"].ToString(), "H1"));



                Table table = new Table();
                table.Append(TabelStyle(BorderValues.BasicThinLines));
                //   table.Append(RMAddRowC1(ParaHeading("Core Focus Areas", "H3")));
                table.Append(ADBAddRC6(ParaHeading("1.PROPOSED POSITION FOR THIS PROJECT:","H2".ToString())));
                table.Append(ADBAddRow1("2.NAME:", ds.Tables[0].Rows[0]["Name"].ToString()));
                //table.Append(AddRow4("2  NAME:", ds.Tables[0].Rows[0]["Name"].ToString()));
                table.Append(ADBAddRow1("3.DATE OF BIRTH:",  ds.Tables[0].Rows[0]["DOB"].ToString()));
                table.Append(ADBAddRow1("4.NATIONALITY:",  ds.Tables[0].Rows[0]["NationalityTitle"].ToString()));

                table.Append(ADBAddRow1("5.PERSONAL ADDRESS: \n TELEPHONE NO:\nFAX NO:\nEMAIL ADDRESS:","".ToString()));

           

                if (ds.Tables[3].Rows.Count >= 1)
                {
                    table.Append(ADBAddRC6(ParaHeading("6 . EDUCATION:", "H2")));
                    //table.Append(AddRc6("","Education Level", "Institution", "Year Of Passing"));
                    table.Append(ADBAddRC5("Degree", "Period", "Status", "Institution", "Country"));
                    //("Degree", "Period", "Status", "Institution", "Country"));    tlQual.Append(AddRC5(ds.Tables[3].Rows[i]["Qualify"].ToString(), ds.Tables[3].Rows[i]["YearPass"].ToString() , ds.Tables[3].Rows[i]["Status"].ToString(), ds.Tables[3].Rows[i]["BoardOrUniver"].ToString(), ds.Tables[3].Rows[i]["Country"].ToString()));
                    for (int i = 0; i <= ds.Tables[3].Rows.Count -1; i++)
                    {
                        table.Append(ADBAddRC5(ds.Tables[3].Rows[i]["Qualify"].ToString(),ds.Tables[3].Rows[i]["YearPass"].ToString(), ds.Tables[3].Rows[i]["Status"].ToString(), ds.Tables[3].Rows[i]["BoardOrUniver"].ToString(), ds.Tables[3].Rows[i]["Country"].ToString()));
                    }
                }

               

                if (ds.Tables[21].Rows.Count >= 1)
                {
                    table.Append(ADBAddRC6(ParaHeading("7. OTHER TRAINING", "H2")));

                    table.Append(ADBAddRow6("Level" ,"Role","Department"));
                    for (int i = 0; i <= ds.Tables[21].Rows.Count - 1; i++)
                    {
                        table.Append(ADBAddRow6(ds.Tables[21].Rows[i]["Level"].ToString(), ds.Tables[21].Rows[i]["Role"].ToString(), ds.Tables[21].Rows[i]["Department"].ToString()));
                    }
                }

                table.Append(ADBAddRC6(ParaHeading("8.  LANGUAGE AND DEGREE OF EFFICIENCY:", "H2")));

                table.Append(ADBAddRC6(ParaHeading("9 .MEMBERSHIP IN PROFESSIONAL SOCIETIES:", "H2")));
                table.Append(ADBAddRC6(ParaHeading("10 .COUNTRIES OF WORK EXPERIENCE:", "H2")));

               

                if (ds.Tables[8].Rows.Count >= 1)
                {
                    table.Append(ADBAddRC6(ParaHeading("11 .EMPLOYMENT RECORD", "H2")));


                    table.Append(ADBAddRow6("Organization"                          ,"From"                     ,"To"));
                    for (int i = 0; i <= ds.Tables[8].Rows.Count - 1; i++)
                    {
                        table.Append(ADBAddRow6(ds.Tables[8].Rows[i]["Organization"].ToString(), ds.Tables[8].Rows[i]["From"].ToString(), ds.Tables[8].Rows[i]["To"].ToString()));
                    }
                }

                table.Append(ADBAddRC6(ParaHeading("12.DETAILED TASKS ASSIGNED:            WORK UNDERTAKEN THAT BEST ILLUSTRATES CAPABILITY TO HANDLE THE TASKS ASSIGNED:","H2")));

                //table.Append(AddRC6(ParaHeading("", "" )));
            
                table.Append(ADBAddRC6(ParaHeading("13 .PERSONAL / PERMANENT EMPLOYMENT STATUS CERTIFICATION:","H2")));

               
                table.Append(ADBAddRC6(ParaHeading( "I am a former ADB Staff member", "H2")));
                table.Append(ADBAddRC6(ParaHeading( "If yes, I retired from the ADB more than twelve (12) months ago", "H2")));
                table.Append(ADBAddRC6(ParaHeading( "I am a close relative of a current ADB staff member", "H2")));
                table.Append(ADBAddRC6(ParaHeading( "I am the spouse of a current ADB staff member", "H2"))); 
                 table.Append(ADBAddRC6(ParaHeading( "I am a regular full-time employee of the Consultant or associated firm", "H2")));
                table.Append(ADBAddRC6(ParaHeading( "I was not involved in the preparation of the terms of reference for this Consulting Services Agreement", "H2")));


                

               body.Append(table);
                mainPart.Document.Append(body);
                mainPart.Document.Save();
            }
        }
        catch(Exception ex)
        {
           // HttpContext.Current.Response.Write(ex.ToString());
        }

    }

    public static void PracticeTemplate(DataSet ds, string strUserCode)
    {
        if (File.Exists(System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Resume/"), strUserCode + ".docx")))
        {
            File.Delete(System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Resume/"), strUserCode + ".docx"));
        }

        try
        {
            using (WordprocessingDocument myDoc = WordprocessingDocument.Create(System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Resume/"), strUserCode + ".docx"), WordprocessingDocumentType.Document))
            {
                HttpContext.Current.Trace.Write(strUserCode.ToString());

                MainDocumentPart mainPart = myDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = new Body();

                StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylePart.Styles = new Styles();
                stylePart.Styles.Append(CreateStyle("P1", "FreightSans", "black", 20, "Normal"));

                Table table = new Table();
                table.Append(TabelStyle(BorderValues.Dotted));

                table.Append(PracticeAddRow1("Proposed position for the project", ""));
                table.Append(PracticeAddRow1("Name of Personnel", ds.Tables[0].Rows[0]["Name"].ToString()));
                table.Append(PracticeAddRow1("Countries of work experience", ""));
                table.Append(PracticeAddRow1("Specialization(Brief - 1 para)", ""));
                table.Append(PracticeAddRow1("Professional Brief(2 para)", ""));

                table.Append(PracticeAddRow2("Educational Qualifications ", "Institution", "Year", "Course"));

                if (ds.Tables[3].Rows.Count >= 1)
                {
                    for (int i = 0; i <= ds.Tables[3].Rows.Count - 1; i++)
                    {
                        table.Append(PracticeAddRow2(ds.Tables[3].Rows[i]["Level"].ToString(), ds.Tables[3].Rows[i]["BoardOrUniver"].ToString(), ds.Tables[3].Rows[i]["YearPass"].ToString(), ds.Tables[3].Rows[i]["Qualify"].ToString()));
                    }
                }

                table.Append(PracticeAddRow2("Employment Record till date", "Employer", "From", "To"));
                if (ds.Tables[8].Rows.Count >= 1)
                {
                    for (int i = 0; i <= ds.Tables[8].Rows.Count - 1; i++)
                    {
                        table.Append(PracticeAddRow2("", ds.Tables[8].Rows[i]["Organization"].ToString(), ds.Tables[8].Rows[i]["From"].ToString(), ds.Tables[8].Rows[i]["To"].ToString()));
                    }
                }

                table.Append(PracticeAddRow1("Key projects and duties", ""));
                table.Append(PracticeAddRow1("Training Experience", ""));

                table.Append(PracticeAddRow1("Teaching Experience", ""));
                if (ds.Tables[21].Rows.Count >= 1)
                {
                    for (int i = 0; i <= ds.Tables[21].Rows.Count - 1; i++)
                    {
                        table.Append(PracticeAddRow2(ds.Tables[21].Rows[i]["Level"].ToString(), ds.Tables[21].Rows[i]["Institute"].ToString(), ds.Tables[21].Rows[i]["Department"].ToString(), ds.Tables[21].Rows[i]["Role"].ToString()));
                    }
                }

                table.Append(PracticeAddRow1("Research Papers &Reports", ""));
                table.Append(PracticeAddRow1("Awards and Recognition", ""));

                table.Append(PracticeAddRow1("Computer Skills", ""));
                table.Append(PracticeAddRow1("Date of Birth", ds.Tables[0].Rows[0]["DOB"].ToString()));
                table.Append(PracticeAddRow1("Nationality", ds.Tables[0].Rows[0]["NationalityTitle"].ToString()));
                table.Append(PracticeAddRow1("Personal address", ""));

                table.Append(PracticeAddRow1("Contact details",""));

                if (ds.Tables[4].Rows.Count == 1)
                {
                    for (int i = 0; i <= ds.Tables[4].Rows.Count - 1; i++)
                    { 
                            table.Append(PracticeAddRow3("Mobile", ":", ds.Tables[4].Rows[0]["MobileCA"].ToString()));
                        table.Append(PracticeAddRow3("Home", ":", ds.Tables[4].Rows[0]["HomePA"].ToString()));
                    }
                }



                if (ds.Tables[4].Rows.Count == 1)
                table.Append(PracticeAddRow1("Email", ds.Tables[4].Rows[0]["Email"].ToString()));

                table.Append(PracticeAddRow1("Computer skills", ""));

                table.Append(PracticeAddRow2("Languages", "Speaking", "Reading", "Writing"));

                if (ds.Tables[20].Rows.Count >= 1)
                {
                    for (int i = 0; i <= ds.Tables[20].Rows.Count - 1; i++)
                    {
                        if (ds.Tables[20].Rows[i]["Skills"].ToString() == "Language")
                        {
                            string s1, s2, s3;

                            if(ds.Tables[20].Rows[i]["ProfR"].ToString()== "True")
                            {
                                s1 = "Yes";
                            }
                            else
                            {
                                s1 = "No";
                            }
                            if (ds.Tables[20].Rows[i]["ProfW"].ToString() == "True")
                            {
                                s2 = "Yes";
                            }
                            else
                            {
                                s2 = "No";
                            }
                            
                             if (ds.Tables[20].Rows[i]["ProfS"].ToString() == "True")
                            {
                                s3 = "Yes";
                            }
                             else
                            {
                                s3 = "No";
                            }

                            table.Append(PracticeAddRow2(ds.Tables[20].Rows[i]["Language"].ToString(), s1.ToString(), s2.ToString(), s3.ToString()));
                        }
                    }
                }

                table.Append(PracticeAddRC1(ParaHeading("Certification", "P1")));
                table.Append(PracticeAddRC1(ParaHeading("1. I am willing to work on the Project and I will be available for entire duration of the Project assignment as required.", "P1")));
                table.Append(PracticeAddRC1(ParaHeading("2. I, the undersigned, certify that to the best of my knowledge and belief, this CV correctly describes my qualifications, my experience and me.", "P1")));
                table.Append(PracticeAddRC1(ParaHeading("Full name of staff member & Signature(Scanned): ", "P1")));
                table.Append(PracticeAddRC1(ParaHeading("Attach scanned copy of the registration certificate with professional bodies(e.g: LEED AP, Council of Architecture, etc.,", "P1")));
                table.Append(PracticeAddRC1(ParaHeading("Resume updated as on: 20th January 2014", "P1")));

                body.Append(table);
                mainPart.Document.Append(body);
                mainPart.Document.Save();
            }
        }
        catch (Exception ex)
        {
            //HttpContext.Current.Response.Write(ex.ToString());
        }
    }

    public static void WBTemplate(DataSet ds, string strUserCode)
        {
            PracticeTemplate(ds, strUserCode);
        }

    public static TableRow RMAddRowC1(Paragraph p1)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        GridSpan gridSpan;
        tr = new TableRow();
        tc = new TableCell();
        tc.Append(p1);
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "100%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 2;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }
    public static TableRow RMAddRowC2(Paragraph p1, Paragraph p2)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        TableCellWidth tw;
        tr = new TableRow();
        tc = new TableCell(p1);
        tw = new TableCellWidth();
        tc.Append(new TableCellProperties(
               new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "3500" }));
        tr.Append(tc);

        tc = new TableCell();
        tc.Append(p2);
        tc.Append(new TableCellProperties(
              new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "6500" }));
        tcp = new TableCellProperties();
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }
    public static TableRow RMAddRowC21(Paragraph p1, Paragraph p2)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        TableCellWidth tw;
        tr = new TableRow();
        tc = new TableCell(p1);
        tw = new TableCellWidth();
        tc.Append(new TableCellProperties(
               new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "3500" }));

        tcp = new TableCellProperties();

        VerticalMerge vm = new VerticalMerge()
        {
            Val = MergedCellValues.Continue
        };

        tcp.Append(vm);

        tr.Append(tc);

        tc = new TableCell();
        tc.Append(p2);
        tc.Append(new TableCellProperties(
              new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "6500" }));
        tcp = new TableCellProperties();
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }


    public static TableRow ARAddRC0(Paragraph p1)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        GridSpan gridSpan;
        tr = new TableRow();
        tc = new TableCell(p1);
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "100%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 6;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }
    public static TableRow ARAddRC1(Paragraph p1)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        GridSpan gridSpan;
        tr = new TableRow();
        tc = new TableCell(p1);
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "100%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 6;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }
    public static TableRow ARAddRC2(Paragraph p1)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        GridSpan gridSpan;
        tr = new TableRow();
        tc = new TableCell(p1);
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "100%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 6;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }
    public static TableRow ARAddRC3(Paragraph p1)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        GridSpan gridSpan;
        tr = new TableRow();
        tc = new TableCell(p1);
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "100%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 6;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }
    public static TableRow ARAddRC4(string strText1, string strText2, string strText3, string strText4, string strText5)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        GridSpan gridSpan;
        TableCellWidth tw;
        tr = new TableRow();
        tc = new TableCell(new Paragraph(new Run(new Text(strText1))));
        tw = new TableCellWidth();
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "20%" }));
        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText2))));
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "20%" }));

        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText3))));
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "20%" }));
        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText4))));
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "20%" }));
        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText5))));
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "15%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 4;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }
    public static TableRow ARAddRow6(string strText1, string strText2, string strText3)
    {
        TableRow tr;
        TableCell tc;
        TableCellWidth tw;
        TableCellProperties tcp;
        GridSpan gridSpan;
        tr = new TableRow();

        tc = new TableCell(new Paragraph(new Run(new Text(strText1))));
        tw = new TableCellWidth();
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "20%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 2;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);

        tc = new TableCell(new Paragraph(new Run(new Text(strText2))));
        tw = new TableCellWidth();
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "10%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 1;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText3))));
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "100%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 2;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }
   

    public static TableRow ADBAddRow1(string strText1, string strText2)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        TableCellWidth tw;
        GridSpan gridSpan;
        tr = new TableRow();
        tc = new TableCell(new Paragraph(new Run(new Text(strText1))));
        tw = new TableCellWidth();
        tc.Append(new TableCellProperties(
               new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "40%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 2;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText2))));
        tc.Append(new TableCellProperties(
        new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "100%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 3;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }
    public static TableRow ADBAddRow6(string strText1, string strText2, string strText3)
    {
        TableRow tr;
        TableCell tc;
        TableCellWidth tw;
        TableCellProperties tcp;
        GridSpan gridSpan;
        tr = new TableRow();

        tc = new TableCell(new Paragraph(new Run(new Text(strText1))));
        tw = new TableCellWidth();
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "20%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 2;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
   

        tc = new TableCell(new Paragraph(new Run(new Text(strText2))));
        tw = new TableCellWidth();
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "10" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 1;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
 

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText3))));
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "100%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 2;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }
    public static TableRow ADBAddRC5(string strText1, string strText2, string strText3, string strText4, string strText5)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        GridSpan gridSpan;
        TableCellWidth tw;
        tr = new TableRow();
        tc = new TableCell(new Paragraph(new Run(new Text(strText1))));
        tw = new TableCellWidth();
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "20%" }));
        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText2))));
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "20%" }));

        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText3))));
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "20%" }));
        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText4))));
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "20%" }));
        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText5))));
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "20" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 1;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }
    public static TableRow ADBAddRC6(Paragraph p1)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        GridSpan gridSpan;
        tr = new TableRow();
        tc = new TableCell(p1);
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "100%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 5;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }



    public static TableRow PracticeAddRow1(string strText1, string strText2)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        TableCellWidth tw;
        GridSpan gridSpan;
        tr = new TableRow();
        tc = new TableCell(new Paragraph(new Run(new Text(strText1))));
        tw = new TableCellWidth();
        tc.Append(new TableCellProperties(
               new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "40%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 2;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText2))));
        tc.Append(new TableCellProperties(
        new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "100%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 4;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }
    public static TableRow PracticeAddRow2(string strText1, string strText2, string strText3, string strText4)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        TableCellWidth tw;
        GridSpan gridSpan;
        tr = new TableRow();
        tc = new TableCell(new Paragraph(new Run(new Text(strText1))));
        tw = new TableCellWidth();
        tc.Append(new TableCellProperties(
               new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "25%" }));
        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText2))));
        tc.Append(new TableCellProperties(
              new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "25%" }));
        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText3))));
        tc.Append(new TableCellProperties(
              new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "25%" }));
        tr.Append(tc);

        tc = new TableCell();

        tc.Append(new Paragraph(new Run(new Text(strText4))));
        tc.Append(new TableCellProperties(
              new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "25%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 6;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }
    public static TableRow PracticeAddRow3(string strText1, string strText2, string strText3)
    {
        TableRow tr;
        TableCell tc;
        TableCellWidth tw;
        TableCellProperties tcp;
        GridSpan gridSpan;
        tr = new TableRow();

        tc = new TableCell(new Paragraph(new Run(new Text(strText1))));
        tw = new TableCellWidth();
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "20%" }));
        tr.Append(tc);

        tc = new TableCell(new Paragraph(new Run(new Text(strText2))));
        tw = new TableCellWidth();
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "10" }));
        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText3))));
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "100%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 4;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }
    public static TableRow PracticeAddRC1(Paragraph p1)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        GridSpan gridSpan;
        tr = new TableRow();
        tc = new TableCell(p1);
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "100%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 4;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }


    
    public static TableRow AddRow1(string strText1, string strText2)
        {
            TableRow tr;
            TableCell tc;
            TableCellProperties tcp;
            TableCellWidth tw;
            GridSpan gridSpan;
            tr = new TableRow();
            tc = new TableCell(new Paragraph(new Run(new Text(strText1))));
            tw = new TableCellWidth();
            tc.Append(new TableCellProperties(
                   new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "40%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 2;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(strText2))));
            tc.Append(new TableCellProperties(
            new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "100%" }));
            tcp = new TableCellProperties();
            gridSpan = new GridSpan();
            gridSpan.Val = 4;
            tcp.Append(gridSpan);
            tc.Append(tcp);
            tr.Append(tc);
            return tr;
        }
    public static TableRow AddRow2(string strText1, string strText2, string strText3, string strText4)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        TableCellWidth tw;
        GridSpan gridSpan;
        tr = new TableRow();
        tc = new TableCell(new Paragraph(new Run(new Text(strText1))));
        tw = new TableCellWidth();
        //tc.Append(new TableCellProperties(
        //        new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "3500" }));
        tc.Append(new TableCellProperties(
               new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "25%" }));
        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText2))));
        tc.Append(new TableCellProperties(
              new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "25%" }));
        tr.Append(tc);

        tc = new TableCell();
        tc.Append(new Paragraph(new Run(new Text(strText3))));
        tc.Append(new TableCellProperties(
              new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "25%" }));
        tr.Append(tc);

        tc = new TableCell();

        tc.Append(new Paragraph(new Run(new Text(strText4))));
        tc.Append(new TableCellProperties(
              new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "25%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val =6;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }
    public static TableRow AddRow3(Paragraph p1)
    {
        TableRow tr;
        TableCell tc;
        TableCellProperties tcp;
        GridSpan gridSpan;
        tr = new TableRow();
        tc = new TableCell(p1);
        tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "100%" }));
        tcp = new TableCellProperties();
        gridSpan = new GridSpan();
        gridSpan.Val = 3;
        tcp.Append(gridSpan);
        tc.Append(tcp);
        tr.Append(tc);
        return tr;
    }

    public static Style CreateStyle(string name, string fnt, string clr, int size, string fs)
    {
        RunProperties rPr = new RunProperties();

        Color color = new Color() { Val = clr.ToString() };

        

        RunFonts rFont = new RunFonts();
        rFont.Ascii = fnt.ToString();
        rPr.Append(color);
        rPr.Append(rFont);
      
        rPr.Append(new Bold());
        rPr.Append(new FontSize() { Val = size.ToString() });

        Style style = new Style();
        style.StyleId = name;
        style.Append(new Name() { Val = name });
        style.Append(new BasedOn() { Val = name });
        style.Append(new NextParagraphStyle() { Val = fs });
        style.Append(rPr);

        return style;
    }

    public static Paragraph ParaHeading(string strText1, string style)
    {
        ParagraphProperties pp2 = new ParagraphProperties();
        pp2.ParagraphStyleId = new ParagraphStyleId() { Val = style };

        ParagraphProperties User_heading_pPr = new ParagraphProperties();
        Justification CenterHeading = new Justification() { Val = JustificationValues.Center };
        User_heading_pPr.Append(CenterHeading);
        User_heading_pPr.ParagraphStyleId = new ParagraphStyleId() { Val = "userheading" };
        pp2.Append(User_heading_pPr);

        Paragraph heading = new Paragraph();
        Run heading_run = new Run();
        Text heading_text = new Text(strText1);
        heading.Append(pp2);
        heading_run.Append(heading_text);
        heading.Append(heading_run);

       
        return heading;
    }

    public static TableProperties TabelStyle(BorderValues tblBorder)
    {
        TableProperties tblPr1 = new TableProperties();
        TableBorders tblBorders1 = new TableBorders();
        tblBorders1.TopBorder = new TopBorder();
        tblBorders1.TopBorder.Val = new EnumValue<BorderValues>(tblBorder);
        tblBorders1.BottomBorder = new BottomBorder();
        tblBorders1.BottomBorder.Val = new EnumValue<BorderValues>(tblBorder);
        tblBorders1.LeftBorder = new LeftBorder();
        tblBorders1.LeftBorder.Val = new EnumValue<BorderValues>(tblBorder);
        tblBorders1.RightBorder = new RightBorder();
        tblBorders1.RightBorder.Val = new EnumValue<BorderValues>(tblBorder);
        tblBorders1.InsideHorizontalBorder = new InsideHorizontalBorder();
        tblBorders1.InsideHorizontalBorder.Val = tblBorder;
        tblBorders1.InsideVerticalBorder = new InsideVerticalBorder();
        tblBorders1.InsideVerticalBorder.Val = tblBorder;
        tblPr1.Append(tblBorders1);
        return tblPr1;
    }

    private static Paragraph AddImageToBody(string relationshipId)
    {
        var element =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = 990000L, Cy = 792000L },
                        new DW.EffectExtent()
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DW.DocProperties()
                        {
                            Id = (UInt32Value)1U,
                            Name = "Picture 1"
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties()
                                        {
                                            Id = (UInt32Value)0U,
                                            Name = "New Bitmap Image.jpg"
                                        },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip(
                                            new A.BlipExtensionList(
                                                new A.BlipExtension()
                                                {
                                                    Uri =
                                                    "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                })
                                        )
                                        {
                                            Embed = relationshipId,
                                            CompressionState =
                                            A.BlipCompressionValues.Print
                                        },
                                        new A.Stretch(
                                            new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset() { X = 0L, Y = 0L },
                                            new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                        new A.PresetGeometry(
                                            new A.AdjustValueList()
                                        )
                                        { Preset = A.ShapeTypeValues.Rectangle }))
                            )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                    )
                    {
                        DistanceFromTop = (UInt32Value)0U,
                        DistanceFromBottom = (UInt32Value)0U,
                        DistanceFromLeft = (UInt32Value)0U,
                        DistanceFromRight = (UInt32Value)0U,
                        EditId = "50D07946"
                    });

        return new Paragraph(new Run(element));
    }

    private static Stream GetStreamFromUrl(string url)
    {
        byte[] imageData = null;
        var wc = new System.Net.WebClient();

        try
        {
            imageData = wc.DownloadData(url);
            return new MemoryStream(imageData);
        }
        catch (Exception ex)
        {
            imageData = wc.DownloadData(System.IO.Path.Combine(HttpContext.Current.Server.MapPath("./Images/"), "nophoto.gif"));
            return new MemoryStream(imageData);
        }
    }
    

}

