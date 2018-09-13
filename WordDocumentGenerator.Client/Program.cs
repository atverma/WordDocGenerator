// ----------------------------------------------------------------------
// <copyright file="Program.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Schema;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using WordDocumentGenerator.Library;

namespace WordDocumentGenerator.Client
{
    class Program
    {
        // Content Control Tags
        private static string PlaceholderIgnoreA = "PlaceholderIgnoreA";
        private static string PlaceholderIgnoreB = "PlaceholderIgnoreB";

        private static string PlaceholderContainerA = "PlaceholderContainerA";

        private static string PlaceholderRecursiveA = "PlaceholderRecursiveA";
        private static string PlaceholderRecursiveB = "PlaceholderRecursiveB";

        private static string PlaceholderNonRecursiveA = "PlaceholderNonRecursiveA";
        private static string PlaceholderNonRecursiveB = "PlaceholderNonRecursiveB";
        private static string PlaceholderNonRecursiveC = "PlaceholderNonRecursiveC";
        private static string PlaceholderNonRecursiveD = "PlaceholderNonRecursiveD";

        static void Main(string[] args)
        {
            // There are three sample document templates i.e. Test_Template - 1.docx, Test_Template - 1.docx and Test_Template - Final.docx
            // The only difference between Test_Template - 1.docx and Test_Template - 2.docx is that Test_Template - 2.docx contains table. This table has it's own
            // content control tags. Rest is same. Hence I'll be reusing the code inside generators wherever possible.
            // Test_Template - Final.docx contains only cover page and is used to create a final report by appending documents.

            Console.WriteLine("Started execution of samples ...");
            Console.WriteLine("Generated documents will be saved to - " + Path.GetFullPath(@"Sample Templates\"));
            Console.WriteLine();

            GenerateDocumentUsingSampleDocGenerator();
            GenerateDocumentUsingSampleRefreshableDocGenerator();
            RefreshDocumentUsingSampleRefreshableDocGenerator();
            GenerateDocumentUsingDocWithTableGenerator();
            RefreshDocumentUsingDocWithTableGenerator();
            GenerateDocumentUsingSampleDocGeneratorUsingDatabinding();
            RefreshDocumentUsingSampleDocGeneratorUsingDatabinding();
            GenerateDocumentUsingSampleDocWithTableGeneratorUsingDatabinding();
            GenerateDocumentUsingSampleDocGeneratorUsingXml();
            GenerateDocumentUsingSampleDocGeneratorUsingXmlAndDatabinding();
            GenerateDocumentUsingSampleGenericDocGeneratorUsingXml();
            GenerateFinalReportByAppendingDocuments();            

            Console.WriteLine();
            Console.WriteLine("Execution Completed.");
            Console.WriteLine("Press any key to exit ...");
            Console.ReadKey();
        }

        /// <summary>
        /// Generates the document using sample doc generator.
        /// </summary>
        private static void GenerateDocumentUsingSampleDocGenerator()
        {
            // Test document generation from template("Test_Template - 1.docx")
            DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SampleDocumentGenerator", "1.0", GetDataContext(),
                                                    "Test_Template - 1.docx", false);

            SampleDocumentGenerator sampleDocumentGenerator = new SampleDocumentGenerator(generationInfo);
            byte[] result = result = sampleDocumentGenerator.GenerateDocument();
            WriteOutputToFile("Test_Template1_Out.docx", "Test_Template - 1.docx", result);
        }

        /// <summary>
        /// Generates the document using sample refreshable doc generator.
        /// </summary>
        /// <returns></returns>
        private static SampleRefreshableDocumentGenerator GenerateDocumentUsingSampleRefreshableDocGenerator()
        {
            // Test refreshable document generation from template("Test_Template - 1.docx")
            DocumentGenerationInfo generationInfo = generationInfo = GetDocumentGenerationInfo("SampleRefreshableDocumentGenerator", "1.0", GetDataContext(),
                                                    "Test_Template - 1.docx", false);
            SampleRefreshableDocumentGenerator sampleRefreshableDocumentGenerator = new SampleRefreshableDocumentGenerator(generationInfo);
            byte[] result = sampleRefreshableDocumentGenerator.GenerateDocument();
            WriteOutputToFile("Test_Template1_BeforeRefresh_Out.docx", "Test_Template - 1.docx", result);
            return sampleRefreshableDocumentGenerator;
        }

        /// <summary>
        /// Refreshes the document using sample refreshable doc generator.
        /// </summary>
        private static void RefreshDocumentUsingSampleRefreshableDocGenerator()
        {
            // Test refreshable document refresh from template("Test_Template1_BeforeRefresh_Out.docx")
            DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SampleRefreshableDocumentGenerator", "1.0", GetDataContextRefresh(GetDataContext()),
                                                    "Test_Template1_BeforeRefresh_Out.docx", false);
            SampleRefreshableDocumentGenerator sampleRefreshableDocumentGenerator = new SampleRefreshableDocumentGenerator(generationInfo);
            byte[] result = sampleRefreshableDocumentGenerator.GenerateDocument();
            WriteOutputToFile("Test_Template1_AfterRefresh_Out.docx", "Test_Template1_BeforeRefresh_Out.docx", result);
        }

        /// <summary>
        /// Generates the document using sample doc generator using databinding.
        /// </summary>
        private static void GenerateDocumentUsingSampleDocGeneratorUsingDatabinding()
        {
            // Test document generation from template("Test_Template - 1.docx.docx")
            DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SampleDocumentGeneratorUsingDatabinding", "1.0", GetDataContext(),
                                                    "Test_Template - 1.docx", true);
            SampleDocumentGeneratorUsingDatabinding sampleDocumentGeneratorUsingDatabinding = new SampleDocumentGeneratorUsingDatabinding(generationInfo);
            byte[] result = sampleDocumentGeneratorUsingDatabinding.GenerateDocument();
            WriteOutputToFile("Test_Template1_Databinding_Out_BeforeRefresh.docx", "Test_Template - 1.docx", result);
        }

        /// <summary>
        /// Refreshes the document using sample doc generator using databinding.
        /// </summary>
        private static void RefreshDocumentUsingSampleDocGeneratorUsingDatabinding()
        {
            // Test document refresh from template("Test_Template1_Databinding_Out_BeforeRefresh.docx")
            DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SampleDocumentGeneratorUsingDatabinding", "1.0", GetDataContextRefresh(GetDataContext()),
                                                    "Test_Template1_Databinding_Out_BeforeRefresh.docx", true);
            SampleDocumentGeneratorUsingDatabinding sampleDocumentGeneratorUsingDatabinding = new SampleDocumentGeneratorUsingDatabinding(generationInfo);
            byte[] result = sampleDocumentGeneratorUsingDatabinding.GenerateDocument();
            WriteOutputToFile("Test_Template1_Databinding_Out_AfterRefresh.docx", "Test_Template1_Databinding_Out_BeforeRefresh.docx", result);
        }

        /// <summary>
        /// Generates the document using sample doc with table generator using databinding.
        /// </summary>
        private static void GenerateDocumentUsingSampleDocWithTableGeneratorUsingDatabinding()
        {
            // Test document generation with table generation from template("Test_Template - 2.docx")
            DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SampleDocumentWithTableGeneratorUsingDatabinding", "1.0", GetDataContextRefresh(GetDataContext()),
                                                    "Test_Template - 2.docx", true);
            SampleDocumentWithTableGeneratorUsingDatabinding sampleDocumentWithTableGeneratorUsingDatabinding = new SampleDocumentWithTableGeneratorUsingDatabinding(generationInfo);
            byte[] result = sampleDocumentWithTableGeneratorUsingDatabinding.GenerateDocument();
            WriteOutputToFile("Test_Template2_Databinding_Table_Out.docx", "Test_Template - 2.docx", result);
        }

        /// <summary>
        /// Generates the document using doc with table generator.
        /// </summary>
        private static void GenerateDocumentUsingDocWithTableGenerator()
        {
            // Test refreshable document with table generation from template("Test_Template - 2.docx")
            DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SampleDocumentWithTableGenerator", "1.0", GetDataContext(),
                                                    "Test_Template - 2.docx", false);
            SampleDocumentWithTableGenerator sampleDocumentWithTableGenerator = new SampleDocumentWithTableGenerator(generationInfo);
            byte[] result = sampleDocumentWithTableGenerator.GenerateDocument();
            WriteOutputToFile("Test_Template2_Out.docx", "Test_Template - 2.docx", result);
        }

        /// <summary>
        /// Refreshes the document using doc with table generator.
        /// </summary>
        private static void RefreshDocumentUsingDocWithTableGenerator()
        {
            // Test refreshable document with table refresh from template("Test_Template2_Out.docx")
            DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SampleDocumentWithTableGenerator", "1.0", GetDataContextRefresh(GetDataContext()),
                                                    "Test_Template2_Out.docx", false);
            SampleDocumentWithTableGenerator sampleDocumentWithTableGenerator = new SampleDocumentWithTableGenerator(generationInfo);
            byte[] result = sampleDocumentWithTableGenerator.GenerateDocument();
            WriteOutputToFile("Test_Template2_AfterRefresh_Out.docx", "Test_Template2_Out.docx", result);
        }

        /// <summary>
        /// Generates the document using sample doc generator using XML.
        /// </summary>
        private static void GenerateDocumentUsingSampleDocGeneratorUsingXml()
        {
            // Generate Content Controls using Xml
            Dictionary<string, ContentControlXmlMetadata> placeHolderTagToContentControlXmlMetadataCollection = new Dictionary<string, ContentControlXmlMetadata>();

            // Handle ignore placeholders
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderIgnoreA, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderIgnoreA, Type = PlaceHolderType.Ignore });
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderIgnoreB, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderIgnoreA, Type = PlaceHolderType.Ignore });

            // Handle container placeholders            
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderContainerA, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderContainerA, Type = PlaceHolderType.Container, ControlTagXPath = "./Id[1]" });

            // Handle recursive placeholders            
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderRecursiveA, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderRecursiveA, Type = PlaceHolderType.Recursive, ControlValueXPath = "./vendors[1]" });
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderRecursiveB, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderRecursiveB, Type = PlaceHolderType.Recursive, ControlValueXPath = "./items[1]" });

            // Handle non recursive placeholders
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderNonRecursiveA, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderNonRecursiveA, Type = PlaceHolderType.NonRecursive, ControlTagXPath = "./Id[1]", ControlValueXPath = "./Name[1]" });
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderNonRecursiveB, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderNonRecursiveB, Type = PlaceHolderType.NonRecursive, ControlTagXPath = "./Id[1]", ControlValueXPath = "./Name[1]" });
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderNonRecursiveC, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderNonRecursiveC, Type = PlaceHolderType.NonRecursive, ControlTagXPath = "./Id[1]", ControlValueXPath = "./Name[1]" });
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderNonRecursiveD, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderNonRecursiveD, Type = PlaceHolderType.NonRecursive, ControlTagXPath = "./Id[1]", ControlValueXPath = "./Description[1]" });

            // Test document generation from template("Test_Template - 1.docx")
            string dataAsXml = "<Order><vendors><Vendor><Id>469c8927-6a68-4f16-b267-acaf38fc2d39</Id><Name>Vendor 1</Name></Vendor><Vendor><Id>48e1f6d0-5060-4725-ae45-51e81b2e89d6</Id><Name>Vendor 2</Name></Vendor><Vendor><Id>b59d289f-93b3-4d03-81f0-e6f49657e611</Id><Name>Vendor 111</Name></Vendor><Vendor><Id>165d4fe3-d445-47ac-b40e-3611a40c4845</Id><Name>Vendor 222</Name></Vendor><Vendor><Id>6d950039-075b-47ea-96f0-6088feca5c9b</Id><Name>Vendor 113</Name></Vendor><Vendor><Id>fe31e8f1-0927-4780-8843-0e166977a505</Id><Name>Vendor 224</Name></Vendor><Vendor><Id>a9edafc3-9117-4499-8820-806ec59a0bc9</Id><Name>Vendor 115</Name></Vendor><Vendor><Id>7eca4fc3-5218-4090-a1cf-10db411b668e</Id><Name>Vendor 226</Name></Vendor><Vendor><Id>8d491555-da72-41a3-ad77-7c001e93b052</Id><Name>Vendor 117</Name></Vendor><Vendor><Id>08654a92-601c-4c38-9218-00a54bc44ec6</Id><Name>Vendor 228</Name></Vendor></vendors><items><Item><Id>81474c98-0094-499d-88d7-678f40581b50</Id><Name>Item 1</Name></Item><Item><Id>60a5d7bc-304a-49fa-be4a-684f91adf6c5</Id><Name>Item 2</Name></Item><Item><Id>bc9d5b75-3861-4bf8-bb26-35689e6b557a</Id><Name>Item 11</Name></Item><Item><Id>72fafb82-179d-4115-a4e7-84bf1a62049d</Id><Name>Item 21</Name></Item></items><Id>736d79d9-0689-419e-b27a-1867f613fd80</Id><Name>Test Order</Name><Description>Test Order Description</Description></Order>";
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(dataAsXml);
            xmlDoc.Schemas.Add(null, @"Test Data\TestDataSchema.xsd");
            xmlDoc.Validate(ValidationCallBack, xmlDoc.DocumentElement);

            DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SampleDocumentGeneratorUsingXml", "1.0", xmlDoc.DocumentElement,
                                                    "Test_Template - 1.docx", false);
            SampleDocumentGeneratorUsingXml sampleDocumentGeneratorUsingXml = new SampleDocumentGeneratorUsingXml(generationInfo, placeHolderTagToContentControlXmlMetadataCollection);
            byte[] result = sampleDocumentGeneratorUsingXml.GenerateDocument();
            WriteOutputToFile("SampleDocumentGeneratorUsingXml.docx", "Test_Template - 1.docx", result);
        }

        /// <summary>
        /// Generates the document using sample doc generator using XML.
        /// </summary>
        private static void GenerateDocumentUsingSampleDocGeneratorUsingXmlAndDatabinding()
        {
            // Generate Content Controls using Xml
            Dictionary<string, ContentControlXmlMetadata> placeHolderTagToContentControlXmlMetadataCollection = new Dictionary<string, ContentControlXmlMetadata>();

            // Handle ignore placeholders
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderIgnoreA, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderIgnoreA, Type = PlaceHolderType.Ignore });
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderIgnoreB, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderIgnoreA, Type = PlaceHolderType.Ignore });

            // Handle container placeholders            
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderContainerA, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderContainerA, Type = PlaceHolderType.Container, ControlTagXPath = "./Id[1]" });

            // Handle recursive placeholders            
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderRecursiveA, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderRecursiveA, Type = PlaceHolderType.Recursive, ControlValueXPath = "./vendors[1]" });
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderRecursiveB, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderRecursiveB, Type = PlaceHolderType.Recursive, ControlValueXPath = "./items[1]" });

            // Handle non recursive placeholders
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderNonRecursiveA, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderNonRecursiveA, Type = PlaceHolderType.NonRecursive, ControlTagXPath = "./Id[1]", ControlValueXPath = "/ns0:DocumentRootNode[1]/ns0:DataBoundControlsDataStore[1]/ns0:Data[1]/Order[1]/vendors[1]/Vendor[{0}]/Name[1]" });
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderNonRecursiveB, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderNonRecursiveB, Type = PlaceHolderType.NonRecursive, ControlTagXPath = "./Id[1]", ControlValueXPath = "/ns0:DocumentRootNode[1]/ns0:DataBoundControlsDataStore[1]/ns0:Data[1]/Order[1]/items[1]/Item[{0}]/Name[1]" });
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderNonRecursiveC, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderNonRecursiveC, Type = PlaceHolderType.NonRecursive, ControlTagXPath = "./Id[1]", ControlValueXPath = "/ns0:DocumentRootNode[1]/ns0:DataBoundControlsDataStore[1]/ns0:Data[1]/Order[1]/Name[1]" });
            placeHolderTagToContentControlXmlMetadataCollection.Add(PlaceholderNonRecursiveD, new ContentControlXmlMetadata() { PlaceHolderName = PlaceholderNonRecursiveD, Type = PlaceHolderType.NonRecursive, ControlTagXPath = "./Id[1]", ControlValueXPath = "/ns0:DocumentRootNode[1]/ns0:DataBoundControlsDataStore[1]/ns0:Data[1]/Order[1]/Description[1]" });

            // Test document generation from template("Test_Template - 1.docx")
            string dataAsXml = "<Order><vendors><Vendor><Id>469c8927-6a68-4f16-b267-acaf38fc2d39</Id><Name>Vendor 1</Name></Vendor><Vendor><Id>48e1f6d0-5060-4725-ae45-51e81b2e89d6</Id><Name>Vendor 2</Name></Vendor><Vendor><Id>b59d289f-93b3-4d03-81f0-e6f49657e611</Id><Name>Vendor 111</Name></Vendor><Vendor><Id>165d4fe3-d445-47ac-b40e-3611a40c4845</Id><Name>Vendor 222</Name></Vendor><Vendor><Id>6d950039-075b-47ea-96f0-6088feca5c9b</Id><Name>Vendor 113</Name></Vendor><Vendor><Id>fe31e8f1-0927-4780-8843-0e166977a505</Id><Name>Vendor 224</Name></Vendor><Vendor><Id>a9edafc3-9117-4499-8820-806ec59a0bc9</Id><Name>Vendor 115</Name></Vendor><Vendor><Id>7eca4fc3-5218-4090-a1cf-10db411b668e</Id><Name>Vendor 226</Name></Vendor><Vendor><Id>8d491555-da72-41a3-ad77-7c001e93b052</Id><Name>Vendor 117</Name></Vendor><Vendor><Id>08654a92-601c-4c38-9218-00a54bc44ec6</Id><Name>Vendor 228</Name></Vendor></vendors><items><Item><Id>81474c98-0094-499d-88d7-678f40581b50</Id><Name>Item 1</Name></Item><Item><Id>60a5d7bc-304a-49fa-be4a-684f91adf6c5</Id><Name>Item 2</Name></Item><Item><Id>bc9d5b75-3861-4bf8-bb26-35689e6b557a</Id><Name>Item 11</Name></Item><Item><Id>72fafb82-179d-4115-a4e7-84bf1a62049d</Id><Name>Item 21</Name></Item></items><Id>736d79d9-0689-419e-b27a-1867f613fd80</Id><Name>Test Order</Name><Description>Test Order Description</Description></Order>";
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(dataAsXml);
            xmlDoc.Schemas.Add(null, @"Test Data\TestDataSchema.xsd");
            xmlDoc.Validate(ValidationCallBack, xmlDoc.DocumentElement);

            DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SampleDocumentGeneratorUsingXmlAndDataBinding", "1.0", xmlDoc.DocumentElement,
                                                    "Test_Template - 1.docx", true);
            SampleDocumentGeneratorUsingXmlAndDataBinding sampleDocumentGeneratorUsingXmlAndDataBinding = new SampleDocumentGeneratorUsingXmlAndDataBinding(generationInfo, placeHolderTagToContentControlXmlMetadataCollection, false);
            byte[] result = sampleDocumentGeneratorUsingXmlAndDataBinding.GenerateDocument();
            WriteOutputToFile("SampleDocumentGeneratorUsingXmlAndDataBinding.docx", "Test_Template - 1.docx", result);
        }

        /// <summary>
        /// Generates the document using sample generic doc generator using XML.
        /// </summary>
        private static void GenerateDocumentUsingSampleGenericDocGeneratorUsingXml()
        {
            // Test document generation from template("Test_Template - 1.docx")
            string dataAsXml = "<field>" +
                                    "<listFields>" + 
                                        "<field orderName=\"Test Order\" orderDescription=\"Test Order Description\" id=\"\" contentControlTagREFS=\"PlaceholderNonRecursiveC PlaceholderNonRecursiveD\"/>" +
                                        "<field id=\"736d79d9-0689-419e-b27a-1867f613fd80\" contentControlTagREFS=\"PlaceholderContainerA\">" +
                                            "<listFields>" +
                                                "<field id=\"736d79d9-0689-419e-b17a-1867f613fd80\" contentControlTagREFS=\"PlaceholderRecursiveA\">" +
                                                    "<listFields>" +
                                                        "<field vendorName=\"Vendor 1\" id=\"736d79d9-1689-419e-c27a-1867f613fd80\" contentControlTagREFS=\"PlaceholderNonRecursiveA\"/>" +
                                                        "<field vendorName=\"Vendor 2\" id=\"736d79d9-0589-419e-c27a-1867f613fd80\" contentControlTagREFS=\"PlaceholderNonRecursiveA\"/>" +
                                                    "</listFields>" +
                                            "</field>" +
                                                "<field id=\"736d79d9-0689-119e-b27a-1867f613fd80\" contentControlTagREFS=\"PlaceholderRecursiveB\">" +
                                                    "<listFields>" +
                                                        "<field itemName=\"Item 1\" id=\"736d79d9-0689-419e-c27a-1867f613fd80\" contentControlTagREFS=\"PlaceholderNonRecursiveB\" />" +
                                                        "<field itemName=\"Item 2\" id=\"736d79d9-0689-419e-b27a-1867f613dd80\" contentControlTagREFS=\"PlaceholderNonRecursiveB\" />" +
                                                    "</listFields>" +
                                                "</field>" +
                                            "</listFields>" +
                                        "</field>" +
                                    "</listFields>" +
                                    "<contentControls>" +
                                        "<contentControl type=\"2\" tag=\"PlaceholderNonRecursiveC\" refControlValue=\"orderName\" refTagValue=\"id\"/>" +
                                        "<contentControl type=\"2\" tag=\"PlaceholderNonRecursiveD\" refControlValue=\"orderDescription\" refTagValue=\"id\"/>" +
                                        "<contentControl type=\"4\" tag=\"PlaceholderContainerA\" refTagValue=\"id\"/>" +
                                        "<contentControl type=\"1\" tag=\"PlaceholderRecursiveA\" refTagValue=\"id\"/>" +
                                        "<contentControl type=\"1\" tag=\"PlaceholderRecursiveB\" refTagValue=\"id\"/>" +
                                        "<contentControl type=\"2\" tag=\"PlaceholderNonRecursiveA\" refControlValue=\"vendorName\" refTagValue=\"id\"/>" +
                                        "<contentControl type=\"2\" tag=\"PlaceholderNonRecursiveB\" refControlValue=\"itemName\" refTagValue=\"id\"/>" +
                                    "</contentControls>" +
                                "</field>";


            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(dataAsXml);
            DocumentGenerationInfo generationInfo = GetDocumentGenerationInfo("SampleGenericDocumentGeneratorUsingXml", "1.0", xmlDoc.DocumentElement,
                                        "Test_Template - 1.docx", false);
            SampleGenericDocumentGeneratorUsingXml sampleGenericDocumentGeneratorUsingXml = new SampleGenericDocumentGeneratorUsingXml(generationInfo);
            byte[] result = sampleGenericDocumentGeneratorUsingXml.GenerateDocument();
            WriteOutputToFile("SampleGenericDocumentGeneratorUsingXml1.docx", "Test_Template - 1.docx", result);

            // Test document generation from template("Test_Template - 2.docx")
            dataAsXml = "<field>" +
                                    "<listFields>" +
                                        "<field orderName=\"Test Order\" orderDescription=\"Test Order Description\" id=\"\" contentControlTagREFS=\"PlaceholderNonRecursiveC PlaceholderNonRecursiveD\"/>" +
                                        "<field id=\"736d79d9-0689-419e-b27a-1867f613fd80\" contentControlTagREFS=\"PlaceholderContainerA\">" +
                                            "<listFields>" +
                                                "<field id=\"736d79d9-0689-419e-b17a-1867f613fd80\" contentControlTagREFS=\"PlaceholderRecursiveA VendorDetailRow\">" +
                                                    "<listFields>" +
                                                        "<field vendorName=\"Vendor 1\" id=\"736d79d9-1689-419e-c87a-1867f613fd80\" contentControlTagREFS=\"PlaceholderNonRecursiveA VendorId VendorName\"/>" +
                                                        "<field vendorName=\"Vendor 2\" id=\"736d79d9-0589-419e-c77a-1867f613fd80\" contentControlTagREFS=\"PlaceholderNonRecursiveA VendorId VendorName\"/>" +
                                                        "<field vendorName=\"Vendor 3\" id=\"736d79d9-0589-419e-c67a-1867f613fd80\" contentControlTagREFS=\"PlaceholderNonRecursiveA VendorId VendorName\"/>" +
                                                        "<field vendorName=\"Vendor 4\" id=\"736d79d9-0589-419e-c57a-1867f613fd80\" contentControlTagREFS=\"PlaceholderNonRecursiveA VendorId VendorName\"/>" +
                                                        "<field vendorName=\"Vendor 5\" id=\"736d79d9-0589-419e-c47a-1867f613fd80\" contentControlTagREFS=\"PlaceholderNonRecursiveA VendorId VendorName\"/>" +
                                                        "<field vendorName=\"Vendor 6\" id=\"736d79d9-0589-419e-c37a-1867f613fd80\" contentControlTagREFS=\"PlaceholderNonRecursiveA VendorId VendorName\"/>" +
                                                        "<field vendorName=\"Vendor 7\" id=\"736d79d9-0589-419e-c27a-1867f613fd80\" contentControlTagREFS=\"PlaceholderNonRecursiveA VendorId VendorName\"/>" +
                                                        "<field vendorName=\"Vendor 8\" id=\"736d79d9-0589-419e-c17a-1867f613fd80\" contentControlTagREFS=\"PlaceholderNonRecursiveA VendorId VendorName\"/>" +
                                                    "</listFields>" +
                                            "</field>" +
                                                "<field id=\"736d79d9-0689-119e-b27a-1867f613fd80\" contentControlTagREFS=\"PlaceholderRecursiveB\">" +
                                                    "<listFields>" +
                                                        "<field itemName=\"Item 1\" id=\"736d79d9-0689-419e-c27a-1867f613fd80\" contentControlTagREFS=\"PlaceholderNonRecursiveB\" />" +
                                                        "<field itemName=\"Item 2\" id=\"736d79d9-0689-419e-b27a-1867f613dd80\" contentControlTagREFS=\"PlaceholderNonRecursiveB\" />" +
                                                        "<field itemName=\"Item 3\" id=\"736d79d9-0689-419e-b12a-1867f613dd80\" contentControlTagREFS=\"PlaceholderNonRecursiveB\" />" +
                                                        "<field itemName=\"Item 4\" id=\"736d79d9-0689-419e-b22a-1867f613dd80\" contentControlTagREFS=\"PlaceholderNonRecursiveB\" />" +
                                                        "<field itemName=\"Item 5\" id=\"736d79d9-0689-419e-b43a-1867f613dd80\" contentControlTagREFS=\"PlaceholderNonRecursiveB\" />" +
                                                        "<field itemName=\"Item 6\" id=\"736d79d9-0689-419e-b87a-1867f613dd80\" contentControlTagREFS=\"PlaceholderNonRecursiveB\" />" +
                                                        "<field itemName=\"Item 7\" id=\"736d79d9-0689-419e-b98a-1867f613dd80\" contentControlTagREFS=\"PlaceholderNonRecursiveB\" />" +
                                                        "<field itemName=\"Item 8\" id=\"736d79d9-0689-419e-b19a-1867f613dd80\" contentControlTagREFS=\"PlaceholderNonRecursiveB\" />" +
                                                    "</listFields>" +
                                                "</field>" +
                                            "</listFields>" +
                                        "</field>" +
                                    "</listFields>" +
                                    "<contentControls>" +
                                        "<contentControl type=\"2\" tag=\"PlaceholderNonRecursiveC\" refControlValue=\"orderName\" refTagValue=\"id\"/>" +
                                        "<contentControl type=\"2\" tag=\"PlaceholderNonRecursiveD\" refControlValue=\"orderDescription\" refTagValue=\"id\"/>" +
                                        "<contentControl type=\"4\" tag=\"PlaceholderContainerA\" refTagValue=\"id\"/>" +
                                        "<contentControl type=\"1\" tag=\"PlaceholderRecursiveA\" refTagValue=\"id\"/>" +
                                        "<contentControl type=\"1\" tag=\"PlaceholderRecursiveB\" refTagValue=\"id\"/>" +
                                        "<contentControl type=\"2\" tag=\"PlaceholderNonRecursiveA\" refControlValue=\"vendorName\" refTagValue=\"id\"/>" +
                                        "<contentControl type=\"2\" tag=\"PlaceholderNonRecursiveB\" refControlValue=\"itemName\" refTagValue=\"id\"/>" +
                                        "<contentControl type=\"1\" tag=\"VendorDetailRow\" refTagValue=\"id\"/>" +
                                        "<contentControl type=\"2\" tag=\"VendorId\" refControlValue=\"id\" refTagValue=\"id\"/>" +
                                        "<contentControl type=\"2\" tag=\"VendorName\" refControlValue=\"vendorName\" refTagValue=\"id\"/>" +
                                    "</contentControls>" +
                                "</field>";


            xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(dataAsXml);
            generationInfo = GetDocumentGenerationInfo("SampleGenericDocumentGeneratorUsingXml", "2.0", xmlDoc.DocumentElement,
                                        "Test_Template - 2.docx", false);
            sampleGenericDocumentGeneratorUsingXml = new SampleGenericDocumentGeneratorUsingXml(generationInfo);
            result = sampleGenericDocumentGeneratorUsingXml.GenerateDocument();
            WriteOutputToFile("SampleGenericDocumentGeneratorUsingXml2.docx", "Test_Template - 2.docx", result);
        }

        /// <summary>
        /// Generates the final report by appending documents.
        /// </summary>
        private static void GenerateFinalReportByAppendingDocuments()
        {
            // Test final report i.e. created by merging documents generation
            byte[] primaryDoc = File.ReadAllBytes(@"Sample Templates\Test_Template - Final.docx");
            List<byte[]> otherDocs = new List<byte[]>();
            otherDocs.Add(File.ReadAllBytes(@"Sample Templates\Test_Template1_Out.docx"));
            otherDocs.Add(File.ReadAllBytes(@"Sample Templates\Test_Template1_BeforeRefresh_Out.docx"));
            otherDocs.Add(File.ReadAllBytes(@"Sample Templates\Test_Template1_AfterRefresh_Out.docx"));
            otherDocs.Add(File.ReadAllBytes(@"Sample Templates\Test_Template2_Out.docx"));
            otherDocs.Add(File.ReadAllBytes(@"Sample Templates\Test_Template2_AfterRefresh_Out.docx"));
            otherDocs.Add(File.ReadAllBytes(@"Sample Templates\Test_Template2_Databinding_Table_Out.docx"));
            otherDocs.Add(File.ReadAllBytes(@"Sample Templates\Test_Template1_Databinding_Out_BeforeRefresh.docx"));
            otherDocs.Add(File.ReadAllBytes(@"Sample Templates\Test_Template1_Databinding_Out_AfterRefresh.docx"));

            // Final report generation            
            OpenXmlHelper openXmlHelper = new OpenXmlHelper(DocumentGenerationInfo.NamespaceUri);
            byte[] result = openXmlHelper.AppendDocumentsToPrimaryDocument(primaryDoc, otherDocs);
            WriteOutputToFile("FinalReport.docx", "Test_Template - Final.docx", result);

            // Final Protected report generation
            using (MemoryStream msfinalDocument = new MemoryStream())
            {
                msfinalDocument.Write(result, 0, result.Length);

                using (WordprocessingDocument finalDocument = WordprocessingDocument.Open(msfinalDocument, true))
                {
                    openXmlHelper.ProtectDocument(finalDocument);
                }

                msfinalDocument.Position = 0;
                result = new byte[msfinalDocument.Length];
                msfinalDocument.Read(result, 0, result.Length);
                msfinalDocument.Close();
            }

            WriteOutputToFile("FinalReport_Protected.docx", "Test_Template - Final.docx", result);
        }

        /// <summary>
        /// Validations the call back.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="System.Xml.Schema.ValidationEventArgs"/> instance containing the event data.</param>
        private static void ValidationCallBack(object sender, ValidationEventArgs e)
        {
            Console.WriteLine(e.Message);
        }

        /// <summary>
        /// Gets the document generation info.
        /// </summary>
        /// <param name="docType">Type of the doc.</param>
        /// <param name="docVersion">The doc version.</param>
        /// <param name="dataContext">The data context.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="useDataBoundControls">if set to <c>true</c> [use data bound controls].</param>
        /// <returns></returns>
        private static DocumentGenerationInfo GetDocumentGenerationInfo(string docType, string docVersion, object dataContext, string fileName, bool useDataBoundControls)
        {
            DocumentGenerationInfo generationInfo = new DocumentGenerationInfo();
            generationInfo.Metadata = new DocumentMetadata() { DocumentType = docType, DocumentVersion = docVersion };
            generationInfo.DataContext = dataContext;
            generationInfo.TemplateData = File.ReadAllBytes(Path.Combine("Sample Templates", fileName));
            generationInfo.IsDataBoundControls = useDataBoundControls;

            return generationInfo;
        }

        /// <summary>
        /// Writes the output to file.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="templateName">Name of the template.</param>
        /// <param name="fileContents">The file contents.</param>
        private static void WriteOutputToFile(string fileName, string templateName, byte[] fileContents)
        {
            ConsoleColor consoleColor = Console.ForegroundColor;

            if (fileContents != null)
            {
                File.WriteAllBytes(Path.Combine("Sample Templates", fileName), fileContents);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(string.Format("Generation succeeded for template({0}) --> {1}", templateName, fileName));
                Console.WriteLine();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(string.Format("Generation failed for template({0}) --> {1}", templateName, fileName));
            }

            Console.ForegroundColor = consoleColor;
        }

        /// <summary>
        /// Gets the data context.
        /// </summary>
        /// <returns></returns>
        private static Order GetDataContext()
        {
            Order order = new Order();

            order.Name = "Test Order";
            order.Description = "Test Order Description";
            order.Id = Guid.NewGuid();

            order.vendors = new System.Collections.Generic.List<Vendor>();
            order.items = new System.Collections.Generic.List<Item>();

            order.vendors.Add(new Vendor() { Name = "Vendor 1", Id = Guid.NewGuid() });
            order.vendors.Add(new Vendor() { Name = "Vendor 2", Id = Guid.NewGuid() });

            order.items.Add(new Item() { Name = "Item 1", Id = Guid.NewGuid() });
            order.items.Add(new Item() { Name = "Item 2", Id = Guid.NewGuid() });

            return order;
        }

        /// <summary>
        /// Gets the data context refresh.
        /// </summary>
        /// <param name="order">The order.</param>
        /// <returns></returns>
        private static Order GetDataContextRefresh(Order order)
        {
            order.vendors.Add(new Vendor() { Name = "Vendor 111", Id = Guid.NewGuid() });
            order.vendors.Add(new Vendor() { Name = "Vendor 222", Id = Guid.NewGuid() });
            order.vendors.Add(new Vendor() { Name = "Vendor 113", Id = Guid.NewGuid() });
            order.vendors.Add(new Vendor() { Name = "Vendor 224", Id = Guid.NewGuid() });
            order.vendors.Add(new Vendor() { Name = "Vendor 115", Id = Guid.NewGuid() });
            order.vendors.Add(new Vendor() { Name = "Vendor 226", Id = Guid.NewGuid() });
            order.vendors.Add(new Vendor() { Name = "Vendor 117", Id = Guid.NewGuid() });
            order.vendors.Add(new Vendor() { Name = "Vendor 228", Id = Guid.NewGuid() });

            order.items.Add(new Item() { Name = "Item 11", Id = Guid.NewGuid() });
            order.items.Add(new Item() { Name = "Item 21", Id = Guid.NewGuid() });

            return order;
        }
    }
}
