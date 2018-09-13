// ----------------------------------------------------------------------
// <copyright file="AddInService.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to usealed the Open Xml 2.0 SDK and VS 2010 ffor document generation. They are unsupported, but you can usealed them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Client
{
    using System.Collections.Generic;
    using System.Xml;
    using WordDocumentGenerator.Library;

    /// <summary>
    /// Class for document generation and used by AddIn only. This mimics the behavior of Service i.e. AddIn can call a Service e.g. WCF and pass only the document stream. Here instead of adding a Service
    /// direct method call is provided.
    /// </summary>
    public class AddInService
    {
        const string PlaceholderIgnoreA = "PlaceholderIgnoreA";
        const string PlaceholderIgnoreB = "PlaceholderIgnoreB";

        const string PlaceholderContainerA = "PlaceholderContainerA";

        const string PlaceholderRecursiveA = "PlaceholderRecursiveA";
        const string PlaceholderRecursiveB = "PlaceholderRecursiveB";

        const string PlaceholderNonRecursiveA = "PlaceholderNonRecursiveA";
        const string PlaceholderNonRecursiveB = "PlaceholderNonRecursiveB";
        const string PlaceholderNonRecursiveC = "PlaceholderNonRecursiveC";
        const string PlaceholderNonRecursiveD = "PlaceholderNonRecursiveD";

        /// <summary>
        /// Generates the document.
        /// </summary>
        /// <param name="documentStream">The document stream.</param>
        /// <returns></returns>
        public static byte[] GenerateDocument(byte[] documentStream)
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

            DocumentGenerationInfo generationInfo = new DocumentGenerationInfo();
            generationInfo.Metadata = new DocumentMetadata() { DocumentType = "SampleDocumentGeneratorUsingXml", DocumentVersion = "1.0" };
            generationInfo.DataContext = xmlDoc.DocumentElement;
            generationInfo.TemplateData = documentStream;
            generationInfo.IsDataBoundControls = true;

            SampleDocumentGeneratorUsingXmlAndDataBinding docGen = new SampleDocumentGeneratorUsingXmlAndDataBinding(generationInfo, placeHolderTagToContentControlXmlMetadataCollection, true);

            return docGen.GenerateDocument();
        }
    }
}