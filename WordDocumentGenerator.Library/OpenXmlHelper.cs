// ----------------------------------------------------------------------
// <copyright file="OpenXmlHelper.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.Library
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Xml;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;

    /// <summary>
    /// Helper class for OpenXml operations for document generation
    /// </summary>
    public class OpenXmlHelper
    {
        #region Memebers

        private readonly string NamespaceUri = string.Empty;

        #endregion
        
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlHelper"/> class.
        /// </summary>
        /// <param name="NamespaceUri">The namespace URI.</param>
        public OpenXmlHelper(string NamespaceUri)
        {
            this.NamespaceUri = NamespaceUri;
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Validates and returns errors of document.
        /// </summary>
        /// <param name="filepath">The document path.</param>
        /// <param name="errors">The document errors.</param>
        /// <returns>Returns true if document has no errors</returns>
        public static bool ValidateWordDocument(string filepath, out List<string> errors)
        {
            errors = new List<string>();

            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
            {
                try
                {
                    OpenXmlValidator validator = new OpenXmlValidator();

                    foreach (ValidationErrorInfo error in validator.Validate(wordprocessingDocument))
                    {
                        errors.Add($"Error: {errors.Count} Description: {error.Description} ErrorType: {error.ErrorType} Node: {error.Node} Path: {error.Path.XPath} Part: { error.Part.Uri} ");
                    }
                }
                catch (Exception ex)
                {
                    errors.Add(ex.Message);
                }

                wordprocessingDocument.Close();
            }

            return errors.Count <= 0;
        }

        /// <summary>
        /// Appends the documents to primary document.
        /// </summary>
        /// <param name="primaryDocument">The primary document.</param>
        /// <param name="documentstoAppend">The documents to append.</param>
        /// <returns></returns>
        public byte[] AppendDocumentsToPrimaryDocument(byte[] primaryDocument, List<byte[]> documentstoAppend)
        {
            if (documentstoAppend == null)
            {
                throw new ArgumentNullException("documentstoAppend");
            }

            if (primaryDocument == null)
            {
                throw new ArgumentNullException("primaryDocument");
            }

            byte[] output = null;

            using (MemoryStream finalDocumentStream = new MemoryStream())
            {
                finalDocumentStream.Write(primaryDocument, 0, primaryDocument.Length);

                using (WordprocessingDocument finalDocument = WordprocessingDocument.Open(finalDocumentStream, true))
                {
                    SectionProperties finalDocSectionProperties = null;
                    this.UnprotectDocument(finalDocument);

                    SectionProperties tempSectionProperties = finalDocument.MainDocumentPart.Document.Descendants<SectionProperties>().LastOrDefault();

                    if (tempSectionProperties != null)
                    {
                        finalDocSectionProperties = tempSectionProperties.CloneNode(true) as SectionProperties;
                    }

                    this.RemoveContentControlsAndKeepContents(finalDocument.MainDocumentPart.Document);

                    foreach (byte[] documentToAppend in documentstoAppend)
                    {
                        AlternativeFormatImportPart subReportPart = finalDocument.MainDocumentPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.WordprocessingML);
                        SectionProperties secProperties = null;

                        using (MemoryStream docToAppendStream = new MemoryStream())
                        {
                            docToAppendStream.Write(documentToAppend, 0, documentToAppend.Length);

                            using (WordprocessingDocument docToAppend = WordprocessingDocument.Open(docToAppendStream, true))
                            {
                                this.UnprotectDocument(docToAppend);

                                tempSectionProperties = docToAppend.MainDocumentPart.Document.Descendants<SectionProperties>().LastOrDefault();

                                if (tempSectionProperties != null)
                                {
                                    secProperties = tempSectionProperties.CloneNode(true) as SectionProperties;
                                }

                                this.RemoveContentControlsAndKeepContents(docToAppend.MainDocumentPart.Document);
                                docToAppend.MainDocumentPart.Document.Save();
                            }

                            docToAppendStream.Position = 0;
                            subReportPart.FeedData(docToAppendStream);
                        }

                        if (documentstoAppend.ElementAtOrDefault(0).Equals(documentToAppend))
                        {
                            AssignSectionProperties(finalDocument.MainDocumentPart.Document, finalDocSectionProperties);
                        }

                        AltChunk altChunk = new AltChunk();
                        altChunk.Id = finalDocument.MainDocumentPart.GetIdOfPart(subReportPart);
                        finalDocument.MainDocumentPart.Document.AppendChild(altChunk);

                        if (!documentstoAppend.ElementAtOrDefault(documentstoAppend.Count - 1).Equals(documentToAppend))
                        {
                            AssignSectionProperties(finalDocument.MainDocumentPart.Document, secProperties);
                        }

                        finalDocument.MainDocumentPart.Document.Save();
                    }

                    finalDocument.MainDocumentPart.Document.Save();
                }

                finalDocumentStream.Position = 0;
                output = new byte[finalDocumentStream.Length];
                finalDocumentStream.Read(output, 0, output.Length);
                finalDocumentStream.Close();
            }

            return output;
        }

        /// <summary>
        /// Gets the SDT content of content control.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <returns></returns>
        public OpenXmlCompositeElement GetSdtContentOfContentControl(SdtElement element)
        {
            SdtRun sdtRunELement = element as SdtRun;
            SdtBlock sdtBlockElement = element as SdtBlock;
            SdtCell sdtCellElement = element as SdtCell;
            SdtRow sdtRowElement = element as SdtRow;

            if (sdtRunELement != null)
            {
                return sdtRunELement.SdtContentRun;
            }
            else if (sdtBlockElement != null)
            {
                return sdtBlockElement.SdtContentBlock;
            }
            else if (sdtCellElement != null)
            {
                return sdtCellElement.SdtContentCell;
            }
            else if (sdtRowElement != null)
            {
                return sdtRowElement.SdtContentRow;
            }

            return null;
        }

        /// <summary>
        /// Protects the document.
        /// </summary>
        /// <param name="wordprocessingDocument">The wordprocessing document.</param>
        public void ProtectDocument(WordprocessingDocument wordprocessingDocument)
        {
            if (wordprocessingDocument == null)
            {
                throw new ArgumentNullException("wordprocessingDocument");
            }

            DocumentSettingsPart documentSettingsPart = wordprocessingDocument.MainDocumentPart.GetPartsOfType<DocumentSettingsPart>().FirstOrDefault();

            if (documentSettingsPart != null)
            {
                var documentProtection = documentSettingsPart.Settings.Elements<DocumentProtection>().FirstOrDefault();

                if (documentProtection != null)
                {
                    documentProtection.Enforcement = true;
                }
                else
                {
                    documentProtection = new DocumentProtection() { Edit = DocumentProtectionValues.Comments, Enforcement = true, CryptographicProviderType = CryptProviderValues.RsaFull, CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash, CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny, CryptographicAlgorithmSid = 4, CryptographicSpinCount = (UInt32Value)100000U, Hash = "2krUoz1qWd0WBeXqVrOq81l8xpk=", Salt = "9kIgmDDYtt2r5U2idCOwMA==" };
                    documentSettingsPart.Settings.Append(documentProtection);
                }
            }

            wordprocessingDocument.MainDocumentPart.Document.Save();
        }

        /// <summary>
        /// Unprotects the document.
        /// </summary>
        /// <param name="wordprocessingDocument">The wordprocessing document.</param>
        public void UnprotectDocument(WordprocessingDocument wordprocessingDocument)
        {
            if (wordprocessingDocument == null)
            {
                throw new ArgumentNullException("wordprocessingDocument");
            }

            DocumentSettingsPart documentSettingsPart = wordprocessingDocument.MainDocumentPart.GetPartsOfType<DocumentSettingsPart>().FirstOrDefault();

            if (documentSettingsPart != null)
            {
                var documentProtection = documentSettingsPart.Settings.Elements<DocumentProtection>().FirstOrDefault();

                if (documentProtection != null)
                {
                    documentProtection.Remove();
                }
            }

            List<OpenXmlLeafElement> permElements = new List<OpenXmlLeafElement>();

            foreach (var permStart in wordprocessingDocument.MainDocumentPart.Document.Body.Descendants<PermStart>())
            {
                if (!permElements.Contains(permStart))
                {
                    permElements.Add(permStart);
                }
            }

            foreach (var permEnd in wordprocessingDocument.MainDocumentPart.Document.Body.Descendants<PermEnd>())
            {
                if (!permElements.Contains(permEnd))
                {
                    permElements.Add(permEnd);
                }
            }

            foreach (var permElem in permElements)
            {
                if (permElem.Parent != null)
                {
                    permElem.Remove();
                }
            }

            wordprocessingDocument.MainDocumentPart.Document.Save();
        }

        /// <summary>
        /// Removes the content controls and keep contents.
        /// </summary>
        /// <param name="document">The document.</param>
        public void RemoveContentControlsAndKeepContents(Document document)
        {
            if (document == null)
            {
                throw new ArgumentNullException("document");
            }

            CustomXmlPartCore customXmlPartCore = new CustomXmlPartCore(this.NamespaceUri);
            CustomXmlPart customXmlPart = customXmlPartCore.GetCustomXmlPart(document.MainDocumentPart);
            XmlDocument customPartDoc = new XmlDocument();

            if (customXmlPart != null)
            {
                using (XmlReader reader = XmlReader.Create(customXmlPart.GetStream(FileMode.Open, FileAccess.Read)))
                {
                    customPartDoc.Load(reader);
                }
            }

            RemoveContentControlsAndKeepContents(document.Body, customPartDoc.DocumentElement);
            document.Save();
        }

        /// <summary>
        /// Removes the content controls and keep contents.
        /// </summary>
        /// <param name="compositeElement">The composite element.</param>
        /// <param name="customXmlPartDocElement">The custom XML part doc element.</param>
        public void RemoveContentControlsAndKeepContents(OpenXmlCompositeElement compositeElement, XmlElement customXmlPartDocElement)
        {
            if (compositeElement == null)
            {
                throw new ArgumentNullException("compositeElement");
            }

            if (compositeElement is SdtElement)
            {
                IList<OpenXmlCompositeElement> elementsList = RemoveContentControlAndKeepContents(compositeElement as SdtElement, customXmlPartDocElement);

                foreach (var innerCompositeElement in elementsList)
                {
                    RemoveContentControlsAndKeepContents(innerCompositeElement, customXmlPartDocElement);
                }
            }
            else
            {
                var childCompositeElements = compositeElement.Elements<OpenXmlCompositeElement>().ToList();

                foreach (var childCompositeElement in childCompositeElements)
                {
                    RemoveContentControlsAndKeepContents(childCompositeElement, customXmlPartDocElement);
                }
            }
        }

        /// <summary>
        /// Assigns the content from custom XML part for databound control.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="customPartDocElement">The custom part doc element.</param>
        public void AssignContentFromCustomXmlPartForDataboundControl(SdtElement element, XmlElement customPartDocElement)
        {
            // This fix is applied only for data bound content controls. It was found MergeDocuments method was not picking up data from CustomXmlPart. Thus
            // default text of the content control was there in the Final report instead of the XPath value.
            // This method copies the text from the CustomXmlPart using XPath specified while creating the Binding element and assigns that to the
            // content control

            DataBinding binding = element.SdtProperties.GetFirstChild<DataBinding>();

            if (binding != null)
            {
                if (binding.XPath.HasValue)
                {
                    string path = binding.XPath.Value;

                    if (customPartDocElement != null)
                    {
                        XmlNamespaceManager mgr = new XmlNamespaceManager(new NameTable());
                        mgr.AddNamespace("ns0", this.NamespaceUri);
                        XmlNode node = customPartDocElement.SelectSingleNode(path, mgr);

                        if (node != null)
                        {
                            this.SetContentOfContentControl(element, node.InnerText);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Gets the text from content control.
        /// </summary>
        /// <param name="contentControl">The content control.</param>
        /// <returns></returns>
        public string GetTextFromContentControl(SdtElement contentControl)
        {
            string result = null;

            if (contentControl != null)
            {
                if (contentControl is SdtRun)
                {
                    if (IsContentControlMultiline(contentControl))
                    {
                        var runs = contentControl.Descendants<SdtContentRun>().First().Elements()
                           .Where(elem => elem is Run || elem is InsertedRun);

                        List<string> runTexts = new List<string>();

                        if (runs != null)
                        {
                            foreach (var run in runs)
                            {
                                foreach (var runChild in run.Elements())
                                {
                                    Text runText = runChild as Text;
                                    Break runBreak = runChild as Break;

                                    if (runText != null)
                                    {
                                        runTexts.Add(runText.Text);
                                    }
                                    else if (runBreak != null)
                                    {
                                        runTexts.Add(Environment.NewLine);
                                    }
                                }
                            }
                        }

                        StringBuilder stringBuilder = new StringBuilder();

                        foreach (string item in runTexts)
                        {
                            stringBuilder.Append(item);
                        }

                        result = stringBuilder.ToString();
                    }
                    else
                    {
                        result = contentControl.InnerText;
                    }
                }
                else
                {
                    result = contentControl.InnerText;
                }
            }

            return result;
        }

        /// <summary>
        /// Generates the paragraph.
        /// </summary>
        /// <returns>
        /// Returns new Paragraph with empty run
        /// </returns>
        public Paragraph GenerateParagraph()
        {
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            paragraph.Append(run);
            return paragraph;
        }

        /// <summary>
        /// Ensures the unique content control ids for main document part.
        /// </summary>
        /// <param name="mainDocumentPart">The main document part.</param>
        public void EnsureUniqueContentControlIdsForMainDocumentPart(MainDocumentPart mainDocumentPart)
        {
            List<int> contentControlIds = new List<int>();

            if (mainDocumentPart != null)
            {
                foreach (HeaderPart part in mainDocumentPart.HeaderParts)
                {
                    SetUniquecontentControlIds(part.Header, contentControlIds);
                    part.Header.Save();
                }

                foreach (FooterPart part in mainDocumentPart.FooterParts)
                {
                    SetUniquecontentControlIds(part.Footer, contentControlIds);
                    part.Footer.Save();
                }

                SetUniquecontentControlIds(mainDocumentPart.Document.Body, contentControlIds);
                mainDocumentPart.Document.Save();
            }
        }

        /// <summary>
        /// Sets the unique content control ids.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="existingIds">The existing ids.</param>
        public void SetUniquecontentControlIds(OpenXmlCompositeElement element, List<int> existingIds)
        {
            Random randomizer = new Random();

            foreach (SdtId sdtId in element.Descendants<SdtId>())
            {
                if (existingIds.Contains(sdtId.Val))
                {
                    int randomId = randomizer.Next(int.MaxValue);

                    while (existingIds.Contains(randomId))
                    {
                        randomizer.Next(int.MaxValue);
                    }

                    sdtId.Val.Value = randomId;
                }
                else
                {
                    existingIds.Add(sdtId.Val);
                }
            }
        }

        /// <summary>
        /// Sets the content of content control.
        /// </summary>
        /// <param name="contentControl">The content control.</param>
        /// <param name="content">The content.</param>
        public void SetContentOfContentControl(SdtElement contentControl, string content)
        {
            if (contentControl == null)
            {
                throw new ArgumentNullException("contentControl");
            }

            content = string.IsNullOrEmpty(content) ? string.Empty : content;
            bool isCombobox = contentControl.SdtProperties.Descendants<SdtContentDropDownList>().FirstOrDefault() != null;            

            if (isCombobox)
            {
                OpenXmlCompositeElement openXmlCompositeElement = GetSdtContentOfContentControl(contentControl);
                Run run = CreateRun(openXmlCompositeElement, content);
                SetSdtContentKeepingPermissionElements(openXmlCompositeElement, run);
            }
            else
            {                
                OpenXmlCompositeElement openXmlCompositeElement = GetSdtContentOfContentControl(contentControl);
                contentControl.SdtProperties.RemoveAllChildren<ShowingPlaceholder>();
                List<Run> runs = new List<Run>();

                if (IsContentControlMultiline(contentControl))
                {
                    List<string> textSplitted = content.Split(Environment.NewLine.ToCharArray()).ToList();
                    bool addBreak = false;

                    foreach (string textSplit in textSplitted)
                    {
                        Run run = CreateRun(openXmlCompositeElement, textSplit);                        

                        if (addBreak)
                        {
                            run.AppendChild<Break>(new Break());
                        }

                        if (!addBreak)
                        {
                            addBreak = true;
                        }

                        runs.Add(run);
                    }
                }
                else
                {
                    runs.Add(CreateRun(openXmlCompositeElement, content));                    
                }

                if (openXmlCompositeElement is SdtContentCell)
                {
                    AddRunsToSdtContentCell(openXmlCompositeElement as SdtContentCell, runs);
                }
                else if (openXmlCompositeElement is SdtContentBlock)
                {
                    Paragraph para = CreateParagraph(openXmlCompositeElement, runs);
                    SetSdtContentKeepingPermissionElements(openXmlCompositeElement, para);
                }
                else
                {
                    SetSdtContentKeepingPermissionElements(openXmlCompositeElement, runs);
                }
            }
        }

        /// <summary>
        /// Sets the tag value.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="tagValue">The tag value.</param>
        public void SetTagValue(SdtElement element, string tagValue)
        {
            Tag tag = GetTag(element);
            tag.Val.Value = tagValue;
        }

        /// <summary>
        /// Gets the tag.
        /// </summary>
        /// <param name="element">The SDT element.</param>
        /// <returns>
        /// Returns Tag of content control
        /// </returns>
        public Tag GetTag(SdtElement element)
        {
            if (element == null)
                throw new ArgumentNullException("element");

            return element.SdtProperties.Elements<Tag>().FirstOrDefault();
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Determines whether [is content control multiline] [the specified content control].
        /// </summary>
        /// <param name="contentControl">The content control.</param>
        /// <returns>
        ///   <c>true</c> if [is content control multiline] [the specified content control]; otherwise, <c>false</c>.
        /// </returns>
        private static bool IsContentControlMultiline(SdtElement contentControl)
        {
            SdtContentText contentText = contentControl.SdtProperties.Elements<SdtContentText>().FirstOrDefault();

            bool isMultiline = false;

            if (contentText != null && contentText.MultiLine != null)
            {
                isMultiline = contentText.MultiLine.Value == true;
            }
            return isMultiline;
        }

        /// <summary>
        /// Sets the SDT content keeping permission elements.
        /// </summary>
        /// <param name="openXmlCompositeElement">The open XML composite element.</param>
        /// <param name="newChild">The new child.</param>
        private void SetSdtContentKeepingPermissionElements(OpenXmlCompositeElement openXmlCompositeElement, OpenXmlElement newChild)
        {
            PermStart start = openXmlCompositeElement.Descendants<PermStart>().FirstOrDefault();
            PermEnd end = openXmlCompositeElement.Descendants<PermEnd>().FirstOrDefault();
            openXmlCompositeElement.RemoveAllChildren();

            if (start != null)
            {
                openXmlCompositeElement.AppendChild(start);
            }

            openXmlCompositeElement.AppendChild(newChild);

            if (end != null)
            {
                openXmlCompositeElement.AppendChild(end);
            }
        }

        /// <summary>
        /// Sets the SDT content keeping permission elements.
        /// </summary>
        /// <param name="openXmlCompositeElement">The open XML composite element.</param>
        /// <param name="newChildren">The new children.</param>
        private void SetSdtContentKeepingPermissionElements(OpenXmlCompositeElement openXmlCompositeElement, List<Run> newChildren)
        {
            PermStart start = openXmlCompositeElement.Descendants<PermStart>().FirstOrDefault();
            PermEnd end = openXmlCompositeElement.Descendants<PermEnd>().FirstOrDefault();
            openXmlCompositeElement.RemoveAllChildren();

            if (start != null)
            {
                openXmlCompositeElement.AppendChild(start);
            }

            foreach (var newChild in newChildren)
            {
                openXmlCompositeElement.AppendChild(newChild);
            }

            if (end != null)
            {
                openXmlCompositeElement.AppendChild(end);
            }
        }

        /// <summary>
        /// Sets the SDT content keeping permission elements.
        /// </summary>
        /// <param name="openXmlCompositeElement">The open XML composite element.</param>
        /// <param name="newChildren">The new children.</param>
        private void SetSdtContentKeepingPermissionElements(OpenXmlCompositeElement openXmlCompositeElement, List<OpenXmlElement> newChildren)
        {
            PermStart start = openXmlCompositeElement.Descendants<PermStart>().FirstOrDefault();
            PermEnd end = openXmlCompositeElement.Descendants<PermEnd>().FirstOrDefault();
            openXmlCompositeElement.RemoveAllChildren();

            if (start != null)
            {
                openXmlCompositeElement.AppendChild(start);
            }

            foreach (var newChild in newChildren)
            {
                openXmlCompositeElement.AppendChild(newChild);
            }

            if (end != null)
            {
                openXmlCompositeElement.AppendChild(end);
            }
        }

        /// <summary>
        /// Adds the runs to SDT content cell.
        /// </summary>
        /// <param name="sdtContentCell">The SDT content cell.</param>
        /// <param name="runs">The runs.</param>
        private void AddRunsToSdtContentCell(SdtContentCell sdtContentCell, List<Run> runs)
        {
            TableCell cell = new TableCell();
            Paragraph para = new Paragraph();
            para.RemoveAllChildren();

            foreach (Run run in runs)
            {
                para.AppendChild<Run>(run);
            }

            cell.AppendChild<Paragraph>(para);
            SetSdtContentKeepingPermissionElements(sdtContentCell, cell);
        }

        /// <summary>
        /// Removes the content control and keep contents.
        /// </summary>
        /// <param name="contentControl">The content control.</param>
        /// <param name="customXmlPartDocElement">The custom XML part doc element.</param>
        /// <returns></returns>
        private IList<OpenXmlCompositeElement> RemoveContentControlAndKeepContents(SdtElement contentControl, XmlElement customXmlPartDocElement)
        {
            IList<OpenXmlCompositeElement> elementsList = new List<OpenXmlCompositeElement>();

            AssignContentFromCustomXmlPartForDataboundControl(contentControl, customXmlPartDocElement);

            foreach (var elem in GetSdtContentOfContentControl(contentControl).Elements())
            {
                OpenXmlElement newElement = contentControl.Parent.InsertBefore(elem.CloneNode(true), contentControl);
                AddToListIfCompositeElement(elementsList, newElement);
            }

            contentControl.Remove();
            return elementsList;
        }

        /// <summary>
        /// Adds to list if composite element.
        /// </summary>
        /// <param name="elementsList">The elements list.</param>
        /// <param name="newElement">The new element.</param>
        private void AddToListIfCompositeElement(IList<OpenXmlCompositeElement> elementsList, OpenXmlElement newElement)
        {
            OpenXmlCompositeElement compositeElement = newElement as OpenXmlCompositeElement;

            if (elementsList == null)
            {
                throw new ArgumentNullException("elementsList");
            }

            if (compositeElement != null)
            {
                elementsList.Add(compositeElement);
            }
        }

        /// <summary>
        /// Assigns the section properties.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="secProperties">The sec properties.</param>
        private void AssignSectionProperties(Document document, SectionProperties secProperties)
        {
            if (document == null)
            {
                throw new ArgumentNullException("document");
            }

            if (secProperties != null)
            {
                PageSize pageSize = secProperties.Descendants<PageSize>().FirstOrDefault();

                if (pageSize != null)
                {
                    pageSize.Remove();
                }

                PageMargin pageMargin = secProperties.Descendants<PageMargin>().FirstOrDefault();

                if (pageMargin != null)
                {
                    pageMargin.Remove();
                }

                document.AppendChild(new Paragraph(new ParagraphProperties(new SectionProperties(pageSize, pageMargin))));
            }
        }

        /// <summary>
        /// Creates the paragraph.
        /// </summary>
        /// <param name="openXmlCompositeElement">The open XML composite element.</param>
        /// <param name="runs">The runs.</param>
        /// <returns></returns>
        private static Paragraph CreateParagraph(OpenXmlCompositeElement openXmlCompositeElement, List<Run> runs)
        {
            ParagraphProperties paragraphProperties = openXmlCompositeElement.Descendants<ParagraphProperties>().FirstOrDefault();
            Paragraph para = null;

            if (paragraphProperties != null)
            {
                para = new Paragraph(paragraphProperties.CloneNode(true));
                foreach (Run run in runs)
                {
                    para.AppendChild<Run>(run);
                }
            }
            else
            {
                para = new Paragraph();
                foreach (Run run in runs)
                {
                    para.AppendChild<Run>(run);
                }
            }
            return para;
        }

        /// <summary>
        /// Creates the run.
        /// </summary>
        /// <param name="openXmlCompositeElement">The open XML composite element.</param>
        /// <param name="content">The content.</param>
        /// <returns></returns>
        private static Run CreateRun(OpenXmlCompositeElement openXmlCompositeElement, string content)
        {   
            RunProperties runProperties = openXmlCompositeElement.Descendants<RunProperties>().FirstOrDefault();
            Run run = null;

            if (runProperties != null)
            {
                run = new Run(runProperties.CloneNode(true), new Text(content));
            }
            else
            {
                run = new Run(new Text(content));
            }

            return run;
        }

        #endregion
    }
}