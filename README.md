# WordDocGenerator

Codeplex project @ https://archive.codeplex.com/?p=worddocgenerator has been moved to GitHub along with updating to latest .NET Framework and DocumentFormat.OpenXml.

WordDocumentGenerator is an utility to generate Word documents from templates using Visual Studio 2017, .NET Framework 4.7 and DocumentFormat.OpenXml 2.8.1.
WordDocumentGenerator helps generate Word documents both non-refresh-able as well as refresh-able based on predefined templates using minimum code changes.
Content controls are used as placeholders for document generation.

Document generation is quite easy as code changes required are very less. Mostly one just needs to override these five methods while coding a new generator class. Sample generators have been provided in utility's source code. 

        /// <summary>
        /// Gets the place holder tag to type collection.
        /// </summary>
        /// <returns></returns>
        protected abstract Dictionary<string, PlaceHolderType> GetPlaceHolderTagToTypeCollection();

        /// <summary>
        /// Ignore placeholder found.
        /// </summary>
        /// <param name="placeholderTag">The placeholder tag.</param>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        protected abstract void IgnorePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext);

        /// <summary>
        /// Non recursive placeholder found.
        /// </summary>
        /// <param name="placeholderTag">The placeholder tag.</param>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        protected abstract void NonRecursivePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext);

        /// <summary>
        /// Recursive placeholder found.
        /// </summary>
        /// <param name="placeholderTag">The placeholder tag.</param>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        protected abstract void RecursivePlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext);

        /// <summary>
        /// Container placeholder found.
        /// </summary>
        /// <param name="placeholderTag">The placeholder tag.</param>
        /// <param name="openXmlElementDataContext">The open XML element data context.</param>
        protected abstract void ContainerPlaceholderFound(string placeholderTag, OpenXmlElementDataContext openXmlElementDataContext);

Development Status
The functionalities that can be achieved are:

Document Generation
Generate document from a Word template using content controls as place holders and populate controls with data(Object) e.g. [SampleDocumentGenerator, SampleRefreshableDocumentGenerator, SampleDocumentWithTableGenerator]
Generate document from a Word template using content controls as place holders(data bound content controls) and populate controls with data(Object is serialized to Xml) e.g. [SampleDocumentGeneratorUsingDatabinding, SampleDocumentWithTableGeneratorUsingDatabinding, SampleDocumentGeneratorUsingXmlAndDatabinding]
Refresh the document from within the document(e.g. right click on document and click Refresh) using document-level projects for Word 2007, Word 2010 and Word 2016
Generate document from a Word template using content controls as place holders and populate controls with data(XmlNode) e.g. [SampleDocumentGeneratorUsingXml]
Generate document from a Word template using content controls as place holders(data bound content controls) and populate controls with data(XmlNode) e.g. [SampleDocumentGeneratorUsingXmlAndDatabinding]
Generate document that can be
Standalone: Once generated document cannot be refreshed.
Refreshable: Once generated document can be refreshed. Content controls will be added/updated/deleted and content control's content will be refreshed as per data.
Append documents using AltChunk
Protect Document
UnProtect Document
Removal of Content Controls from a document while keeping contents
Removal of Foot notes
Ensuring the each content control has unique Id's by fixing the duplicate Id's if any for a document
Serializing an Object to Xml using XmlSerializer(Used for document generation using data bound content controls as serialized object is written to CustomXmlPart)

Content Controls
Set text of a content control(not applicable for data bound content controls)
Get text from a content control(not applicable for data bound content controls)
Set text of content control while keeping PermStart and PermEnd elements(not applicable for data bound content controls)
Set Tag of a content control
Get Tag of a content control
Set data binding of a content control
Set text of a data bound content control from CustomXmlPart manually. This is helpful in cases when CustomXmlPart needs to be removed and this copies the text from the CustomXmlPart node using XPath.

CustomXmlPart
Adding a CustomXmlPart to a document
Removing CustomXmlPart from a document
Getting CustomXmlPart from a document
Add/Update a Xml element node inside CustomXmlPart. This is required
To keep Document related metadata e.g. Document type, version etc.
To make the Document self-refreshable. In this case the container content control is persisted inside a Placeholder node, the first time document is generated from template. Onwards when refreshing document we fetch the container content control from CustomXmlPart
Saving the Xml e.g. serialized object which will be the data store for data bound content controls


You can read more about it at
http://blogs.msdn.com/b/atverma/archive/2011/12/31/utility-to-generate-word-documents-from-templates-using-visual-studio-2010-and-open-xml-2-0-sdk.aspx
http://blogs.msdn.com/b/atverma/archive/2012/01/08/utility-to-generate-word-documents-from-templates-using-visual-studio-2010-and-open-xml-2-0-sdk-part-2-samples-updated.aspx
http://blogs.msdn.com/b/atverma/archive/2012/01/11/utility-to-generate-word-documents-from-templates-using-visual-studio-2010-and-open-xml-2-0-sdk-part-3.aspx
http://blogs.msdn.com/b/atverma/archive/2012/01/11/utility-to-generate-word-documents-from-templates-using-visual-studio-2010-and-open-xml-2-0-sdk-part-4.aspx
