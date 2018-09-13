// ----------------------------------------------------------------------
// <copyright file="OPCConverter.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.WordRefreshableDocumentAddin
{
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.IO.Packaging;
    using System.Linq;
    using System.Text;
    using System.Xml;
    using System.Xml.Linq;
    using Microsoft.Office.Core;
    using WordDocumentGenerator.Library;
    using Word = Microsoft.Office.Interop.Word;

    /// <summary>
    /// OPC helper methods
    /// </summary>
    public static class OPCHelper
    {
        /// <summary>
        /// Gets the package stream from range.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <returns></returns>
        public static byte[] GetPackageStream(this Word.Document document)
        {
            XDocument doc = XDocument.Parse(document.WordOpenXML);
            XNamespace pkg =
               "http://schemas.microsoft.com/office/2006/xmlPackage";
            XNamespace rel =
                "http://schemas.openxmlformats.org/package/2006/relationships";
            Package InmemoryPackage = null;
            byte[] output = null;

            using (MemoryStream memStream = new MemoryStream())
            {
                using (InmemoryPackage = Package.Open(memStream, FileMode.Create))
                {
                    // add all parts (but not relationships)
                    foreach (var xmlPart in doc.Root
                        .Elements()
                        .Where(p =>
                            (string)p.Attribute(pkg + "contentType") !=
                            "application/vnd.openxmlformats-package.relationships+xml"))
                    {
                        string name = (string)xmlPart.Attribute(pkg + "name");
                        string contentType = (string)xmlPart.Attribute(pkg + "contentType");
                        if (contentType.EndsWith("xml"))
                        {
                            Uri u = new Uri(name, UriKind.Relative);
                            PackagePart part = InmemoryPackage.CreatePart(u, contentType,
                                CompressionOption.Normal);
                            using (Stream str = part.GetStream(FileMode.Create))
                            using (XmlWriter xmlWriter = XmlWriter.Create(str))
                                xmlPart.Element(pkg + "xmlData")
                                    .Elements()
                                    .First()
                                    .WriteTo(xmlWriter);
                        }
                        else
                        {
                            Uri u = new Uri(name, UriKind.Relative);
                            PackagePart part = InmemoryPackage.CreatePart(u, contentType,
                                CompressionOption.Normal);
                            using (Stream str = part.GetStream(FileMode.Create))
                            using (BinaryWriter binaryWriter = new BinaryWriter(str))
                            {
                                string base64StringInChunks =
                               (string)xmlPart.Element(pkg + "binaryData");
                                char[] base64CharArray = base64StringInChunks
                                    .Where(c => c != '\r' && c != '\n').ToArray();
                                byte[] byteArray =
                                    System.Convert.FromBase64CharArray(base64CharArray,
                                    0, base64CharArray.Length);
                                binaryWriter.Write(byteArray);
                            }
                        }
                    }
                    foreach (var xmlPart in doc.Root.Elements())
                    {
                        string name = (string)xmlPart.Attribute(pkg + "name");
                        string contentType = (string)xmlPart.Attribute(pkg + "contentType");
                        if (contentType ==
                            "application/vnd.openxmlformats-package.relationships+xml")
                        {
                            // add the package level relationships
                            if (name == "/_rels/.rels")
                            {
                                foreach (XElement xmlRel in
                                    xmlPart.Descendants(rel + "Relationship"))
                                {
                                    string id = (string)xmlRel.Attribute("Id");
                                    string type = (string)xmlRel.Attribute("Type");
                                    string target = (string)xmlRel.Attribute("Target");
                                    string targetMode =
                                        (string)xmlRel.Attribute("TargetMode");
                                    if (targetMode == "External")
                                        InmemoryPackage.CreateRelationship(
                                            new Uri(target, UriKind.Absolute),
                                            TargetMode.External, type, id);
                                    else
                                        InmemoryPackage.CreateRelationship(
                                            new Uri(target, UriKind.Relative),
                                            TargetMode.Internal, type, id);
                                }
                            }
                            else
                            // add part level relationships
                            {
                                string directory = name.Substring(0, name.IndexOf("/_rels"));
                                string relsFilename = name.Substring(name.LastIndexOf('/'));
                                string filename =
                                    relsFilename.Substring(0, relsFilename.IndexOf(".rels"));
                                PackagePart fromPart = InmemoryPackage.GetPart(
                                    new Uri(directory + filename, UriKind.Relative));
                                foreach (XElement xmlRel in
                                    xmlPart.Descendants(rel + "Relationship"))
                                {
                                    string id = (string)xmlRel.Attribute("Id");
                                    string type = (string)xmlRel.Attribute("Type");
                                    string target = (string)xmlRel.Attribute("Target");
                                    string targetMode =
                                        (string)xmlRel.Attribute("TargetMode");
                                    if (targetMode == "External")
                                        fromPart.CreateRelationship(
                                            new Uri(target, UriKind.Absolute),
                                            TargetMode.External, type, id);
                                    else
                                        fromPart.CreateRelationship(
                                            new Uri(target, UriKind.Relative),
                                            TargetMode.Internal, type, id);
                                }
                            }
                        }
                    }

                    InmemoryPackage.Flush();
                }

                memStream.Position = 0;
                output = new byte[memStream.Length];
                memStream.Read(output, 0, output.Length);
            }

            return output;
        }

        /// <summary>
        /// Stores the custom XML part.
        /// </summary>
        /// <param name="document">The document.</param>
        /// <param name="customXmlPartDocument">The custom XML part document.</param>
        /// <returns></returns>
        public static CustomXMLPart StoreCustomXmlPart(this Word._Document document, XDocument customXmlPartDocument)
        {
            if (document != null && customXmlPartDocument != null)
            {
                CustomXMLParts parts = document.CustomXMLParts.SelectByNamespace(DocumentGenerationInfo.NamespaceUri);

                if (parts.Count > 0)
                {
                    Debug.Assert(parts.Count == 1);
                    parts[1].Delete();
                }

                return document.CustomXMLParts.Add(customXmlPartDocument.ToString());
            }

            return null;
        }

        /// <summary>
        /// Opcs to flat opc.
        /// </summary>
        /// <param name="package">The package.</param>
        /// <returns></returns>
        public static XDocument OpcToFlatOpc(Package package)
        {
            XNamespace
                pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
            XDeclaration
                declaration = new XDeclaration("1.0", "UTF-8", "yes");
            XDocument doc = new XDocument(
                declaration,
                new XProcessingInstruction("mso-application", "progid=\"Word.Document\""),
                new XElement(pkg + "package",
                    new XAttribute(XNamespace.Xmlns + "pkg", pkg.ToString()),
                    package.GetParts().Select(part => GetContentsAsXml(part))
                )
            );
            return doc;
        }

        /// <summary>
        /// Gets the contents as XML.
        /// </summary>
        /// <param name="part">The part.</param>
        /// <returns></returns>
        private static XElement GetContentsAsXml(PackagePart part)
        {
            XNamespace pkg =
               "http://schemas.microsoft.com/office/2006/xmlPackage";
            if (part.ContentType.EndsWith("xml"))
            {
                using (Stream partstream = part.GetStream())
                using (StreamReader streamReader = new StreamReader(partstream))
                {
                    string streamString = streamReader.ReadToEnd();
                    XElement newXElement =
                        new XElement(pkg + "part", new XAttribute(pkg + "name", part.Uri),
                            new XAttribute(pkg + "contentType", part.ContentType),
                            new XElement(pkg + "xmlData", XElement.Parse(streamString)));
                    return newXElement;
                }
            }
            else
            {
                using (Stream str = part.GetStream())
                using (BinaryReader binaryReader = new BinaryReader(str))
                {
                    int len = (int)binaryReader.BaseStream.Length;
                    byte[] byteArray = binaryReader.ReadBytes(len);
                    // the following expression creates the base64String, then chunks
                    // it to lines of 76 characters long
                    string base64String = (System.Convert.ToBase64String(byteArray))
                        .Select
                        (
                            (c, i) => new
                            {
                                Character = c,
                                Chunk = i / 76
                            }
                        )
                        .GroupBy(c => c.Chunk)
                        .Aggregate(
                            new StringBuilder(),
                            (s, i) =>
                                s.Append(
                                    i.Aggregate(
                                        new StringBuilder(),
                                        (seed, it) => seed.Append(it.Character),
                                        sb => sb.ToString()
                                    )
                                )
                                .Append(Environment.NewLine),
                            s => s.ToString()
                        );

                    return new XElement(pkg + "part",
                        new XAttribute(pkg + "name", part.Uri),
                        new XAttribute(pkg + "contentType", part.ContentType),
                        new XAttribute(pkg + "compression", "store"),
                        new XElement(pkg + "binaryData", base64String)
                    );
                }
            }
        }
    }
}
