// ----------------------------------------------------------------------
// <copyright file="ThisDocument.cs" author="Atul Verma">
//     Copyright (c) Atul Verma. This utility along with samples demonstrate how to use the Open Xml 2.0 SDK and VS 2010 for document generation. They are unsupported, but you can use them as-is.
// </copyright>
// ------------------------------------------------------------------------

namespace WordDocumentGenerator.WordRefreshableDocumentAddin
{
    using System.Collections.Generic;
    using System.IO;
    using System.Xml;
    using System.Xml.Linq;
    using DocumentFormat.OpenXml.Packaging;
    using Microsoft.Office.Core;
    using WordDocumentGenerator.Client;
    using WordDocumentGenerator.Library;

    /// <summary>
    /// Word document that can be refreshed from Server(instead of Service call showed a direct call to Proxy). At service side WordDocumentGenerator API will be there. 
    /// 1. Pass document stream to server i.e. byte[]
    /// 2. Server generates/refreshed the document and returns the document stream i.e. byte[]
    /// 3. Refresh the document contents
    /// </summary>
    public partial class ThisDocument
    {        
        Microsoft.Office.Interop.Word.Application app;
        CommandBars commandbars = null;
        CommandBar textCommandBar = null;
        CommandBarButton refreshDocumentCommandBarButton = null;
        List<string> commandBarsTags = new List<string>();

        /// <summary>
        /// Handles the Startup event of the ThisDocument control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            app = ThisApplication;
            commandbars = Globals.ThisDocument.CommandBars;
            textCommandBar = commandbars["Text"] as CommandBar;
            refreshDocumentCommandBarButton = AddCommandBar(textCommandBar, new _CommandBarButtonEvents_ClickEventHandler(RefreshDocumentCommandBarButton_Click), 1, "refreshDocument", "Refresh Document");
        }

        /// <summary>
        /// Handles the Shutdown event of the ThisDocument control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
            foreach (CommandBarControl v in textCommandBar.Controls)
            {
                if (commandBarsTags.Contains(v.Tag))
                {
                    v.Delete(missing);
                }
            }

            app = null;
        }

        /// <summary>
        /// Adds the command bar.
        /// </summary>
        /// <param name="cmdBr">The CMD br.</param>
        /// <param name="handler">The handler.</param>
        /// <param name="index">The index.</param>
        /// <param name="tag">The tag.</param>
        /// <param name="caption">The caption.</param>
        /// <returns></returns>
        private CommandBarButton AddCommandBar(CommandBar cmdBr, _CommandBarButtonEvents_ClickEventHandler handler, int index, string tag, string caption)
        {
            CommandBarButton cmdBtn = (CommandBarButton)cmdBr.FindControl(MsoControlType.msoControlButton, 0, tag, missing, missing);

            if ((cmdBtn != null))
            {
                cmdBtn.Delete(true);
            }

            cmdBtn = (CommandBarButton)cmdBr.Controls.Add(MsoControlType.msoControlButton, missing, missing, index, true);
            cmdBtn.Style = MsoButtonStyle.msoButtonCaption;
            cmdBtn.Caption = caption;
            cmdBtn.Tag = tag;
            cmdBtn.Visible = true;

            cmdBtn.Click -= handler;
            cmdBtn.Click += handler;

            if (!commandBarsTags.Contains(tag))
            {
                commandBarsTags.Add(tag);
            }

            return cmdBtn;
        }

        /// <summary>
        /// Refreshes the document command bar button_ click.
        /// </summary>
        /// <param name="cmdBarbutton">The CMD barbutton.</param>
        /// <param name="cancel">if set to <c>true</c> [cancel].</param>
        private void RefreshDocumentCommandBarButton_Click(CommandBarButton cmdBarbutton, ref bool cancel)
        {
            app.ScreenUpdating = false;            
            Microsoft.Office.Interop.Word.Document doc = app.ActiveDocument;
            
            // Get the active documents as stream of bytes
            byte[] input = doc.GetPackageStream();            

            // Generate document on the Server. AddInService can be a proxy to service, however here it's direct call
            byte[] output = AddInService.GenerateDocument(input);

            if (output != null)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    ms.Write(output, 0, output.Length);

                    using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(ms, true))
                    {
                        XDocument xDoc = OPCHelper.OpcToFlatOpc(wordDocument.Package);
                        string openxml = xDoc.ToString();                        
                        doc.Range().InsertXML(openxml);

                        // Add CustomXmlPart
                        CustomXmlPartCore customXmlPartCore = new CustomXmlPartCore(DocumentGenerationInfo.NamespaceUri);
                        CustomXmlPart customPart = customXmlPartCore.GetCustomXmlPart(wordDocument.MainDocumentPart);

                        if (customPart != null)
                        {
                            XDocument customPartDoc = null;

                            using (XmlReader reader = XmlReader.Create(customPart.GetStream(FileMode.Open, FileAccess.Read)))
                            {
                                customPartDoc = XDocument.Load(reader);
                            }

                            doc.StoreCustomXmlPart(customPartDoc);
                        }                       
                    }
                }
            }

            app.ScreenUpdating = true;
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(ThisDocument_Shutdown);
        }

        #endregion
    }
}
