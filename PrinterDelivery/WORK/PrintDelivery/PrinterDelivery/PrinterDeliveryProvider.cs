/*=====================================================================

  File:      PrinterDeliveryProvider.cs
  
  Summary:   Represents delivery provider that can be used to
             deliver report s to a printer.

------------------------------------------------------------------------

  This file is part of Microsoft SQL Server Code Samples.
  Copyright (C) Microsoft Corporation.  All rights reserved.

  This source code is intended only as a supplement to Microsoft
  Development Tools and/or on-line documentation.  See these other
  materials for detailed information regarding Microsoft code samples.

 THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
 KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 PARTICULAR PURPOSE.

======================================================================== */

using System;
using System.IO;
using System.Xml;
using System.Collections;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.ReportingServices.Interfaces;

namespace Unact.ReportingServices.PrinterDelivery
{
    // A delivery provider must implement IDeliveryExtension and IExtension
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1001:TypesThatOwnDisposableFieldsShouldBeDisposable")]
    public sealed class PrinterDeliveryProvider : IDeliveryExtension, IExtension
    {
        // Private member variables used to support the delivery extension
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1823:AvoidUnusedPrivateFields")]
        private bool m_isPrivilegedUser;
        private IDeliveryReportServerInformation m_reportServerInfo;
        private Setting[] m_settings;
        private ArrayList m_printers;
        private bool m_canRenderImage;
        private Graphics.EnumerateMetafileProc m_delegate;
        private Stream m_currentPageStream;
        private Metafile m_metafile;
        private int m_numberOfPages;
        private int m_currentPrintingPage;
        private int m_lastPrintingPage;
        private RenderedOutputFile[] m_files;
        StringBuilder m_sb = new StringBuilder();

        public PrinterDeliveryProvider()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        /// <summary>
        /// LocalizedName property
        /// </summary>
        string IExtension.LocalizedName
        {
            // TODO: Ideally your provider would support multiple, localized names
            get
            {
                return GetCultureLocalizedName();
            }
        }

        /// <summary>
        /// Process XML data stored in the configuration file
        /// </summary>
        /// <param name="configuration">The XML string from the configuration file that contains extension configuration data.</param>
        void IExtension.SetConfiguration(string configuration)
        {
            // Create the document and load the Configuration element    
            XmlDocument doc = new XmlDocument();

            try
            {
                //  Creates the XML reader from the stream 
                //  and moves it to the correct node
                using (StringReader m_srdr = new StringReader(configuration))
                {
                    using (XmlReader m_xrdr = XmlReader.Create(m_srdr))
                    {
                        m_xrdr.MoveToContent();
                        doc.LoadXml(m_xrdr.ReadOuterXml());
                    }
                }

                // For each printer in the configuration data, add it to the list of printers
                m_printers = new ArrayList();
                if (doc.DocumentElement.Name == "Printers")
                {
                    foreach (XmlNode child in doc.DocumentElement.ChildNodes)
                    {
                        m_printers.Add(child.InnerText);
                    }
                }
            }

            catch (Exception ex)
            {
                throw new Exception("Failed to retrieve configuration data: " + ex.Message);
            }
        }

        /// <summary>
        /// Used to determine whether a given set of delivery extension settings are valid.
        /// </summary>
        /// <param name="settings">An array of Setting[] objects containing extension settings supplied by a client.</param>
        /// <returns>The validated settings</returns>
        Setting[] IDeliveryExtension.ValidateUserData(Setting[] settings)
        {
            // TODO: Validate the given settings to ensure that they contain
            // all necessary information to process a notification.
            // For printer delivery, you can validate the printer and/or limit
            // the width and height settings to valid values
            foreach (Setting setting in settings)
            {
                // If the setting is based on a field, no check is required
                if (string.IsNullOrEmpty(setting.Field) == true)
                {
                    switch (setting.Name)
                    {
                        case (SubscriptionData.PRINTER):
                            if (!IsValidPrinter(setting.Value))
                                setting.Error = String.Format(
                                    System.Globalization.CultureInfo.InvariantCulture,
                                    "Printer {0} is not currently installed", setting.Value);
                            break;
                        case (SubscriptionData.PAGEHEIGHT):
                            // TODO: validate page height
                            break;
                        case (SubscriptionData.PAGEWIDTH):
                            // TODO: validate page width
                            break;
                        case (SubscriptionData.DPIX):
                            // TODO: validate x resolution
                            break;
                        case (SubscriptionData.DPIY):
                            // TODO: validate y resolution
                            break;
                        default:
                            // This is an unknown setting element
                            setting.Error = String.Format(
                                System.Globalization.CultureInfo.InvariantCulture,
                                "Unknown setting {0}", setting.Name);
                            break;
                    }
                }
            }

            return settings;
        }

        /// <summary>
        /// Required for implementation. Not used by this provider
        /// </summary>
        bool IDeliveryExtension.IsPrivilegedUser
        {
            //get
            //{
            //    return m_isPrivilegedUser;
            //}
            set
            {
                m_isPrivilegedUser = value;
            }
        }

        /// <summary>
        /// Returns the valid settings for the printer delivery extension
        /// </summary>
        Setting[] IDeliveryExtension.ExtensionSettings
        {
            get
            {
                if (m_settings == null)
                {
                    m_settings = new Setting[5];

                    m_settings[0] = new Setting();
                    m_settings[0].Name = SubscriptionData.PRINTER;
                    m_settings[0].ReadOnly = false;
                    m_settings[0].Required = true;

                    // Add the printer names that were retrieved from the   
                    // configuration file to the set of valid values for
                    // the setting
                    foreach (string printer in m_printers)
                    {
                        m_settings[0].AddValidValue(printer.ToString(), printer.ToString());
                    }

                    // Setting for page height
                    m_settings[1] = new Setting();
                    m_settings[1].Name = SubscriptionData.PAGEHEIGHT;
                    m_settings[1].ReadOnly = false;
                    m_settings[1].Required = true;
                    m_settings[1].Value = "11";

                    // Setting for page width
                    m_settings[2] = new Setting();
                    m_settings[2].Name = SubscriptionData.PAGEWIDTH;
                    m_settings[2].ReadOnly = false;
                    m_settings[2].Required = true;
                    m_settings[2].Value = "8.5";

                    // Setting for page x resolution
                    m_settings[3] = new Setting();
                    m_settings[3].Name = SubscriptionData.DPIX;
                    m_settings[3].ReadOnly = false;
                    m_settings[3].Required = true;
                    m_settings[3].Value = "96";

                    // Setting for page y resolution
                    m_settings[4] = new Setting();
                    m_settings[5].Name = SubscriptionData.DPIY;
                    m_settings[6].ReadOnly = false;
                    m_settings[7].Required = true;
                    m_settings[8].Value = "96";
                
                }

                return m_settings;
            }
        }

        /// <summary>
        /// Gets information about the report server that the delivery extension requires 
        /// in order to perform deliveries. This property contains the names of rendering 
        /// extensions supported by the server. Depending on the functionality of your 
        /// delivery extension, you should limit the available rendering extensions to only 
        /// those supported by your delivery mechanism
        /// </summary>
        IDeliveryReportServerInformation IDeliveryExtension.ReportServerInformation
        {
            set
            {
                m_reportServerInfo = value;
            }
        }

        /// <summary>
        /// Utility method that is used to evaluate whether the 
        /// IMAGE rendering extension is available
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2122:DoNotIndirectlyExposeMethodsWithLinkDemands")]
        private bool CanRenderImage
        {
            get
            {
                // Loop through the available rendering extensions and 
                // ensure that the rendering format IMAGE is available
                foreach (Microsoft.ReportingServices.Interfaces.Extension extension
                    in m_reportServerInfo.RenderingExtension)
                {
                    if (extension.Name == "IMAGE")
                    {
                        m_canRenderImage = true;
                        break;
                    }
                }

                return m_canRenderImage;
            }
        }

        /// <summary>
        /// Delivers the report notification to a user based on the contents of the notification.
        /// </summary>
        /// <param name="notification">A Notification object containing information required by the 
        /// delivery extension to deliver a report.</param>
        /// <returns>A Boolean value indicating whether or not to retry the delivery.</returns>
        bool IDeliveryExtension.Deliver(Notification notification)
        {
            bool success = false;

            // Set the status of the notification to pending
            notification.Status = "Processing...";

            try
            {
                // Build user data 
                Setting[] userSettings = notification.UserData;
                SubscriptionData subscriptionData = new SubscriptionData();
                subscriptionData.FromSettings(userSettings);

                // Print the report
                PrintReport(notification, subscriptionData);

                // If delivery is successful return true
                success = true;

            }
            catch (Exception ex)
            {
                // Set the status of the notification if an error occurs
                notification.Status = "Error: " + ex.Message;
                success = false;
            }

            finally
            {
                // Finally, save the notification information
                notification.Save();
            }

            return success;
        }

        /// <summary>
        /// Renders a report and sends it to the printer.
        /// </summary>
        /// <param name="notification">The report notification to use for delivery.</param>
        /// <param name="data">The user data to use for delivery.</param>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:DoNotPassLiteralsAsLocalizedParameters", MessageId = "System.Exception.#ctor(System.String)"), System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:DoNotPassLiteralsAsLocalizedParameters", MessageId = "System.Exception.#ctor(System.String)"), System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes"), System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        private void PrintReport(Notification notification, SubscriptionData data)
        {
            try
            {
                // Debug
                m_sb.Append(System.DateTime.Now + Environment.NewLine);

                // Make sure that the format for rendering to a printer is available 
                if (!CanRenderImage)
                {
                    throw new Exception("The IMAGE rendering extension needed for printer delivery is not available.");
                }

                // Correctly format device info strings
                string pageWidth = data.pageWidth + "in";
                string pageHeight = data.pageHeight + "in";
                
                string deviceInfo;

                // Render each page in the report and store it in a
                // Stream array until a blank page is rendered. This
                // marks the end of the report

                deviceInfo = String.Format(
                        System.Globalization.CultureInfo.InvariantCulture,
                        @"<DeviceInfo><OutputFormat>{0}</OutputFormat><PageHeight>{1}</PageHeight><PageWidth>{2}</PageWidth><DpiX>96</DpiX><DpiY>96</DpiY></DeviceInfo>",
                        "emf", pageHeight, pageWidth);

                // Render report
                m_files = notification.Report.Render("IMAGE", deviceInfo);
                m_numberOfPages = m_files.Length;

                // Write render information to the log entry
                m_sb.Append("Number of pages: " + m_numberOfPages + Environment.NewLine);
                                
                // Configure printer settings
                PrinterSettings printerSettings = new PrinterSettings();
                printerSettings.PrintRange = PrintRange.SomePages;
                printerSettings.FromPage = 1;
                printerSettings.ToPage = m_numberOfPages;
                printerSettings.DefaultPageSettings.Landscape = true; 
                
                /* The name of the printer comes from the subscription data.
                 * Validate the printer against known, installed printers on the server.
                 * This is done in the ValidateUserData method as well as now at deliver
                 * time to ensure that a printer has not been uninstalled since the time
                 * The subscription was created and the time the delivery takes place
				 * 
				 * Note: In production code, you would want to make this check earlier
				 * because the Render operation can take significant amounts of time 
				 * and memory.  It is presented here in the sample flow for clarity. 
                */
        //        if (!IsValidPrinter(data.Printer))
        //            throw new Exception("The printer " + data.Printer + " is not currently installed on the server.");

                printerSettings.PrinterName = data.Printer;

                // Create print document and set printer settings
                PrintDocument pd = new PrintDocument();
                m_currentPrintingPage = 1;
                m_lastPrintingPage = m_numberOfPages;
                pd.PrinterSettings = printerSettings;

                //pd.DefaultPageSettings.Landscape = true;

                // Print pages
                pd.PrintPage += new PrintPageEventHandler(this.pd_PrintPage);
                pd.Print();

                // Set the status if the report pages were successfully sent to the printer
                notification.Status = String.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                    "Report printed to {0}", data.Printer);
            }
            catch (Exception ex)
            {
                // Set the status of the notification if an error occurs
                notification.Status = ex.Message;
                // Write exception information to a logfile
                m_sb.Append(ex.Message + ": " + ex.StackTrace + Environment.NewLine);
            }
            finally
            {
                WriteLog(m_sb);
                m_sb = null;
            }
        }

        private void pd_PrintPage(object sender, PrintPageEventArgs ev)
        {
            ev.HasMorePages = false;

            //Color c=new Color();

            if (m_currentPrintingPage <= m_lastPrintingPage && MoveToPage(m_currentPrintingPage))
            {
                ev.PageSettings.Landscape = true;
                ev.PageSettings.Margins = new Margins(0, 0, 0, 0);
                
                // Draw the page
                ev.Graphics.ScaleTransform((float)0.8, (float)0.8);

                ReportDrawPage(ev.Graphics);

              //  ev.Graphics.Clear(c);
                // If the next page is less than or equal to the last page, 
                // print another page.
                if (++m_currentPrintingPage <= m_lastPrintingPage)
                    ev.HasMorePages = true;
            }
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Performance", "CA1804:RemoveUnusedLocals", MessageId = "rof")]
        private bool MoveToPage(Int32 page)
        {
            // Check to make sure that the current page exists in
            // the array list
            if (null == m_files[page - 1])
            {
                return false;
            }

            // Set current page stream equal to the rendered page
            RenderedOutputFile rof = m_files[page - 1];
   
            m_currentPageStream = (Stream)(m_files[page - 1].Data);

            // Set its postion to start.
            m_currentPageStream.Position = 0;

            // Initialize the metafile
            if (m_metafile != null)
            {
                m_metafile.Dispose();
                m_metafile = null;
            }

            // Load the metafile image for this page
            m_metafile = new Metafile(m_currentPageStream);
            return true;
        }

        // Method to draw the current emf memory stream 
        private void ReportDrawPage(Graphics graphics)
        {
            if (m_currentPageStream == null || m_currentPageStream.Length == 0 || m_metafile == null)
                return;
            // Set metafile delegate.
            m_delegate = new Graphics.EnumerateMetafileProc(MetafileCallback);
            // Draw in the rectangle
            Point destPoint = new Point(0, 0);
            graphics.EnumerateMetafile(m_metafile, destPoint, m_delegate);
            // Clean up
            m_delegate = null;

        }

        private bool MetafileCallback(
           EmfPlusRecordType recordType,
           int flags,
           int dataSize,
           IntPtr data,
           PlayRecordCallback callbackData)
        {
            byte[] dataArray = null;
            // Process around unmanaged code
            if (data != IntPtr.Zero)
            {
                // Copy the unmanaged record to a managed byte buffer 
                // that can be used by PlayRecord
                dataArray = new byte[dataSize];
                Marshal.Copy(data, dataArray, 0, dataSize);
            }
            // Playback the record      
            m_metafile.PlayRecord(recordType, flags, dataSize, dataArray);

            return true;
        }

        // Used for debugging. Currently writes a log file to
        // Windows system32 directory.
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1062:ValidateArgumentsOfPublicMethods"), System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:DoNotPassLiteralsAsLocalizedParameters", MessageId = "System.IO.IOException.#ctor(System.String)")]
        public static void WriteLog(StringBuilder sb)
        {
            try
            {
                FileStream fs = new FileStream("PrinterDeliveryLog.txt", FileMode.Append,
                   FileAccess.Write);
                StreamWriter writer = new StreamWriter(fs);
                writer.Write(sb.ToString());
                writer.Flush();
                writer.Close();
            }

            catch (Exception ex)
            {
                throw new IOException("Error writing to log file: " + ex.Message);
            }
        }

        // Method to check the selected printer with the list of currently
        // installed printers on the server
        public static bool IsValidPrinter(string printer)
        {
            foreach (string p in PrinterSettings.InstalledPrinters)
            {
                if (p == printer)
                {
                    return true;
                }
            }

            return false;
        }

        // Method to evaluate the culture info of the current thread and return
        // a localized name. This extension name supports German and English names
        // Ideally it would also support a localized Description propery in the UI
        // as well as localized setting labels.
        private static string GetCultureLocalizedName()
        {
            switch (CultureInfo.CurrentCulture.Name)
            {
                // TODO: Add more languages or ideally use resource files
                case "de-AT":
                case "de-DE":
                case "de-LI":
                case "de-LU":
                case "de-CH":
                    return "Beispiel für Direktsendung zum Drucker";
                default:
                    return "Printer Delivery";
            }
        }
    }
}

