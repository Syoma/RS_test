/*=====================================================================

  File:      SubscriptionData.cs
  
  Summary:   Represents a centralized store of the extension settings
             needed to delivery reports to a printer.

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
using System.Collections;

using Microsoft.ReportingServices.Interfaces;

namespace Unact.ReportingServices.PrinterDelivery
{
    internal class SubscriptionData
    {

        // Initalize variables to default values
        public string Printer = "";
        public double pageHeight = 11.0;
        public double pageWidth = 8.5;
        public double Dpi = 300;
        public string Orientation = "Landscape";


        // Initialize setting names
        internal const string PRINTER = "Принтер";
        internal const string PAGEHEIGHT = "Высота документа";
        internal const string PAGEWIDTH = "Ширина документа";
        internal const string DPI = "Разрешение документа";
        internal const string ORIENTATION = "Ориентация документа";

        public SubscriptionData()
        {
            // TODO: Add constructor code here
        }

        // Populate the object from an array of setting elements
        // No validation is done, it is assumed that the settings
        // contains all relevant information
        public void FromSettings(Setting[] settings)
        {
            foreach (Setting setting in settings)
            {
                switch (setting.Name)
                {
                    case (PRINTER):
                        Printer = setting.Value;
                        break;
                    case (PAGEHEIGHT):
                        this.pageHeight = System.Convert.ToDouble(setting.Value,
                            System.Globalization.CultureInfo.InvariantCulture);
                        break;
                    case (PAGEWIDTH):
                        this.pageWidth = System.Convert.ToDouble(setting.Value,
                            System.Globalization.CultureInfo.InvariantCulture);
                        break;
                    case (DPI):
                        this.Dpi = System.Convert.ToDouble(setting.Value,
                            System.Globalization.CultureInfo.InvariantCulture);
                        break;
                    case (ORIENTATION):
                        this.Orientation = setting.Value;
                        break;
                    default:
                        break;
                }
            }
        }

        // Creates an array of the settings
        public Setting[] ToSettingArray()
        {
            ArrayList list = new ArrayList();

            list.Add(CreateSetting(PRINTER, this.Printer ));
            list.Add(CreateSetting(PAGEHEIGHT, System.Convert.ToString(this.pageHeight,
                System.Globalization.CultureInfo.InvariantCulture)));
            list.Add(CreateSetting(PAGEWIDTH, System.Convert.ToString(this.pageWidth,
                System.Globalization.CultureInfo.InvariantCulture)));
            list.Add(CreateSetting(DPI, System.Convert.ToString(this.Dpi,
                System.Globalization.CultureInfo.InvariantCulture)));
            list.Add(CreateSetting(ORIENTATION, this.Orientation ));

            return list.ToArray(typeof(Setting)) as Setting[];
        }

        // Creates a single instance of a Setting
        private static Setting CreateSetting(string name, string val)
        {
            Setting setting = new Setting();
            setting.Name = name;
            setting.Value = val;

            return setting;
        }
    }
}

