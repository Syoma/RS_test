/*=====================================================================

  File:      PrinterDeliveryUIProvider.cs
  
  Summary:   Represents delivery provider user interface that can be 
             used to creating printer delivery subscriptions in 
             Report Manager.

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
using System.Web;
using System.Text;
using System.Web.UI;
using System.Collections;
using System.Globalization;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using Microsoft.ReportingServices.Interfaces;
using System.Xml;


namespace Unact.ReportingServices.PrinterDelivery
{
    // PrinterDeliveryUIProvider implements a UserControl that plugs into the Reporting Services Report Manager application.
    public class PrinterDeliveryUIProvider : System.Web.UI.WebControls.WebControl, ISubscriptionBaseUIUserControl
    {
        // Variables used to store information about the subscription being created
        private double m_pagewidth = 8.5;
        private double m_pageheight = 11;
        private ArrayList m_printers;

        // Labels for UI controls
        internal const string PRINTERLABEL = "Select a printer:";
        internal const string PAGEWIDTHLABEL = "Enter a page width:";
        internal const string PAGEHEIGHTLABEL = "Enter a page height:";

        // IDs used to refer to controls on the UI page.
        internal const string PRINTERCONTROLID = "PRINTERSELECTLIST";
        internal const string PAGEWIDTHCONTROLID = "PAGEWIDTHTEXTBOX";
        internal const string PAGEHEIGHTCONTROLID = "PAGEHEIGHTTEXTBOX";

        // Strings to enable validation to occur (client side validation needs JavaScript)
        internal const string LANGUAGEATTRIBUTE = "language";
        internal const string SCRIPTDEFAULTLANGUAGE = "Javascript";
        internal const string SCRIPTTYPEATTRIBUTE = "type";
        internal const string SCRIPTDEFAULTTYPE = "text/Javascript";
        internal const string SCRIPTTAG = "script";
        //internal const string ONCLICK = "onclick";

        // Used to keep track of whether we have values specified by the user
        private bool m_hasUserData;

        #region Controls
        // Controls used on the UI page
        private LiteralControl m_validatorScript = new LiteralControl();

        // HTML table variables
        private HtmlTable m_outerTable = new HtmlTable();
        private HtmlTableRow m_currentRow;
        private HtmlTableCell m_currentCell;
        private HtmlGenericControl m_pageLevelScript = new HtmlGenericControl(SCRIPTTAG);

        // Labels for controls
        private Label m_printerDropDownLabel = new Label();
        private Label m_pageHeightLabel = new Label();
        private Label m_pageWidthLabel = new Label();

        // Control types
        private DropDownList m_printersDropDownList = new DropDownList();
        private TextBox m_pageHeightTextBox = new TextBox();
        private TextBox m_pageWidthTextBox = new TextBox();

        // Placeholders for validators
        private PlaceHolder m_invalidPrinterName = new PlaceHolder();
        private PlaceHolder m_invalidPageHeight = new PlaceHolder();
        private PlaceHolder m_invalidPageWidth = new PlaceHolder();

        // Field validators for UI
        private RequiredFieldValidator m_printerRequired = new RequiredFieldValidator();
        private RequiredFieldValidator m_pageHeightRequired = new RequiredFieldValidator();
        private RequiredFieldValidator m_pageWidthRequired = new RequiredFieldValidator();

        #endregion

        // Provider constructor
        public PrinterDeliveryUIProvider()
        {

            this.Init += new EventHandler(Control_Init);
            this.Load += new EventHandler(Control_Load);
            this.PreRender += new EventHandler(Control_PreRender);
        }

        #region Event Handlers

        private void Control_PreRender(object sender, EventArgs args)
        {
            // Use this event to enable/disable controls based on the 
            // user's selection before the control is rendered.
            // The PrinterDeliverySample does not use this event handler.
        }


        /// <summary>Perform all step needed when the control has been loaded</summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void Control_Load(object sender, EventArgs args)
        {
            if (!Page.IsPostBack)
            {
                //if you have non-required Extension settings, initialize the values of the controls here
            }
        }


        /// <summary>
        /// Initialize the control
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void Control_Init(object sender, EventArgs args)
        {
            Controls.Add(m_pageLevelScript);

            // Create page level script
            m_pageLevelScript.Attributes.Add(LANGUAGEATTRIBUTE, SCRIPTDEFAULTLANGUAGE);
            m_pageLevelScript.Attributes.Add(SCRIPTTYPEATTRIBUTE, SCRIPTDEFAULTTYPE);
            Controls.Add(m_validatorScript);
            SetValidatorScript();

            // Build a table row for selecting a printer from the printersDropDownList
            #region Printer DropDown Row
            m_currentRow = new HtmlTableRow();
            m_outerTable.Rows.Add(m_currentRow);

            // Create the first cell that contains the label for the printer drop down list
            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "10%";

            //TODO: Specify label here
            this.m_printerDropDownLabel.Text = HttpUtility.HtmlEncode(PRINTERLABEL);
            m_currentCell.Controls.Add(this.m_printerDropDownLabel);

            // Add the cell that contains the drop down list of printer names
            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "25%";
            //TODO: Add control here
            this.m_printersDropDownList.SelectedIndex = 0;
            //TODO: Include list of printers here...
            //TODO: Define default values here (merge UserData and ServerInfo)
            if (!this.m_hasUserData)
            {
                Setting printers = null;
                this.m_printers = new ArrayList();

                // Grab the settings from the server
                foreach (Setting setting in this.m_rsInformation.ServerSettings)
                {
                    if (setting.Name.Equals(SubscriptionData.PRINTER))
                    {
                        printers = setting;
                    }
                }

                if (printers != null)
                {
                    foreach (ValidValue validValues in printers.ValidValues)
                    {
                        ListItem li = new ListItem(validValues.Label, validValues.Value);
                        if (validValues.Value.Equals(printers.Value))
                        {
                            li.Selected = true;
                        }
                        this.m_printersDropDownList.Items.Add(li);
                        this.m_printers.Add(validValues.Value);
                    }
                }
            }

            this.m_printersDropDownList.ID = PRINTERCONTROLID;
            m_currentCell.Controls.Add(this.m_printersDropDownList);

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "40%";

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "100%";

            m_currentCell.Controls.Add(this.m_invalidPrinterName);
            this.m_printerRequired.Display = ValidatorDisplay.Dynamic;
            this.m_printerRequired.ControlToValidate = PRINTERCONTROLID;
            String printerRequiredError = "Specify one of the valid values for the printer setting.";
            this.m_printerRequired.Controls.Add(ErrorMessage(printerRequiredError, true));
            this.m_invalidPrinterName.Controls.Add(this.m_printerRequired);

            PrinterInListOfPrintersValidator pilValidator = new PrinterInListOfPrintersValidator();
            pilValidator.ValidValues = this.m_printers;
            pilValidator.Display = ValidatorDisplay.Dynamic;
            pilValidator.ControlToValidate = PRINTERCONTROLID;
            pilValidator.Controls.Add(ErrorMessage("The specified value is not among the valid values for this setting.  Specify another value.", true));
            this.m_invalidPrinterName.Controls.Add(pilValidator);

            #endregion

            // Build a table row for entering a page width
            #region Page Width Textbox Row
            m_currentRow = new HtmlTableRow();
            m_outerTable.Rows.Add(m_currentRow);

            // Add cell for page width label
            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "10%";

            // Add label
            this.m_pageWidthLabel.Text = HttpUtility.HtmlEncode(PAGEWIDTHLABEL);
            m_currentCell.Controls.Add(this.m_pageWidthLabel);

            // Add text box for entering value
            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "25%";

            this.m_pageWidthTextBox.Text = System.Convert.ToString(
                this.m_pagewidth,
                System.Globalization.CultureInfo.InvariantCulture);

            m_pageWidthTextBox.ID = PAGEWIDTHCONTROLID;
            m_pageWidthTextBox.Style.Add("font-family", "Verdana, Sans-Serif");
            m_pageWidthTextBox.Style.Add("font-size", "x-small");


            this.m_pageWidthTextBox.ID = PAGEWIDTHCONTROLID;
            m_currentCell.Controls.Add(this.m_pageWidthTextBox);

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "40%";

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "100%";

            // Add validator here
            m_currentCell.Controls.Add(this.m_invalidPageWidth);
            this.m_pageWidthRequired.Display = ValidatorDisplay.Dynamic;
            this.m_pageWidthRequired.ControlToValidate = PAGEWIDTHCONTROLID;
            String pageWidthRequiredError = "Specify a value for the page width.";
            this.m_pageWidthRequired.Controls.Add(ErrorMessage(pageWidthRequiredError, true));
            this.m_invalidPageWidth.Controls.Add(this.m_pageWidthRequired);

            ValueGreaterThanZeroValidator vgtzValidator1 = new ValueGreaterThanZeroValidator();
            vgtzValidator1.Display = ValidatorDisplay.Dynamic;
            vgtzValidator1.ControlToValidate = PAGEWIDTHCONTROLID;
            vgtzValidator1.Controls.Add(ErrorMessage("The specified must be greater than zero (0).", true));
            this.m_invalidPageWidth.Controls.Add(vgtzValidator1);

            #endregion

            // Build a table row for Entering a page height 
            #region Page Height Textbox Row
            m_currentRow = new HtmlTableRow();
            m_outerTable.Rows.Add(m_currentRow);

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "10%";

            this.m_pageHeightLabel.Text = HttpUtility.HtmlEncode(PAGEHEIGHTLABEL);
            m_currentCell.Controls.Add(this.m_pageHeightLabel);

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "25%";

            this.m_pageHeightTextBox.Text = System.Convert.ToString(this.m_pageheight,
                System.Globalization.CultureInfo.InvariantCulture);

            m_pageHeightTextBox.ID = PAGEHEIGHTCONTROLID;
            m_pageHeightTextBox.Style.Add("font-family", "Verdana, Sans-Serif");
            m_pageHeightTextBox.Style.Add("font-size", "x-small");


            this.m_pageHeightTextBox.ID = PAGEHEIGHTCONTROLID;
            m_currentCell.Controls.Add(this.m_pageHeightTextBox);

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "40%";

            m_currentCell = new HtmlTableCell();
            m_currentCell.NoWrap = true;
            m_currentRow.Cells.Add(m_currentCell);
            m_currentCell.Width = "100%";

            m_currentCell.Controls.Add(this.m_invalidPageHeight);
            this.m_pageHeightRequired.Display = ValidatorDisplay.Dynamic;
            this.m_pageHeightRequired.ControlToValidate = PAGEHEIGHTCONTROLID;
            String pageHeightRequired = "Specify a value for the page height.";
            this.m_pageHeightRequired.Controls.Add(ErrorMessage(pageHeightRequired, true));
            this.m_invalidPageHeight.Controls.Add(this.m_pageHeightRequired);

            ValueGreaterThanZeroValidator vgtzValidator2 = new ValueGreaterThanZeroValidator();
            vgtzValidator2.Display = ValidatorDisplay.Dynamic;
            vgtzValidator2.ControlToValidate = PAGEHEIGHTCONTROLID;
            vgtzValidator2.Controls.Add(ErrorMessage("The specified must be greater than zero (0).", true));
            this.m_invalidPageHeight.Controls.Add(vgtzValidator2);

            #endregion

            m_outerTable.Attributes.Add("class", "msrs-normal");
            m_outerTable.CellPadding = 0;
            m_outerTable.CellSpacing = 0;
            m_outerTable.Width = "100%";
            Controls.Add(m_outerTable);

        }

        #region VALIDATORSCRIPTFUNCTION
        private const string VALIDATORSCRIPTFUNCTION =
           @"<script language='Javascript' type='text/Javascript'>
                function ValidateValueGreaterThanZero(source, args)
                {
                   var obj = document.all(source.controltovalidate);
   
               if ((obj.value != null) && (!isNaN(number(obj.value))) )
                    {
                  if (number(obj.value) <=0)
                  {
                     args.IsValid = false;
                  }
                    }

                  args.IsValid = true;
                }
                </script>
                ";
        #endregion
        private void SetValidatorScript()
        {
            m_validatorScript.Text = VALIDATORSCRIPTFUNCTION;
        }

        protected Control ErrorMessage(string error, bool noWrap)
        {
            string imgUrl = Page.Request.ApplicationPath + "/images/line_err1.gif";
            HtmlImage htmlImg = new HtmlImage();
            htmlImg.Src = imgUrl;
            htmlImg.Alt = "Value specified contains an error.";

            HtmlTable tbl = new HtmlTable();
            tbl.Rows.Add(new HtmlTableRow());
            HtmlTableCell cell = new HtmlTableCell();
            cell.VAlign = "middle";
            //Note: you can reuse the Report Manger style sheet.
            cell.Attributes.Add("class", "msrs-validationerror");
            cell.Controls.Add(htmlImg);
            tbl.Rows[0].Cells.Add(cell);
            cell = new HtmlTableCell();
            cell.NoWrap = noWrap;
            cell.VAlign = "middle";
            cell.Attributes.Add("class", "msrs-validationerror");
            cell.Controls.Add(new LiteralControl(HttpUtility.HtmlEncode(error)));
            tbl.Rows[0].Cells.Add(cell);
            return tbl;
        }


        #endregion


        #region IExtension methods
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase")]
        public String LocalizedName
        {
            get
            {
                return "Printer Delivery Sample";
            }
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes"), System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase")]
        public void SetConfiguration(String configuration)
        {
            CultureInfo info = System.Threading.Thread.CurrentThread.CurrentCulture;
            try
            {
                //no configuration data for the printer delivery UI.
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, null);
            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = info;
            }
        }

        #endregion

        #region ISubscriptionBaseUIUserControl methods

        private bool m_isPrivilegedUser;

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase")]
        public bool IsPrivilegedUser
        {
            get
            {
                return m_isPrivilegedUser;
            }
            set
            {
                m_isPrivilegedUser = value;
            }
        }

        // Validate that all selected information is correct
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase")]
        public bool Validate()
        {
            // Nothing additional to validate
            return true;
        }

        // Get and Set the user data
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase"), System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase")]
        public Setting[] UserData
        {
            get
            {
                SubscriptionData data = new SubscriptionData();
                data.Printer = this.m_printersDropDownList.SelectedValue;
                data.pageHeight = System.Convert.ToDouble(this.m_pageHeightTextBox.Text,
                    System.Globalization.CultureInfo.InvariantCulture);
                data.pageWidth = System.Convert.ToDouble(this.m_pageWidthTextBox.Text,
                    System.Globalization.CultureInfo.InvariantCulture);

                return data.ToSettingArray();
            }

            set
            {
                this.m_hasUserData = true;

                SubscriptionData data = new SubscriptionData();
                data.FromSettings(value);

                this.m_pagewidth = System.Convert.ToDouble(data.pageWidth);
                this.m_pageWidthTextBox.Text = System.Convert.ToString(this.m_pagewidth,
                    System.Globalization.CultureInfo.InvariantCulture);

                this.m_pageheight = System.Convert.ToDouble(data.pageHeight);
                this.m_pageHeightTextBox.Text = System.Convert.ToString(this.m_pageheight,
                    System.Globalization.CultureInfo.InvariantCulture);


                bool found = false;
                Setting[] serverSettings = m_rsInformation.ServerSettings;
                Setting printers = null;
                foreach (Setting s in serverSettings)
                {
                    if (s.Name.Equals(SubscriptionData.PRINTER))
                    {
                        printers = s;
                    }
                }

                this.m_printersDropDownList.Items.Clear();
                this.m_printers = new ArrayList();

                foreach (ValidValue vv in printers.ValidValues)
                {
                    this.m_printers.Add(vv.Value);

                    ListItem li = new ListItem(vv.Label, vv.Value);

                    if (!found && String.Compare(vv.Value, data.Printer, true, CultureInfo.InvariantCulture) == 0)
                    {
                        found = true;
                        li.Selected = true;
                    }

                    this.m_printersDropDownList.Items.Add(li);
                }

                // If the printer was not found in the list of possible printers it
                // is a printer that is no longer allowed.  We still allow  
                // the printer and display it as selected         
                if (!found)
                {
                    ListItem li = new ListItem(data.Printer, data.Printer);
                    li.Selected = true;
                    this.m_printersDropDownList.Items.Add(li);
                }
            }
        }

        // Get the description that displays for the subscription
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase")]
        public String Description
        {
            get
            {
                return "Print report to " + this.m_printersDropDownList.SelectedItem.Text + ".";
            }
        }

        private IDeliveryReportServerInformation m_rsInformation;
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2123:OverrideLinkDemandsShouldBeIdenticalToBase")]
        public IDeliveryReportServerInformation ReportServerInformation
        {
            set
            {
                m_rsInformation = value;
            }
            get
            {
                return m_rsInformation;
            }
        }

        #endregion


    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1501:AvoidExcessiveInheritance")]
    internal class PrinterInListOfPrintersValidator : CustomValidator
    {
        public PrinterInListOfPrintersValidator()
            : base()
        {
            ServerValidate += new ServerValidateEventHandler(Validate_Server);
        }

        private ArrayList m_validValues;

        public ArrayList ValidValues
        {
            set
            {
                m_validValues = value;
            }
        }

        // Validate function which ensures that the drop down has a valid value selected
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Globalization", "CA1303:DoNotPassLiteralsAsLocalizedParameters", MessageId = "System.Exception.#ctor(System.String,System.Exception)"), System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2201:DoNotRaiseReservedExceptionTypes")]
        private void Validate_Server(object source, ServerValidateEventArgs args)
        {
            DropDownList ddl;

            try
            {
                ddl = (DropDownList)FindControl(((CustomValidator)source).ControlToValidate);

            }
            catch (Exception ex)
            {
                throw new Exception("Error locating dropdownlist.", ex);
            }

            if (m_validValues == null)
            {
                throw new Exception("ValidValues is null.");
            }

            foreach (string curValue in m_validValues)
            {
                if (ddl.SelectedItem.Value.Equals(curValue))
                {
                    args.IsValid = true;
                    return;
                }
            }

            args.IsValid = false;
            return;
        }
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Maintainability", "CA1501:AvoidExcessiveInheritance")]
    internal class ValueGreaterThanZeroValidator : CustomValidator
    {
        public ValueGreaterThanZeroValidator()
            : base()
        {
            ServerValidate += new ServerValidateEventHandler(Validate_Server);

        }

        private void Validate_Server(object source, ServerValidateEventArgs args)
        {
            TextBox tb = (TextBox)FindControl(((CustomValidator)source).ControlToValidate);
            double d = System.Convert.ToDouble(tb.Text,
                System.Globalization.CultureInfo.InvariantCulture);
            if (d <= 0)
            {
                args.IsValid = false;
            }
            else
            {
                args.IsValid = true;
            }
        }
    }
}


