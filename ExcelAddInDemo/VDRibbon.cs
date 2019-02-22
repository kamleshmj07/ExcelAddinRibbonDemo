using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddInDemo
{
    public partial class VDRibbon
    {
        private void VDRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            this.ddlTable_LoadData();
        }

        private void ddlTable_LoadData()
        {
            string[] itemCollection = new string[3] { "Tbl01", "Tbl02", "Tbl03" };

            AddDefaultDropDownItem(ddlTable, "Select Table");

            foreach (var item in itemCollection)
            {
                RibbonDropDownItem ribbonItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                ribbonItem.Label = item;
                ddlTable.Items.Add(ribbonItem);
            }
        }

        private void AddDefaultDropDownItem(RibbonDropDown ribbonDropDown, string defaultLabel)
        {
            ribbonDropDown.Items.Clear();
            var defaultItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            defaultItem.Label = defaultLabel;
            ribbonDropDown.Items.Add(defaultItem);
        }
    }

}
