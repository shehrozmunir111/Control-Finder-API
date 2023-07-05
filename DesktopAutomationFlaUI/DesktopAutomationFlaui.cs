using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;
using PS19.ATM.ReturnStatus;

namespace DesktopAutomationFlaUI
{
    public class DesktopAutomationFlaui
    {
        public static string CSVLogsPath { get; set; } = Path.Combine(Directory.GetCurrentDirectory() + @"\Logs.txt");
        public static string AppPath { get; set; } = @"C:\Program Files (x86)\Smart Wires\SmartInterface\SmartInterface.exe";
        public static List<string> resultChild { get; set; } = new List<string>();

        public string AutomationId { get; set; }
        public string Name { get; set; }
        public string ControlType { get; set; }
        public string Parent { get; set; }
        public string LocationX { get; set; }
        public string LocationY { get; set; }
        public string Width { get; set; }
        public string Height { get; set; }
        public string Result { get; set; }

        private FlaUI.Core.Application app;

        public DesktopAutomationFlaui()
        {
            var process = Process.Start(new ProcessStartInfo(DesktopAutomationFlaui.AppPath));
            app = new FlaUI.Core.Application(process, false);
            Result = $"Automation Id,Name,Control Type,LocationX,LocationY,Height,Width";
            resultChild.Add(Result);
        }

        public Status OpenSmartInterface()
        {
            Status status = new Status(false, "", 0);
            try
            {
                using (var automation = new UIA3Automation())
                {
                    status = SelectComboboxValueByAutomationId(automation, "cmbMode", "Configuration Mode");
                    if (status.ErrorOccurred) { status.ReturnedMessage = status.ReturnedMessage + " 'cmbMode'"; return status; }
                    status = SetEditBoxValueByAutomationId(automation, "txtUserName", "administrator");
                    if (status.ErrorOccurred) { status.ReturnedMessage = status.ReturnedMessage + " 'txtUserName'"; return status; }
                    status = SetEditBoxValueByAutomationId(automation, "txtPwd", "P@s@1@9#a#t#m");
                    if (status.ErrorOccurred) { status.ReturnedMessage = status.ReturnedMessage + " 'txtPwd'"; return status; }
                    status = SelectComboboxValueByAutomationId(automation, "cmbAuthenticationMod", "Local Authentication(SmartInterface)");
                    if (status.ErrorOccurred) { status.ReturnedMessage = status.ReturnedMessage + " 'cmbAuthenticationMod'"; return status; }
                    status = InvokeButtonById(automation, "btnLogin");
                    if (status.ErrorOccurred) { status.ReturnedMessage = status.ReturnedMessage + " 'btnLogin'"; return status; }
                }
            }
            catch (Exception ex)
            {
                status.ErrorOccurred = true;
                status.ReturnedMessage = "Failed To Login SmartInterface";
            }
            return status;
        }

        public Status SelectComboboxValueByAutomationId(UIA3Automation automation, string comboBoxId, string valueToSelect)
        {
            Status status = new Status(false, "", 0);
            try
            {
                var window = app.GetMainWindow(automation);
                Thread.Sleep(500);
                var combo = window.FindFirstDescendant(cf => cf.ByAutomationId(comboBoxId)).AsComboBox();
                Thread.Sleep(500);
                if (combo.IsReadOnly || combo.IsEditable)
                {
                    combo.Select(valueToSelect);
                }
            }
            catch (Exception ex)
            {
                status.ErrorOccurred = true;
                status.ReturnedMessage = "Control Not Found";
            }
            return status;
            //var selectedItem = combo.SelectedItem;
            //Assert.That(selectedItem, Is.Not.Null);
            //Assert.That(selectedItem.Text, Is.EqualTo(valueToSelect));
        }

        public Status SetEditBoxValueByAutomationId(UIA3Automation automation, string editBoxId, string valueToEnter)
        {
            Status status = new Status(false, "", 0);
            try
            {
                var window = app.GetMainWindow(automation);
                Thread.Sleep(300);
                var editBox = window.FindFirstDescendant(cf => cf.ByAutomationId(editBoxId)).AsTextBox();
                Thread.Sleep(300);
                editBox.Enter(valueToEnter);
            }
            catch (Exception ex)
            {
                status.ErrorOccurred = true;
                status.ReturnedMessage = "Control Not Found";
            }

            //var selectedItem = combo.SelectedItem;
            //Assert.That(selectedItem, Is.Not.Null);
            //Assert.That(selectedItem.Text, Is.EqualTo("Configuration Mode"));
            return status;
        }
        public Status InvokeButtonById(UIA3Automation automation, string editBoxId)
        {
            Status status = new Status(false, "", 0);
            try
            {
                var window = app.GetMainWindow(automation);
                var btn = window.FindFirstDescendant(cf => cf.ByAutomationId(editBoxId)).AsButton();
                btn.Click();
            }
            catch (Exception ex)
            {
                status.ErrorOccurred = true;
                status.ReturnedMessage = "Control Not Found";
            }

            //var selectedItem = combo.SelectedItem;
            //Assert.That(selectedItem, Is.Not.Null);
            //Assert.That(selectedItem.Text, Is.EqualTo("Configuration Mode"));
            return status;
        }

    }
}
