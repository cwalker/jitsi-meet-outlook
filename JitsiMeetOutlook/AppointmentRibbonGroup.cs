using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace JitsiMeetOutlook
{
    public partial class AppointmentRibbonGroup
    {

        private void AppointmentRibbonGroup_Load(object sender, RibbonUIEventArgs e)
        {
            if (Properties.Settings.Default.disableCustomRoomId)
            {
                fieldRoomID.Visible = false;
                buttonRandomRoomID.Visible = false;
            }

            initialise();
        }

        private void buttonDialogLauncher_Click(object sender, RibbonControlEventArgs e)
        {

            FormSettings settingsWindow = new FormSettings();
            settingsWindow.Show();
        }

        private void buttonCustomiseJitsiMeeting_Click(object sender, RibbonControlEventArgs e)
        {
            randomiseRoomId();
        }

        private void buttonRequireDisplayName_Click(object sender, RibbonControlEventArgs e)
        {
            toggleRequireName();
        }

        private void buttonStartWithAudioMuted_Click(object sender, RibbonControlEventArgs e)
        {
            toggleMuteOnStart();
        }

        private void RoomID_TextChanged(object sender, RibbonControlEventArgs e)
        {
            _ = setRoomId(fieldRoomID.Text);
        }

        private void buttonStartWithVideoMuted_Click(object sender, RibbonControlEventArgs e)
        {
            toggleVideoOnStart();
        }

        private void buttonExtUrl_CheckedChanged(object sender, RibbonControlEventArgs e)
        {
            if (buttonExtUrl.Checked)
            {
                Document wordDocument = appointmentItem.GetInspector.WordEditor as Document;
                wordDocument.Select();
                var endSel = wordDocument.Application.Selection;
                object missing = System.Reflection.Missing.Value;
                var extlink = JitsiUrl.getExtUrlBase() + roomId;
                endSel.InsertAfter("\n");
                endSel.MoveDown(Word.WdUnits.wdLine);
                endSel.EndKey(Word.WdUnits.wdLine);
                var hyperLinkExt = wordDocument.Hyperlinks.Add(endSel.Range, extlink, ref missing, ref missing, extlink, ref missing);
                hyperLinkExt.Range.Font.Size = 10;
                hyperLinkExt.Application.Options.CtrlClickHyperlinkToOpen = false;
                hyperLinkExt.TextToDisplay = Globals.ThisAddIn.getElementTranslation("appointmentItem", "textBodyMessageExt");
                endSel.EndKey(Word.WdUnits.wdLine);
                endSel.MoveDown(Word.WdUnits.wdLine);
            }
            else
            {
                Document wordDocument = appointmentItem.GetInspector.WordEditor as Document;
                wordDocument.Select();
                var endSel = wordDocument.Application.Selection;
                endSel.MoveDown(Word.WdUnits.wdLine, 1);
                endSel.Expand(Word.WdUnits.wdLine);
                endSel.Delete();
                endSel.EndKey(Word.WdUnits.wdLine);

            }
        }


        private void buttonNewJitsiMeeting_Click(object sender, RibbonControlEventArgs e)
        {
            addJitsiMeeting();
        }
    }
}
