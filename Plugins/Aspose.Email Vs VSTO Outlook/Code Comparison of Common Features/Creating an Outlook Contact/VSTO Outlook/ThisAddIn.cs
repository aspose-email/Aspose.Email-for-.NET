using Microsoft.Office.Interop.Outlook;

namespace VSTO_Outlook
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Microsoft.Office.Interop.Outlook._Application OutlookObject = new Microsoft.Office.Interop.Outlook.Application();
            //Create a new Contact Item
            Microsoft.Office.Interop.Outlook.ContactItem contact = OutlookObject.CreateItem(
                                    Microsoft.Office.Interop.Outlook.OlItemType.olContactItem);

            //Set different properties of this Contact Item.
            contact.FirstName = "Mellissa";
            contact.LastName = "MacBeth";
            contact.JobTitle = "Account Representative";
            contact.CompanyName = "Contoso Ltd.";
            contact.OfficeLocation = "36/2529";
            contact.BusinessTelephoneNumber = "4255551212 x432";
            contact.BusinessAddressStreet = "1 Microsoft Way";
            contact.BusinessAddressCity = "Redmond";
            contact.BusinessAddressState = "WA";
            contact.BusinessAddressPostalCode = "98052";
            contact.BusinessAddressCountry = "United States of America";
            contact.Email1Address = "melissa@contoso.com";
            contact.Email1AddressType = "SMTP";
            contact.Email1DisplayName = "Melissa MacBeth (mellissa@contoso.com)";

            //Save the Contact to disc
            contact.SaveAs("OutlookContact.vcf", OlSaveAsType.olVCard);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
