using System.Windows.Forms;
using System.Text.RegularExpressions;


namespace FCHelper_v001
{
    class GetNonstandardFileData
    {
        //external variables

        public static string[] EmployerID;
        public static string EmployerName;
        public static string Requester;
        public static string ApprovingManager;
        public static string FilesPurpose;
        public static string DatasOrigin;
        public static string ClientsApproval;
        public static string FileCreatorsName;
        public static string FileCreatorsPhone;
        public static string FileCreatorsEmail;
        public static string TechnicalReason;
        public static string BusinessImpact;
        


        public void GetNonstandardFileDataMethod(string nonstandardReqFormFilePathName)
        {
            PrivateGetNonstandardFileDataMethod(nonstandardReqFormFilePathName);
        }

        private void PrivateGetNonstandardFileDataMethod(string nonstandardReqFormFilePathName)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Documents docs = app.Documents;
            Microsoft.Office.Interop.Word.Document doc = docs.Open(nonstandardReqFormFilePathName, ReadOnly: true);

            Microsoft.Office.Interop.Word.Table t1 = doc.Tables[1];
            Microsoft.Office.Interop.Word.Range r1 = t1.Range;
            Microsoft.Office.Interop.Word.Cells cells1 = r1.Cells;


            //convert tables to text to get rid of bullet points at the end of text fields
            Regex rg = new Regex("^[A-Za-z0-9/$ !|\\[]{}%&()]$");//regex filter on output text accepting characters alphanumeric and /

            string employerID = rg.Replace(cells1[3].Range.Text, "");
            string employerName = rg.Replace(cells1[5].Range.Text, "");
            string rquestr = rg.Replace(cells1[8].Range.Text, "");
            string appManagr = rg.Replace(cells1[10].Range.Text, "");
            string filesPurpose = rg.Replace(cells1[13].Range.Text, "");
            string datasOrigin = rg.Replace(cells1[15].Range.Text, "");
            string clientsApproval = rg.Replace(cells1[17].Range.Text, "");

            string fileCreatorName = rg.Replace(cells1[20].Range.Text, "");
            string fileCreatorPhone = rg.Replace(cells1[22].Range.Text, "");
            string fileCreatorEmail = rg.Replace(cells1[24].Range.Text, "");

            string technicalReason = rg.Replace(cells1[26].Range.Text, "");
            string businessImpact = rg.Replace(cells1[28].Range.Text, "");
            
            
            
            //Remove MS Word table1 character '•' from the ends of all the text fields

            employerName = employerName.Remove(employerName.Length - 1, 1);
            employerID = employerID.Remove(employerID.Length - 1, 1);
            rquestr = rquestr.Remove(rquestr.Length - 1, 1);
            appManagr = appManagr.Remove(appManagr.Length - 1, 1);
            filesPurpose = filesPurpose.Remove(filesPurpose.Length - 1, 1);
            datasOrigin = datasOrigin.Remove(datasOrigin.Length - 1, 1);
            clientsApproval = clientsApproval.Remove(clientsApproval.Length - 1, 1);
            fileCreatorName = fileCreatorName.Remove(fileCreatorName.Length - 1, 1);
            fileCreatorPhone = fileCreatorPhone.Remove(fileCreatorPhone.Length - 1, 1);
            fileCreatorEmail = fileCreatorEmail.Remove(fileCreatorEmail.Length - 1, 1);
            technicalReason = technicalReason.Remove(technicalReason.Length - 1, 1);
            businessImpact = businessImpact.Remove(businessImpact.Length - 1, 1);
           
            //get rid of carraige return in all fields

            employerName = employerName.TrimEnd('\r', '\n');
            employerID = employerID.TrimEnd('\r', '\n');
            rquestr = rquestr.TrimEnd('\r', '\n');
            appManagr = appManagr.TrimEnd('\r', '\n');
            filesPurpose = filesPurpose.TrimEnd('\r', '\n');
            datasOrigin = datasOrigin.TrimEnd('\r', '\n');
            clientsApproval = clientsApproval.TrimEnd('\r', '\n');
            fileCreatorName = fileCreatorName.TrimEnd('\r', '\n');
            fileCreatorPhone = fileCreatorPhone.TrimEnd('\r', '\n');
            fileCreatorEmail = fileCreatorEmail.TrimEnd('\r', '\n');
            technicalReason = technicalReason.TrimEnd('\r', '\n');
            businessImpact = businessImpact.TrimEnd('\r', '\n');


            if (employerID.Contains(","))
            {
                string[] employerIDs = employerID.Split(',');
                EmployerID = employerIDs;
            }
            else
            {

                string[] employerIDs = {employerID};

                EmployerID = employerIDs;
            }



            
            EmployerName = employerName;
            Requester = rquestr;
            ApprovingManager = appManagr;
            FilesPurpose = filesPurpose;
            DatasOrigin = datasOrigin;
            ClientsApproval = clientsApproval;
            FileCreatorsName = fileCreatorName;
            FileCreatorsPhone = fileCreatorPhone;
            FileCreatorsEmail = fileCreatorEmail;
            TechnicalReason = technicalReason;
            BusinessImpact = businessImpact;
        }


    }
}
