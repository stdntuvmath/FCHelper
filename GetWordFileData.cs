using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace FCHelper_v001
{
    class GetWordFileData
    {
        public string GetWordFileDataMethod(string fileName, out string output1, out string output2,
                                            out string output3, out string output4, out string output5,
                                            out string output6, out string output7, out string output8,
                                            out string output9, out string output10, out string output11,
                                            out string output12, out string output13, out string output14,
                                            out string output15, out string output16)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Documents docs = app.Documents;
            Microsoft.Office.Interop.Word.Document doc = docs.Open(fileName, ReadOnly: true);

            Microsoft.Office.Interop.Word.Table t1 = doc.Tables[1];
            Microsoft.Office.Interop.Word.Range r1 = t1.Range;
            Microsoft.Office.Interop.Word.Cells cells1 = r1.Cells;

            Microsoft.Office.Interop.Word.Table t2 = doc.Tables[2];
            Microsoft.Office.Interop.Word.Range r2 = t2.Range;
            Microsoft.Office.Interop.Word.Cells cells2 = r2.Cells;


            

            //convert tables to text to get rid of bullet points at the end of text fields
            Regex rg = new Regex("^[A-Za-z0-9/$ !|\\[]{}%&()]$");//regex filter on output text accepting characters alphanumeric and /
        

            Regex rgg = new Regex("^[A-Za-z0-9/$ .!|\\[]{}%&()]$");//regex filter on output text accepting characters alphanumeric and /

            //if (!rg.IsMatch(cells1[2].Range.Text))
            //{
            //    string employerName = "";
            //}
            //else
            //{
            //    string employerName = rg.Replace(cells1[2].Range.Text, "");
            //}

            //Table 1


            string employerName = rg.Replace(cells1[2].Range.Text, "");
            string employerID = rg.Replace(cells1[4].Range.Text, "");
            string region = rg.Replace(cells1[6].Range.Text, "");
            string segment = rg.Replace(cells1[8].Range.Text, "");
            string benefitEffectiveDate = rg.Replace(cells1[10].Range.Text, "");
            string currentProducts = rg.Replace(cells1[12].Range.Text, "");
            string addedProducts = rg.Replace(cells1[14].Range.Text, "");
            string newImpFlag = rg.Replace(cells1[16].Range.Text, "");
            string IM_AM = rg.Replace(cells1[18].Range.Text, "");
            string impDeadline = rg.Replace(cells1[20].Range.Text, "");
            string sftpCreds = "";

            //Remove MS Word table1 character '•' from the ends of all the text fields

            employerName = employerName.Remove(employerName.Length - 1, 1);
            employerID = employerID.Remove(employerID.Length - 1, 1);
            region = region.Remove(region.Length - 1, 1);
            segment = segment.Remove(segment.Length - 1, 1);
            benefitEffectiveDate = benefitEffectiveDate.Remove(benefitEffectiveDate.Length - 1, 1);
            currentProducts = currentProducts.Remove(currentProducts.Length - 1, 1);
            addedProducts = addedProducts.Remove(addedProducts.Length - 1, 1);
            newImpFlag = newImpFlag.Remove(newImpFlag.Length - 1, 1);
            IM_AM = IM_AM.Remove(IM_AM.Length - 1, 1);
            impDeadline = impDeadline.Remove(impDeadline.Length - 1, 1);

            //get rid of carraige return in all fields

            employerName = employerName.TrimEnd('\r', '\n');
            employerID = employerID.TrimEnd('\r', '\n');
            region = region.TrimEnd('\r', '\n');
            segment = segment.TrimEnd('\r', '\n');
            benefitEffectiveDate = benefitEffectiveDate.TrimEnd('\r', '\n');
            currentProducts = currentProducts.TrimEnd('\r', '\n');
            addedProducts = addedProducts.TrimEnd('\r', '\n');
            newImpFlag = newImpFlag.TrimEnd('\r', '\n');
            IM_AM = IM_AM.TrimEnd('\r', '\n');
            impDeadline = impDeadline.TrimEnd('\r', '\n');
            sftpCreds = sftpCreds.TrimEnd('\r', '\n');



            //added field to the implementation table
            if (t1.Rows.Count == 11)
            {
                sftpCreds = rg.Replace(cells1[22].Range.Text, "");
            }
            else if (t1.Rows.Count == 10)
            {
                sftpCreds = string.Empty;
            }

            


            output1 = employerName;
            output2 = employerID;
            output3 = region;
            output4 = segment;
            output5 = benefitEffectiveDate;
            output6 = currentProducts;
            output7 = addedProducts;
            output8 = newImpFlag;
            output9 = IM_AM;
            output10 = impDeadline;
            output11 = sftpCreds;
            




            //Table 2


            string contactName = rgg.Replace(cells2[6].Range.Text, "");
            string contactphoneNumber = rgg.Replace(cells2[7].Range.Text, "");
            string contactEmail = rgg.Replace(cells2[8].Range.Text, "");
            string contactType = rgg.Replace(cells2[9].Range.Text, "");
            string fileType = rgg.Replace(cells2[10].Range.Text, "");

            //Remove MS Word table2 character '•' from the ends of all the text fields

            contactName = contactName.Remove(contactName.Length - 1, 1);
            contactphoneNumber = contactphoneNumber.Remove(contactphoneNumber.Length - 1, 1);
            contactEmail = contactEmail.Remove(contactEmail.Length - 1, 1);
            contactType = contactType.Remove(contactType.Length - 1, 1);
            fileType = fileType.Remove(fileType.Length - 1, 1);
           

            //assign output variables
            output12 = contactName;
            output13 = contactphoneNumber;
            output14 = contactEmail;
            output15 = contactType;
            output16 = fileType;



            GetterSetterString getData = new GetterSetterString();

            getData.DataValue = employerName;


            //docs.Close();
            app.Quit();


            return ""; //This is the output door for the outputs to go through

            

        }
    }
}
