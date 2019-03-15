using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ADODB;
using Aspose.Words;
using System.IO;
using System.Configuration;
namespace InfoServerSub
{
    class Program
    {

        static void Main(string[] args)
        {
            string refDoc, tipDoc;
            Aspose.Words.License license= new Aspose.Words.License();
            //crear fitxer log
            //var log = System.Configuration.ConfigurationManager.AppSettings["pathLogs"] + "log.txt";
            //if (File.Exists(log))
            //    File.Delete(log);

            //var fileLog = File.Create(log);


            try
            {
                string strLicense = System.Configuration.ConfigurationManager.AppSettings["pathLogs"];
                license.SetLicense(strLicense + "\\Aspose.Words.lic");



                if (args.Length == 0)
                {
                    //string refdoc1 = 
                    //refDoc = "T017020183415720180808111334";
                    refDoc = "T074A20183426620190301140522";
                    tipDoc = refDoc.Substring(0,5);
                }
                else {
                    
                    refDoc = args[0].ToString().Trim();
                    tipDoc = refDoc.Substring(0, 5);
                    //Console.WriteLine(refDoc);
                    //Console.ReadKey();
                }

                if (refDoc != "")
                {

                    //connectar AS400
                    string str = "Provider=IBMDA400;Data Source=192.168.1.12";


                    Connection conn = new Connection();
                    Command cmd1 = new Command();
                    Recordset rs1 = new Recordset();
                    rs1.CursorLocation = CursorLocationEnum.adUseServer;
                    rs1.CursorType = CursorTypeEnum.adOpenForwardOnly;
                    rs1.LockType = LockTypeEnum.adLockReadOnly;
  
                    Object objRecAff = Type.Missing;
                    
                    conn.Open(str, "COMUNIC5", "COMUNIC5", -1);
                    //cmd1.ActiveConnection = conn;

                    rs1 = conn.Execute("select * from interven2.fdoc where DOC000= '" + refDoc.Trim() +"'", out objRecAff, -1);

                    if (rs1.EOF) {
                        rs1.Close();
                        conn.Close();
                        //Console.WriteLine("no existeix dades..");
                        throw new Exception("no existeix dades");
                    }

                        //proba obrir plantilla
                        
                        //string pathDot = @"//sarroca/aplica/Visatge/dot/solicitud_sub.docx";
                        //generar sol·licitud
                        string strFileTarget = Directory.GetCurrentDirectory();
                        string docTemplate = System.Configuration.ConfigurationManager.AppSettings["pathTemplates"];

                        //string STRAMIT = args[0].ToString();
                        if (tipDoc == "T0070" || tipDoc=="T0040")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "DOCPS.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC022":
                                        bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                        break;
                                    case "DOC023":
                                        bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                        break;
                                    case "DOC024":
                                        bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                        break;
                                    case "DOC025":
                                        bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                        break;
                                    case "DOC026":
                                        bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                        break;
                                    case "DOC027":
                                        bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                        break;
                                    case "DOC028":
                                        bookmark.Text = rs1.Fields["DOC028"].Value.ToString();
                                        break;
                                    case "DOC029":
                                        bookmark.Text = rs1.Fields["DOC029"].Value.ToString();
                                        break;
                                    case "DOC030":
                                        bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                        break;
                                    case "DOC031":
                                        bookmark.Text = rs1.Fields["DOC031"].Value.ToString();
                                        break;
                                    case "DOC032":
                                        bookmark.Text = rs1.Fields["DOC032"].Value.ToString();
                                        break;
                                    case "DOC033":
                                        bookmark.Text = rs1.Fields["DOC033"].Value.ToString();
                                        break;
                                    case "DOC034":
                                        bookmark.Text = rs1.Fields["DOC034"].Value.ToString();
                                        break;
                                    case "DOC035":
                                        bookmark.Text = rs1.Fields["DOC035"].Value.ToString();
                                        break;
                                    case "DOC036":
                                        bookmark.Text = rs1.Fields["DOC036"].Value.ToString();
                                        break;
                                    case "DOC037":
                                        bookmark.Text = rs1.Fields["DOC037"].Value.ToString();
                                        break;
                                    case "DOC038":
                                        bookmark.Text = rs1.Fields["DOC038"].Value.ToString();
                                        break;
                                    case "DOC039":
                                        bookmark.Text = rs1.Fields["DOC039"].Value.ToString();
                                        break;
                                    case "DOC040":
                                        bookmark.Text = rs1.Fields["DOC041"].Value.ToString();
                                        break;
                                    case "DOC041":
                                        bookmark.Text = rs1.Fields["DOC041"].Value.ToString();
                                        break;
                                    case "DOC042":
                                        bookmark.Text = rs1.Fields["DOC042"].Value.ToString();
                                        break;
                                    case "DOC043":
                                        bookmark.Text = rs1.Fields["DOC043"].Value.ToString();
                                        break;
                                    case "DOC044":
                                        bookmark.Text = rs1.Fields["DOC044"].Value.ToString();
                                        break;
                                    case "DOC045":
                                        bookmark.Text = rs1.Fields["DOC045"].Value.ToString();
                                        break;
                                    case "DOC046":
                                        bookmark.Text = rs1.Fields["DOC046"].Value.ToString();
                                        break;
                                    case "DOC047":
                                        bookmark.Text = rs1.Fields["DOC047"].Value.ToString();
                                        break;
                                    case "DOC048":
                                        bookmark.Text = rs1.Fields["DOC048"].Value.ToString();
                                        break;
                                    case "DOC049":
                                        bookmark.Text = rs1.Fields["DOC049"].Value.ToString();
                                        break;
                                    case "DOC050":
                                        bookmark.Text = rs1.Fields["DOC050"].Value.ToString();
                                        break;
                                    case "DOC051":
                                        bookmark.Text = rs1.Fields["DOC051"].Value.ToString();
                                        break;
                                    case "DOC052":
                                        bookmark.Text = rs1.Fields["DOC052"].Value.ToString();
                                        break;

                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }
                     
                        if (tipDoc == "T024B" || tipDoc == "T050B")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "DOCRCA.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC022":
                                        bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                        break;
                                    case "DOC023":
                                        bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                        break;
                                    case "DOC024":
                                        bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                        break;
                                    case "DOC025":
                                        bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                        break;
                                    case "DOC026":
                                        bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                        break;
                                    case "DOC027":
                                        bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                        break;
                                    case "DOC028":
                                        bookmark.Text = rs1.Fields["DOC028"].Value.ToString();
                                        break;
                                    case "DOC029":
                                        bookmark.Text = rs1.Fields["DOC029"].Value.ToString();
                                        break;
                                    case "DOC030":
                                        bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                        break;
                                    case "DOC031":
                                        bookmark.Text = rs1.Fields["DOC031"].Value.ToString();
                                        break;
                                    case "DOC032":
                                        bookmark.Text = rs1.Fields["DOC032"].Value.ToString();
                                        break;
                                    case "DOC033":
                                        bookmark.Text = rs1.Fields["DOC033"].Value.ToString();
                                        break;
                                    case "DOC034":
                                        bookmark.Text = rs1.Fields["DOC034"].Value.ToString();
                                        break;
                                    case "DOC035":
                                        bookmark.Text = rs1.Fields["DOC035"].Value.ToString();
                                        break;
                                    case "DOC036":
                                        bookmark.Text = rs1.Fields["DOC036"].Value.ToString();
                                        break;
                                    case "DOC037":
                                        bookmark.Text = rs1.Fields["DOC037"].Value.ToString();
                                        break;
                                    case "DOC038":
                                        bookmark.Text = rs1.Fields["DOC038"].Value.ToString();
                                        break;
                                    case "DOC039":
                                        bookmark.Text = rs1.Fields["DOC039"].Value.ToString();
                                        break;
                                    case "DOC040":
                                        bookmark.Text = rs1.Fields["DOC040"].Value.ToString();
                                        break;
                                    case "DOC041":
                                        bookmark.Text = rs1.Fields["DOC041"].Value.ToString();
                                        break;
                                    case "DOC042":
                                        bookmark.Text = rs1.Fields["DOC042"].Value.ToString();
                                        break;
                                    case "DOC043":
                                        bookmark.Text = rs1.Fields["DOC043"].Value.ToString();
                                        break;
                                    case "DOC044":
                                        bookmark.Text = rs1.Fields["DOC044"].Value.ToString();
                                        break;
                                    case "DOC045":
                                        bookmark.Text = rs1.Fields["DOC045"].Value.ToString();
                                        break;
                                    case "DOC046":
                                        bookmark.Text = rs1.Fields["DOC046"].Value.ToString();
                                        break;
                                    case "DOC047":
                                        bookmark.Text = rs1.Fields["DOC047"].Value.ToString();
                                        break;
                                    case "DOC048":
                                        bookmark.Text = rs1.Fields["DOC048"].Value.ToString();
                                        break;
                                    case "DOC049":
                                        bookmark.Text = rs1.Fields["DOC049"].Value.ToString();
                                        break;
                                    case "DOC050":
                                        bookmark.Text = rs1.Fields["DOC050"].Value.ToString();
                                        break;
                                    case "DOC051":
                                        bookmark.Text = rs1.Fields["DOC051"].Value.ToString();
                                        break;
                                    case "DOC052":
                                        bookmark.Text = rs1.Fields["DOC052"].Value.ToString();
                                        break;
                                    case "DOC053":
                                        bookmark.Text = rs1.Fields["DOC053"].Value.ToString();
                                        break;
                                    case "DOC054":
                                        bookmark.Text = rs1.Fields["DOC054"].Value.ToString();
                                        break;
                                    case "DOC055":
                                        bookmark.Text = rs1.Fields["DOC055"].Value.ToString();
                                        break;
                                    case "DOC056":
                                        bookmark.Text = rs1.Fields["DOC056"].Value.ToString();
                                        break;
                                    case "DOC057":
                                        bookmark.Text = rs1.Fields["DOC057"].Value.ToString();
                                        break;
                                    case "DOC058":
                                        bookmark.Text = rs1.Fields["DOC058"].Value.ToString();
                                        break;
                                    case "DOC059":
                                        bookmark.Text = rs1.Fields["DOC059"].Value.ToString();
                                        break;
                                    case "DOC060":
                                        bookmark.Text = rs1.Fields["DOC060"].Value.ToString();
                                        break;
                                    case "DOC061":
                                        bookmark.Text = rs1.Fields["DOC061"].Value.ToString();
                                        break;
                                    case "DOC062":
                                        bookmark.Text = rs1.Fields["DOC062"].Value.ToString();
                                        break;
                                    case "DOC063":
                                        bookmark.Text = rs1.Fields["DOC063"].Value.ToString();
                                        break;
                                    case "DOC064":
                                        bookmark.Text = rs1.Fields["DOC064"].Value.ToString();
                                        break;
                                    case "DOC065":
                                        bookmark.Text = rs1.Fields["DOC065"].Value.ToString();
                                        break;
                                    case "DOC066":
                                        bookmark.Text = rs1.Fields["DOC066"].Value.ToString();
                                        break;
                                    case "DOC067":
                                        bookmark.Text = rs1.Fields["DOC067"].Value.ToString();
                                        break;
                                    case "DOC068":
                                        bookmark.Text = rs1.Fields["DOC068"].Value.ToString();
                                        break;
                                    case "DOC069":
                                        bookmark.Text = rs1.Fields["DOC069"].Value.ToString();
                                        break;
                                   
                                        
                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }


                        if (tipDoc == "T024A" || tipDoc=="T050A" || tipDoc=="T0114" || tipDoc=="T0142")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "DOCJP.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC022":
                                        bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                        break;
                                    case "DOC023":
                                        bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                        break;
                                    case "DOC024":
                                        bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                        break;
                                    case "DOC025":
                                        bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                        break;
                                    case "DOC026":
                                        bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                        break;
                                    case "DOC027":
                                        bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                        break;
                                    case "DOC028":
                                        bookmark.Text = rs1.Fields["DOC028"].Value.ToString();
                                        break;
                                    case "DOC029":
                                        bookmark.Text = rs1.Fields["DOC029"].Value.ToString();
                                        break;
                                    case "DOC030":
                                        bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                        break;
                                    case "DOC031":
                                        bookmark.Text = rs1.Fields["DOC031"].Value.ToString();
                                        break;
                                    case "DOC032":
                                        bookmark.Text = rs1.Fields["DOC032"].Value.ToString();
                                        break;
                                    case "DOC033":
                                        bookmark.Text = rs1.Fields["DOC033"].Value.ToString();
                                        break;
                                    case "DOC034":
                                        bookmark.Text = rs1.Fields["DOC034"].Value.ToString();
                                        break;
                                    case "DOC035":
                                        bookmark.Text = rs1.Fields["DOC035"].Value.ToString();
                                        break;
                                    case "DOC036":
                                        bookmark.Text = rs1.Fields["DOC036"].Value.ToString();
                                        break;
                                    case "DOC037":
                                        bookmark.Text = rs1.Fields["DOC037"].Value.ToString();
                                        break;
                                    case "DOC038":
                                        bookmark.Text = rs1.Fields["DOC038"].Value.ToString();
                                        break;
                                    case "DOC039":
                                        bookmark.Text = rs1.Fields["DOC039"].Value.ToString();
                                        break;
                                    case "DOC040":
                                        bookmark.Text = rs1.Fields["DOC040"].Value.ToString();
                                        break;
                                    case "DOC041":
                                        bookmark.Text = rs1.Fields["DOC041"].Value.ToString();
                                        break;
                                    case "DOC042":
                                        bookmark.Text = rs1.Fields["DOC042"].Value.ToString();
                                        break;
                                    case "DOC043":
                                        bookmark.Text = rs1.Fields["DOC043"].Value.ToString();
                                        break;
                                    case "DOC044":
                                        bookmark.Text = rs1.Fields["DOC044"].Value.ToString();
                                        break;
                                    case "DOC045":
                                        bookmark.Text = rs1.Fields["DOC045"].Value.ToString();
                                        break;
                                    case "DOC046":
                                        bookmark.Text = rs1.Fields["DOC046"].Value.ToString();
                                        break;
                                    case "DOC047":
                                        bookmark.Text = rs1.Fields["DOC047"].Value.ToString();
                                        break;
                                    case "DOC048":
                                        bookmark.Text = rs1.Fields["DOC048"].Value.ToString();
                                        break;
                                    case "DOC049":
                                        bookmark.Text = rs1.Fields["DOC049"].Value.ToString();
                                        break;
                                    case "DOC050":
                                        bookmark.Text = rs1.Fields["DOC050"].Value.ToString();
                                        break;
                                    case "DOC051":
                                        bookmark.Text = rs1.Fields["DOC051"].Value.ToString();
                                        break;
                                    case "DOC052":
                                        bookmark.Text = rs1.Fields["DOC052"].Value.ToString();
                                        break;
                                    case "DOC053":
                                        bookmark.Text = rs1.Fields["DOC053"].Value.ToString();
                                        break;
                                    case "DOC054":
                                        bookmark.Text = rs1.Fields["DOC054"].Value.ToString();
                                        break;
                                    case "DOC055":
                                        bookmark.Text = rs1.Fields["DOC055"].Value.ToString();
                                        break;
                                    case "DOC056":
                                        bookmark.Text = rs1.Fields["DOC056"].Value.ToString();
                                        break;
                                    case "DOC057":
                                        bookmark.Text = rs1.Fields["DOC057"].Value.ToString();
                                        break;
                                    case "DOC058":
                                        bookmark.Text = rs1.Fields["DOC058"].Value.ToString();
                                        break;
                                    case "DOC059":
                                        bookmark.Text = rs1.Fields["DOC059"].Value.ToString();
                                        break;
                                    case "DOC060":
                                        bookmark.Text = rs1.Fields["DOC060"].Value.ToString();
                                        break;
                                    case "DOC061":
                                        bookmark.Text = rs1.Fields["DOC061"].Value.ToString();
                                        break;
                                    case "DOC062":
                                        bookmark.Text = rs1.Fields["DOC062"].Value.ToString();
                                        break;
                                        
                                        

                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }


                        if (tipDoc == "T0126")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "DOCPP.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC022":
                                        bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                        break;
                                    case "DOC023":
                                        bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                        break;
                                    case "DOC024":
                                        bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                        break;
                                    case "DOC025":
                                        bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                        break;
                                    case "DOC026":
                                        bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                        break;
                                    case "DOC027":
                                        bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                        break;
                                    case "DOC028":
                                        bookmark.Text = rs1.Fields["DOC028"].Value.ToString();
                                        break;
                                    case "DOC029":
                                        bookmark.Text = rs1.Fields["DOC029"].Value.ToString();
                                        break;
                                    case "DOC030":
                                        bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                        break;
                                    case "DOC031":
                                        bookmark.Text = rs1.Fields["DOC031"].Value.ToString();
                                        break;
                                    case "DOC032":
                                        bookmark.Text = rs1.Fields["DOC032"].Value.ToString();
                                        break;
                                    case "DOC033":
                                        bookmark.Text = rs1.Fields["DOC033"].Value.ToString();
                                        break;
                                    case "DOC034":
                                        bookmark.Text = rs1.Fields["DOC034"].Value.ToString();
                                        break;
                                    case "DOC035":
                                        bookmark.Text = rs1.Fields["DOC035"].Value.ToString();
                                        break;
                                    case "DOC036":
                                        bookmark.Text = rs1.Fields["DOC036"].Value.ToString();
                                        break;
                                    case "DOC037":
                                        bookmark.Text = rs1.Fields["DOC037"].Value.ToString();
                                        break;
                                    case "DOC038":
                                        bookmark.Text = rs1.Fields["DOC038"].Value.ToString();
                                        break;
                                    case "DOC039":
                                        bookmark.Text = rs1.Fields["DOC039"].Value.ToString();
                                        break;
                                    case "DOC040":
                                        bookmark.Text = rs1.Fields["DOC040"].Value.ToString();
                                        break;
                                    case "DOC041":
                                        bookmark.Text = rs1.Fields["DOC041"].Value.ToString();
                                        break;
                                    case "DOC042":
                                        bookmark.Text = rs1.Fields["DOC042"].Value.ToString();
                                        break;
                                    case "DOC043":
                                        bookmark.Text = rs1.Fields["DOC043"].Value.ToString();
                                        break;
                                    case "DOC044":
                                        bookmark.Text = rs1.Fields["DOC044"].Value.ToString();
                                        break;
                                    case "DOC045":
                                        bookmark.Text = rs1.Fields["DOC045"].Value.ToString();
                                        break;
                                    case "DOC046":
                                        bookmark.Text = rs1.Fields["DOC046"].Value.ToString();
                                        break;
                                    case "DOC047":
                                        bookmark.Text = rs1.Fields["DOC047"].Value.ToString();
                                        break;
                                    case "DOC048":
                                        bookmark.Text = rs1.Fields["DOC048"].Value.ToString();
                                        break;
                                    case "DOC049":
                                        bookmark.Text = rs1.Fields["DOC049"].Value.ToString();
                                        break;
                                   

                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }

                        if (tipDoc == "T0170")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "DOCCD.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC022":
                                        bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                        break;
                                    case "DOC023":
                                        bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                        break;
                                    case "DOC024":
                                        bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                        break;
                                    case "DOC025":
                                        bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                        break;
                                    case "DOC026":
                                        bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                        break;
                                    case "DOC027":
                                        bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                        break;
                                    case "DOC028":
                                        bookmark.Text = rs1.Fields["DOC028"].Value.ToString();
                                        break;
                                    case "DOC029":
                                        bookmark.Text = rs1.Fields["DOC029"].Value.ToString();
                                        break;
                                    case "DOC030":
                                        bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                        break;
                                    case "DOC031":
                                        bookmark.Text = rs1.Fields["DOC031"].Value.ToString();
                                        break;
                                    case "DOC032":
                                        bookmark.Text = rs1.Fields["DOC032"].Value.ToString();
                                        break;
                                    case "DOC033":
                                        bookmark.Text = rs1.Fields["DOC033"].Value.ToString();
                                        break;
                                    case "DOC034":
                                        bookmark.Text = rs1.Fields["DOC034"].Value.ToString();
                                        break;
                                    case "DOC035":
                                        bookmark.Text = rs1.Fields["DOC035"].Value.ToString();
                                        break;
                                    case "DOC036":
                                        bookmark.Text = rs1.Fields["DOC036"].Value.ToString();
                                        break;
                                    case "DOC037":
                                        bookmark.Text = rs1.Fields["DOC037"].Value.ToString();
                                        break;
                                    case "DOC038":
                                        bookmark.Text = rs1.Fields["DOC038"].Value.ToString();
                                        break;
                                    case "DOC039":
                                        bookmark.Text = rs1.Fields["DOC039"].Value.ToString();
                                        break;
                                    case "DOC040":
                                        bookmark.Text = rs1.Fields["DOC040"].Value.ToString();
                                        break;
                                    case "DOC041":
                                        bookmark.Text = rs1.Fields["DOC041"].Value.ToString();
                                        break;
                                    case "DOC042":
                                        bookmark.Text = rs1.Fields["DOC042"].Value.ToString();
                                        break;
                                    case "DOC043":
                                        bookmark.Text = rs1.Fields["DOC043"].Value.ToString();
                                        break;
                                    case "DOC044":
                                        bookmark.Text = rs1.Fields["DOC044"].Value.ToString();
                                        break;
                                    case "DOC045":
                                        bookmark.Text = rs1.Fields["DOC045"].Value.ToString();
                                        break;
                                    case "DOC046":
                                        bookmark.Text = rs1.Fields["DOC046"].Value.ToString();
                                        break;
                                    case "DOC047":
                                        bookmark.Text = rs1.Fields["DOC047"].Value.ToString();
                                        break;
                                    case "DOC048":
                                        bookmark.Text = rs1.Fields["DOC048"].Value.ToString();
                                        break;
                                    case "DOC049":
                                        bookmark.Text = rs1.Fields["DOC049"].Value.ToString();
                                        break;
                                    case "DOC050":
                                        bookmark.Text = rs1.Fields["DOC050"].Value.ToString();
                                        break;
                                   

                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                           
                            doc.Print();
                        }

                    if (tipDoc == "T0010")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T10.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "tipDoc":
                                        bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                        break;
                                    case "DOC022":
                                        bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                        break;
                                    case "DOC026":
                                        bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC024":
                                        bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                        break;
                                    case "DOC025":
                                        bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                        break;
                                    case "DOC030":
                                        bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                        break;
                                    case "DOC040":
                                        bookmark.Text = rs1.Fields["DOC040"].Value.ToString();
                                        break;

                              }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }

                        if (tipDoc == "T0012")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T12.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "tipDoc":
                                        bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC022":
                                        bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                        break;                    


                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }

                  
                     if (tipDoc == "T0128")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T128.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "tipDoc":
                                        bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC022":
                                        bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                        break;
                                    case "DOC026":
                                        bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                        break;
                                    case "DOC027":
                                        bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                        break;
                                    case "DOC028":
                                        bookmark.Text = rs1.Fields["DOC028"].Value.ToString();
                                        break;

                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }
                    if (tipDoc == "T0071")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T71.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "tipDoc":
                                        bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC08"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                               }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }

                    if (tipDoc == "T0018")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T18.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "tipDoc":
                                        bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                        break;
                                    case "DOC030":
                                        bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC022":
                                        bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                        break;
                                    case "DOC023":
                                        bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                        break;
                                    case "DOC024":
                                        bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                        break;
                                    case "DOC025":
                                        bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                        break;
                                    case "DOC026":
                                        bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                        break;
                                    case "DOC027":
                                        bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                        break;
                                    case "DOC028":
                                        bookmark.Text = rs1.Fields["DOC028"].Value.ToString();
                                        break;
                                    case "DOC029":
                                        bookmark.Text = rs1.Fields["DOC029"].Value.ToString();
                                        break;
                                    case "DOC031":
                                        bookmark.Text = rs1.Fields["DOC031"].Value.ToString();
                                        break;
                                    case "DOC032":
                                        bookmark.Text = rs1.Fields["DOC032"].Value.ToString();
                                        break;
                                    case "DOC033":
                                        bookmark.Text = rs1.Fields["DOC033"].Value.ToString();
                                        break;
                                    case "DOC034":
                                        bookmark.Text = rs1.Fields["DOC034"].Value.ToString();
                                        break;
                                    case "DOC035":
                                        bookmark.Text = rs1.Fields["DOC035"].Value.ToString();
                                        break;

                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }
                     if (tipDoc == "T0014")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T14.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "tipDoc":
                                        bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC038":
                                        bookmark.Text = rs1.Fields["DOC038"].Value.ToString();
                                        break;
                                    case "DOC039":
                                        bookmark.Text = rs1.Fields["DOC039"].Value.ToString();
                                        break;
                                    case "DOC045":
                                        bookmark.Text = rs1.Fields["DOC045"].Value.ToString();
                                        break;
                                    case "DOC046":
                                        bookmark.Text = rs1.Fields["DOC046"].Value.ToString();
                                        break;
                                    case "DOC47":
                                        bookmark.Text = rs1.Fields["DOC047"].Value.ToString();
                                        break;

                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }
                     if (tipDoc == "T0020")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T20.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "tipDoc":
                                        bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC022":
                                        bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                        break;
                                    case "DOC023":
                                        bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                        break;
                                    case "DOC025":
                                        bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                        break;
                                    case "DOC026":
                                        bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                        break;

                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }

                     if (tipDoc == "T0024")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T24.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "tipDoc":
                                        bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC023":
                                        bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                        break;
                                  

                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }

                     if (tipDoc == "E0024" || tipDoc== "E0050")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T24B.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "tipDoc":
                                        bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC024":
                                        bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                        break;
                                    case "DOC025":
                                        bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                        break;
                                    case "DOC026":
                                        bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                        break;
                                    case "DOC027":
                                        bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                        break;
                                    case "DOC028":
                                        bookmark.Text = rs1.Fields["DOC028"].Value.ToString();
                                        break;
                                    case "DOC029":
                                        bookmark.Text = rs1.Fields["DOC029"].Value.ToString();
                                        break;
                                    case "DOC030":
                                        bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                        break;

                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }

                     if (tipDoc == "E0028"  ||  tipDoc== "E0082")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T28B.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "tipDoc":
                                        bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC024":
                                        bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                        break;
                                    case "DOC025":
                                        bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                        break;
                                    case "DOC026":
                                        bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                        break;
                                    case "DOC027":
                                        bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                        break;
                                    case "DOC028":
                                        bookmark.Text = rs1.Fields["DOC028"].Value.ToString();
                                        break;
                                    case "DOC029":
                                        bookmark.Text = rs1.Fields["DOC029"].Value.ToString();
                                        break;
                                    case "DOC030":
                                        bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                        break;


                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }

                    if (tipDoc == "T0116")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T28.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "tipDoc":
                                        bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC023":
                                        bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                        break;

                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }

                     if (tipDoc == "T0028")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T28.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "tipDoc":
                                        bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC023":
                                        bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                        break;

                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }

                     if (tipDoc == "T0072")
                     {
                         if (rs1.Fields["DOC036"].Value.ToString() == "P")
                         {
                             Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T72b.dot");
                         }
                         else
                         {
                             Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T72.dot");                       
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC030":
                                     bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                     break;
                                 case "DOC031":
                                     bookmark.Text = rs1.Fields["DOC031"].Value.ToString();
                                     break;
                                 case "DOC001":
                                     bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC003":
                                     bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                     break;
                                 case "DOC004":
                                     bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                     break;
                                 case "DOC005":
                                     bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                     break;
                                 case "DOC006":
                                     bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                     break;
                                 case "DOC007":
                                     bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                     break;
                                 case "DOC008":
                                     bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                     break;
                                 case "DOC011":
                                     bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                     break;
                                 case "DOC021":
                                     bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                     break;
                                 case "DOC025":
                                     bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                     break;
                                 case "DOC026":
                                     bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                     break;
                                 case "DOC027":
                                     bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                     break;
                                 case "DOC028":
                                     bookmark.Text = rs1.Fields["DOC028"].Value.ToString();
                                     break;
                                 case "DOC032":
                                     bookmark.Text = rs1.Fields["DOC032"].Value.ToString();
                                     break;
                                 case "DOC033":
                                     bookmark.Text = rs1.Fields["DOC033"].Value.ToString();
                                     break;
                                 case "DOC035":
                                     bookmark.Text = rs1.Fields["DOC035"].Value.ToString();
                                     break;
                             
                                }
                            }                        
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                         }
                        }
                         
                     if (tipDoc == "T0072P")
                     {
                         Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T72PAM.dot");
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC030":
                                     bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                     break;
                                 case "DOC031":
                                     bookmark.Text = rs1.Fields["DOC031"].Value.ToString();
                                     break;
                                 case "DOC001":
                                     bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC0311":
                                     bookmark.Text = rs1.Fields["DOC0311"].Value.ToString();
                                     break;
                                 case "DOC004":
                                     bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                     break;
                                 case "DOC005":
                                     bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                     break;
                                 case "DOC050":
                                     bookmark.Text = rs1.Fields["DOC050"].Value.ToString();
                                     break;
                                 case "DOC009":
                                     bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                     break;
                                 case "DOC021":
                                     bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                     break;
                                 case "DOC022":
                                     bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                     break;
                                 case "DOC023":
                                     bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                     break;
                                 case "DOC024":
                                     bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                     break;
                                 case "DOC025":
                                     bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                     break;
                                 case "DOC026":
                                     bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                     break;
                                 case "DOC027":
                                     bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                     break;


                             }
                         }
                         Console.WriteLine("Document al forn... un segon!");
                         doc.Print();
                     }

                     if (tipDoc == "T0030")
                     {
                         Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T30N.dot");
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC034":
                                     bookmark.Text = rs1.Fields["DOC034"].Value.ToString();
                                     break;
                                 case "DOC001":
                                     bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC003":
                                     bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                     break;
                                 case "DOC006":
                                     bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                     break;
                                 case "DOC011":
                                     bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                     break;
                                 case "DOC012":
                                     bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                     break;
                                 case "DOC021":
                                     bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                     break;
                                 case "DOC022":
                                     bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                     break;
                                 case "DOC023":
                                     bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                     break;
                                 case "DOC024":
                                     bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                     break;
                                 case "DOC025":
                                     bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                     break;
                                 case "DOC026":
                                     bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                     break;
                                 case "DOC027":
                                     bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                     break;
                                 case "DOC030":
                                     bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                     break;
                                 case "DOC031":
                                     bookmark.Text = rs1.Fields["DOC031"].Value.ToString();
                                     break;
                                 case "DOC032":
                                     bookmark.Text = rs1.Fields["DOC032"].Value.ToString();
                                     break;
                                 case "DOC033":
                                     bookmark.Text = rs1.Fields["DOC033"].Value.ToString();
                                     break;
                                 case "DOC035":
                                     bookmark.Text = rs1.Fields["DOC035"].Value.ToString();
                                     break;
                                 case "DOC036":
                                     bookmark.Text = rs1.Fields["DOC036"].Value.ToString();
                                     break;
                                 case "DOC037":
                                     bookmark.Text = rs1.Fields["DOC037"].Value.ToString();
                                     break;
                                 case "DOC038":
                                     bookmark.Text = rs1.Fields["DOC038"].Value.ToString();
                                     break;
                                 case "DOC039":
                                     bookmark.Text = rs1.Fields["DOC039"].Value.ToString();
                                     break;
                                 case "DOC040":
                                     bookmark.Text = rs1.Fields["DOC040"].Value.ToString();
                                     break;
                                 
                             }
                         }
                         Console.WriteLine("Document al forn... un segon!");
                         doc.Print();
                     }

                     if (tipDoc == "T0094")
                     {
                         Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T94.dot");
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC030":
                                     bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                     break;
                                 case "DOC031":
                                     bookmark.Text = rs1.Fields["DOC031"].Value.ToString();
                                     break;
                                 case "DOC001":
                                     bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC003":
                                     bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                     break;
                                 case "DOC004":
                                     bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                     break;
                                 case "DOC005":
                                     bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                     break;
                                 case "DOC006":
                                     bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                     break;
                                 case "DOC007":
                                     bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                     break;
                                 case "DOC008":
                                     bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                     break;
                                 case "DOC021":
                                     bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                     break;
                                 case "DOC022":
                                     bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                     break;
                                 case "DOC023":
                                     bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                     break;
                                 case "DOC024":
                                     bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                     break;
                                 case "DOC025":
                                     bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                     break;
                                 case "DOC026":
                                     bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                     break;
                                 case "DOC027":
                                     bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                     break;
                                 case "DOC032":
                                     bookmark.Text = rs1.Fields["DOC032"].Value.ToString();
                                     break;


                             }
                         }
                         Console.WriteLine("Document al forn... un segon!");
                         doc.Print();
                     }

                     if (tipDoc == "T0080")
                     {                         
                         Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T80.dot");
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC020":
                                     bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                     break;
                                 case "DOC021":
                                     bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                     break;
                                 case "DOC022":
                                     bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                     break;
                                 case "DOC001":
                                     bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC003":
                                     bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                     break;
                                 case "DOC004":
                                     bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                     break;
                                 case "DOC005":
                                     bookmark.Text = rs1.Fields["DOC05"].Value.ToString();
                                     break;
                                 case "DOC006":
                                     bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                     break;
                                 case "DOC008":
                                     bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                     break;
                                 case "DOC009":
                                     bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                     break;
                                 case "DOC010":
                                     bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                     break;
                                 case "DOC014":
                                     bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                     break;
                                 case "DOC015":
                                     bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                     break;
                                 case "DOC016":
                                     bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                     break;


                             }
                         }
                         Console.WriteLine("Document al forn... un segon!");
                         doc.Print();
                     }

                     if (tipDoc == "T0068")
                     {
                         Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T68.dot");
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC032":
                                     bookmark.Text = rs1.Fields["DOC032"].Value.ToString();
                                     break;
                                 case "DOC033":
                                     bookmark.Text = rs1.Fields["DOC033"].Value.ToString();
                                     break;
                                 case "DOC001":
                                     bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC003":
                                     bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                     break;
                                 case "DOC004":
                                     bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                     break;
                                 case "DOC005":
                                     bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                     break;
                                 case "DOC006":
                                     bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                     break;
                                 case "DOC007":
                                     bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                     break;
                                 case "DOC008":
                                     bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                     break;
                                 case "DOC011":
                                     bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                     break;
                                 case "DOC021":
                                     bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                     break;
                                 case "DOC022":
                                     bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                     break;
                                 case "DOC023":
                                     bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                     break;
                                 case "DOC024":
                                     bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                     break;
                                 case "DOC025":
                                     bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                     break;
                                 case "DOC026":
                                     bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                     break;
                                 case "DOC027":
                                     bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                     break;
                                 case "DOC028":
                                     bookmark.Text = rs1.Fields["DOC028"].Value.ToString();
                                     break;
                                 case "DOC029":
                                     bookmark.Text = rs1.Fields["DOC029"].Value.ToString();
                                     break;
                                 case "DOC030":
                                     bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                     break;
                                 case "DOC031":
                                     bookmark.Text = rs1.Fields["DOC031"].Value.ToString();
                                     break;

                             }
                         }
                         Console.WriteLine("Document al forn... un segon!");
                         doc.Print();
                     }

                     if (tipDoc == "T0130")
                     {
                         if (rs1.Fields["DOC040"].Value.ToString() == "N")
                         {
                             Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T130.dot");
                         }
                         else
                         {
                         Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T130N.dot");                                             
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC030":
                                     bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                     break;
                                 case "DOC031":
                                     bookmark.Text = rs1.Fields["DOC031"].Value.ToString();
                                     break;
                                 case "DOC001":
                                     bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC003":
                                     bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                     break;
                                 case "DOC004":
                                     bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                     break;
                                 case "DOC006":
                                     bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                     break;
                                 case "DOC011":
                                     bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                     break;
                                 case "DOC012":
                                     bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                     break;
                                 case "DOC020":
                                     bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                     break;
                                 case "DOC021":
                                     bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                     break;
                                 case "DOC022":
                                     bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                     break;
                                 case "DOC023":
                                     bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                     break;
                                 case "DOC024":
                                     bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                     break;
                                 case "DOC025":
                                     bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                     break;
                                 case "DOC026":
                                     bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                     break;
                                 case "DOC027":
                                     bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                     break;
                                 case "DOC032":
                                     bookmark.Text = rs1.Fields["DOC032"].Value.ToString();
                                     break;
                                 case "DOC033":
                                     bookmark.Text = rs1.Fields["DOC033"].Value.ToString();
                                     break;


                             }
                         }
                         Console.WriteLine("Document al forn... un segon!");
                         doc.Print();
                         }
                     }

                     if (tipDoc == "T5268")
                     {
                         Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T5268.dot");
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC022":
                                     bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC003":
                                     bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                     break;
                                 case "DOC004":
                                     bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                     break;
                                 case "DOC005":
                                     bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                     break;
                                 case "DOC009":
                                     bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                     break;
                                 case "DOC010":
                                     bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                     break;
                                 case "DOC011":
                                     bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                     break;
                                 case "DOC012":
                                     bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                     break;
                                 case "DOC013":
                                     bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                     break;
                                 case "DOC014":
                                     bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                     break;
                                 case "DOC015":
                                     bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                     break;
                                 case "DOC016":
                                     bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                     break;
                                 case "DOC017":
                                     bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                     break;
                                 case "DOC018":
                                     bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                     break;
                                 case "DOC019":
                                     bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                     break;
                                 case "DOC020":
                                     bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                     break;
                                 case "DOC021":
                                     bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                     break;
             
                             }
                         }
                         Console.WriteLine("Document al forn... un segon!");
                         doc.Print();
                     }

                     if (tipDoc == "T2272")
                     {
                         Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T2272.dot");
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC022":
                                     bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                     break;
                                 case "DOC023":
                                     bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC003":
                                     bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                     break;
                                 case "DOC004":
                                     bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                     break;
                                 case "DOC005":
                                     bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                     break;
                                 case "DOC009":
                                     bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                     break;
                                 case "DOC010":
                                     bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                     break;
                                 case "DOC011":
                                     bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                     break;
                                 case "DOC012":
                                     bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                     break;
                                 case "DOC013":
                                     bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                     break;
                                 case "DOC014":
                                     bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                     break;
                                 case "DOC015":
                                     bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                     break;
                                 case "DOC016":
                                     bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                     break;
                                 case "DOC017":
                                     bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                     break;
                                 case "DOC018":
                                     bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                     break;
                                 case "DOC019":
                                     bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                     break;
                                 case "DOC020":
                                     bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                     break;
                                 case "DOC021":
                                     bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                     break;

                             }
                         }
                         Console.WriteLine("Document al forn... un segon!");
                         doc.Print();
                     }

                     if (tipDoc == "T0044")
                     {
                         Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T44.dot");
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC018":
                                     bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                     break;
                                 case "DOC001":
                                     bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC003":
                                     bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                     break;
                                 case "DOC004":
                                     bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                     break;
                                 case "DOC005":
                                     bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                     break;
                                 case "DOC006":
                                     bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                     break;
                                 case "DOC007":
                                     bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                     break;
                                 case "DOC008":
                                     bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                     break;
                                 case "DOC009":
                                     bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                     break;
                                 case "DOC010":
                                     bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                     break;
                                 case "DOC011":
                                     bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                     break;
                                 case "DOC014":
                                     bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                     break;
                                 case "DOC015":
                                     bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                     break;
                                 case "DOC016":
                                     bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                     break;

                             }
                         }
                         Console.WriteLine("Document al forn... un segon!");
                         doc.Print();
                     }

                     if (tipDoc == "T0056")
                     {
                         Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T56.dot");
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC017":
                                     bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                     break;
                                 case "DOC001":
                                     bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC003":
                                     bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                     break;
                                 case "DOC004":
                                     bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                     break;
                                 case "DOC005":
                                     bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                     break;
                                 case "DOC006":
                                     bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                     break;
                                 case "DOC007":
                                     bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                     break;
                                 case "DOC008":
                                     bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                     break;
                                 case "DOC009":
                                     bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                     break;
                                 case "DOC010":
                                     bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                     break;
                                 case "DOC011":
                                     bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                     break;
                                 case "DOC014":
                                     bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                     break;
                                 case "DOC015":
                                     bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                     break;
                                 case "DOC016":
                                     bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                     break;
                                 case "DOC018":
                                     bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                     break;
                                 case "DOC019":
                                     bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                     break;

                             }
                         }
                         Console.WriteLine("Document al forn... un segon!");
                         doc.Print();
                     }

                     if (tipDoc == "T0034")
                     {
                         Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T34.dot");
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC001":
                                     bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC003":
                                     bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                     break;
                                 case "DOC005":
                                     bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                     break;
                                 case "DOC006":
                                     bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                     break;
                                 case "DOC009":
                                     bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                     break;
                                 case "DOC010":
                                     bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                     break;
                                 case "DOC013":
                                     bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                     break;
                                 case "DOC030":
                                     bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                     break;
                                 case "DOC031":
                                     bookmark.Text = rs1.Fields["DOC031"].Value.ToString();
                                     break;
                                 case "DOC032":
                                     bookmark.Text = rs1.Fields["DOC032"].Value.ToString();
                                     break;
                                 case "DOC033":
                                     bookmark.Text = rs1.Fields["DOC033"].Value.ToString();
                                     break;
                                 case "DOC034":
                                     bookmark.Text = rs1.Fields["DOC034"].Value.ToString();
                                     break;
                                 case "DOC036":
                                     bookmark.Text = rs1.Fields["DOC036"].Value.ToString();
                                     break;
                                 case "DOC036x":
                                     bookmark.Text = rs1.Fields["DOC036x"].Value.ToString();
                                     break;
                                 case "DOC037":
                                     bookmark.Text = rs1.Fields["DOC037"].Value.ToString();
                                     break;
                                 case "DOC037x":
                                     bookmark.Text = rs1.Fields["DOC037x"].Value.ToString();
                                     break;
                                 case "DOC038":
                                     bookmark.Text = rs1.Fields["DOC038"].Value.ToString();
                                     break;
                                 case "DOC039":
                                     bookmark.Text = rs1.Fields["DOC039"].Value.ToString();
                                     break;
                                 case "DOC040":
                                     bookmark.Text = rs1.Fields["DOC040"].Value.ToString();
                                     break;
                                 case "DOC041":
                                     bookmark.Text = rs1.Fields["DOC041"].Value.ToString();
                                     break;
                                 case "DOCINC":
                                     bookmark.Text = rs1.Fields["DOCINC"].Value.ToString();
                                     break;
                                 case "OBSER":
                                     bookmark.Text = rs1.Fields["OBSER"].Value.ToString();
                                     break;                                 

                             }
                         }
                         Console.WriteLine("Document al forn... un segon!");
                         doc.Print();
                     }

                     if (tipDoc == "T0088")
                     {
                         Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T88.dot");
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC007":
                                     bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                     break;
                                 case "DOC001":
                                     bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC003":
                                     bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                     break;
                                 case "DOC004":
                                     bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                     break;
                                 case "DOC005":
                                     bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                     break;
                                 case "DOC006":
                                     bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                     break;
                                 case "DOC006x":
                                     bookmark.Text = rs1.Fields["DOC006x"].Value.ToString();
                                     break;
                                 case "DOC008":
                                     bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                     break;
                                 case "DOC014":
                                     bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                     break;
                                 case "DOC015":
                                     bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                     break;
                                 case "DOC016":
                                     bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                     break;
                                 case "DOC018":
                                     bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                     break;
                                 case "DOC019":
                                     bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                     break;
                                 case "DOC020":
                                     bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                     break;
                                  
                             }
                         }
                         Console.WriteLine("Document al forn... un segon!");
                         doc.Print();
                     }

                     if (tipDoc == "T0160")
                     {
                         Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T160.dot");
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC007":
                                     bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                     break;
                                 case "DOC001":
                                     bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC003":
                                     bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                     break;
                                 case "DOC004":
                                     bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                     break;
                                 case "DOC005":
                                     bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                     break;
                                 case "DOC006x":
                                     bookmark.Text = rs1.Fields["DOC006x"].Value.ToString();
                                     break;
                                 case "DOC008":
                                     bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                     break;
                                 case "DOC014":
                                     bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                     break;
                                 case "DOC015":
                                     bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                     break;
                                 case "DOC016":
                                     bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                     break;
                                 case "DOC018":
                                     bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                     break;
                                 case "DOC019":
                                     bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                     break;
                                 case "DOC020":
                                     bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                     break;
                                 case "DOC021":
                                     bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                     break;
                                 case "DOC022":
                                     bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                     break;
                                 case "DOC023":
                                     bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                     break;
                                 case "DOC024":
                                     bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                     break;
                                 case "DOC025":
                                     bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                     break;
                                 case "DOC026":
                                     bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                     break;
                                 case "DOC027":
                                     bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                     break;
                                 case "DOC028":
                                     bookmark.Text = rs1.Fields["DOC028"].Value.ToString();
                                     break;
                                 case "INCI":
                                     bookmark.Text = rs1.Fields["INCI"].Value.ToString();
                                     break;
                                 case "OBSER":
                                     bookmark.Text = rs1.Fields["OBSER"].Value.ToString();
                                     break;

                             }
                         }
                         Console.WriteLine("Document al forn... un segon!");
                         doc.Print();
                     }

                     if (tipDoc == "T0090")
                     {
                         Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T90.dot");
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC021":
                                     bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC003":
                                     bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                     break;
                                 case "DOC004":
                                     bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                     break;
                                 case "DOC005":
                                     bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                     break;
                                 case "DOC006":
                                     bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                     break;
                                 case "DOC007":
                                     bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                     break;
                                 case "DOC008":
                                     bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                     break;
                                 case "DOC009":
                                     bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                     break;
                                 case "DOC010":
                                     bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                     break;
                                 case "DOC012":
                                     bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                     break;
                                 case "DOC013":
                                     bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                     break;
                                 case "DOC014":
                                     bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                     break;
                                 case "DOC015":
                                     bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                     break;
                                 case "DOC016":
                                     bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                     break;
                                 case "DOC017":
                                     bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                     break;
                                 case "DOC018":
                                     bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                     break;
                                 case "DOC019":
                                     bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                     break;
                                 case "DOC020":
                                     bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                     break;
                                 case "DOC022":
                                     bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                     break;
                                 case "DOC023":
                                     bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                     break;
                                   
                             }
                         }
                         Console.WriteLine("Document al forn... un segon!");
                         doc.Print();
                     }

                     if (tipDoc == "T0176")
                     {
                         Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "T176.dot");
                         Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                         foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                         {
                             switch (bookmark.Name)
                             {
                                 case "tipDoc":
                                     bookmark.Text = rs1.Fields["tipDoc"].Value.ToString();
                                     break;
                                 case "DOC030":
                                     bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                     break;
                                 case "DOC002":
                                     bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                     break;
                                 case "DOC003":
                                     bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                     break;
                                 case "DOC004":
                                     bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                     break;
                                 case "DOC005":
                                     bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                     break;
                                 case "DOC006":
                                     bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                     break;
                                 case "DOC007":
                                     bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                     break;
                                 case "DOC008":
                                     bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                     break;
                                 case "DOC009":
                                     bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                     break;
                                 case "DOC012":
                                     bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                     break;
                                 case "DOC013":
                                     bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                     break;
                                 case "DOC014":
                                     bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                     break;
                                 case "DOC015":
                                     bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                     break;
                                 case "DOC016":
                                     bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                     break;
                                 case "DOC017":
                                     bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                     break;
                                 case "DOC018":
                                     bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                     break;
                                 case "DOC019":
                                     bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                     break;
                                 case "DOC020":
                                     bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                     break;
                                 case "DOC021":
                                     bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                     break;
                                 case "DOC024":
                                     bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                     break;
                                 case "DOC025":
                                     bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                     break;
                                 case "DOC026":
                                     bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                     break;
                                 case "DOC027":
                                     bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                     break;
                                 case "DOC029":
                                     bookmark.Text = rs1.Fields["DOC029"].Value.ToString();
                                     break;
                                 case "DOC031":
                                     bookmark.Text = rs1.Fields["DOC031"].Value.ToString();
                                     break;

                             }
                         }
                         Console.WriteLine("Document al forn... un segon!");
                         doc.Print();
                     }

                        if (tipDoc == "T0174" )
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "DOCPB.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC022":
                                        bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                        break;
                                    case "DOC023":
                                        bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                        break;
                                    case "DOC024":
                                        bookmark.Text = rs1.Fields["DOC024"].Value.ToString();
                                        break;
                                    case "DOC025":
                                        bookmark.Text = rs1.Fields["DOC025"].Value.ToString();
                                        break;
                                    case "DOC026":
                                        bookmark.Text = rs1.Fields["DOC026"].Value.ToString();
                                        break;
                                    case "DOC027":
                                        bookmark.Text = rs1.Fields["DOC027"].Value.ToString();
                                        break;
                                    case "DOC028":
                                        bookmark.Text = rs1.Fields["DOC028"].Value.ToString();
                                        break;
                                    case "DOC029":
                                        bookmark.Text = rs1.Fields["DOC029"].Value.ToString();
                                        break;
                                    case "DOC030":
                                        bookmark.Text = rs1.Fields["DOC030"].Value.ToString();
                                        break;
                                    case "DOC031":
                                        bookmark.Text = rs1.Fields["DOC031"].Value.ToString();
                                        break;
                                    case "DOC032":
                                        bookmark.Text = rs1.Fields["DOC032"].Value.ToString();
                                        break;
                                    case "DOC033":
                                        bookmark.Text = rs1.Fields["DOC033"].Value.ToString();
                                        break;
                                    case "DOC034":
                                        bookmark.Text = rs1.Fields["DOC034"].Value.ToString();
                                        break;
                                    case "DOC035":
                                        bookmark.Text = rs1.Fields["DOC035"].Value.ToString();
                                        break;
                                    case "DOC036":
                                        bookmark.Text = rs1.Fields["DOC036"].Value.ToString();
                                        break;
                                    case "DOC037":
                                        bookmark.Text = rs1.Fields["DOC037"].Value.ToString();
                                        break;
                                    case "DOC038":
                                        bookmark.Text = rs1.Fields["DOC038"].Value.ToString();
                                        break;
                                    case "DOC039":
                                        bookmark.Text = rs1.Fields["DOC039"].Value.ToString();
                                        break;
                                    case "DOC040":
                                        bookmark.Text = rs1.Fields["DOC040"].Value.ToString();
                                        break;
                                    case "DOC041":
                                        bookmark.Text = rs1.Fields["DOC041"].Value.ToString();
                                        break;
                                    case "DOC042":
                                        bookmark.Text = rs1.Fields["DOC042"].Value.ToString();
                                        break;
                                    case "DOC043":
                                        bookmark.Text = rs1.Fields["DOC043"].Value.ToString();
                                        break;
                                    case "DOC044":
                                        bookmark.Text = rs1.Fields["DOC044"].Value.ToString();
                                        break;
                                    case "DOC045":
                                        bookmark.Text = rs1.Fields["DOC045"].Value.ToString();
                                        break;
                                    case "DOC046":
                                        bookmark.Text = rs1.Fields["DOC046"].Value.ToString();
                                        break;
                                    case "DOC047":
                                        bookmark.Text = rs1.Fields["DOC047"].Value.ToString();
                                        break;
                                    case "DOC048":
                                        bookmark.Text = rs1.Fields["DOC049"].Value.ToString();
                                        break;
                                    case "DOC050":
                                        bookmark.Text = rs1.Fields["DOC050"].Value.ToString();
                                        break;
                                    case "DOC051":
                                        bookmark.Text = rs1.Fields["DOC051"].Value.ToString();
                                        break;
                                    case "DOC052":
                                        bookmark.Text = rs1.Fields["DOC052"].Value.ToString();
                                        break;
                                    case "DOC053":
                                        bookmark.Text = rs1.Fields["DOC053"].Value.ToString();
                                        break;
                                    case "DOC054":
                                        bookmark.Text = rs1.Fields["DOC054"].Value.ToString();
                                        break;
                                    case "DOC055":
                                        bookmark.Text = rs1.Fields["DOC055"].Value.ToString();
                                        break;
                                    case "DOC056":
                                        bookmark.Text = rs1.Fields["DOC056"].Value.ToString();
                                        break;
                                    case "DOC057":
                                        bookmark.Text = rs1.Fields["DOC057"].Value.ToString();
                                        break;
                                    case "DOC058":
                                        bookmark.Text = rs1.Fields["DOC058"].Value.ToString();
                                        break;
                                   

                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }

                        if (tipDoc == "T074A")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "DECJP.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                   
                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }

                        if (tipDoc == "T074B")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "DECRCA1.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                        break;
                                    case "DOC022":
                                        bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                        break;                                   
                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }

                        if (tipDoc == "T074C")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "DECRCA2.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;                                    
                                    case "DOC018":
                                         bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    case "DOC021":
                                        bookmark.Text = rs1.Fields["DOC021"].Value.ToString();
                                        break;
                                    case "DOC022":
                                        bookmark.Text = rs1.Fields["DOC022"].Value.ToString();
                                        break;
                                    case "DOC023":
                                        bookmark.Text = rs1.Fields["DOC023"].Value.ToString();
                                        break;                                   
                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }

                        if (tipDoc == "T0122")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "DECPB.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                    
                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }

                        if (tipDoc == "T0106")
                        {
                            Aspose.Words.Document doc = new Aspose.Words.Document(docTemplate + "DECPBP.dot");
                            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                            foreach (Aspose.Words.Bookmark bookmark in doc.Range.Bookmarks)
                            {
                                switch (bookmark.Name)
                                {
                                    case "DOC001":
                                        bookmark.Text = rs1.Fields["DOC001"].Value.ToString();
                                        break;
                                    case "DOC002":
                                        bookmark.Text = rs1.Fields["DOC002"].Value.ToString();
                                        break;
                                    case "DOC003":
                                        bookmark.Text = rs1.Fields["DOC003"].Value.ToString();
                                        break;
                                    case "DOC004":
                                        bookmark.Text = rs1.Fields["DOC004"].Value.ToString();
                                        break;
                                    case "DOC005":
                                        bookmark.Text = rs1.Fields["DOC005"].Value.ToString();
                                        break;
                                    case "DOC006":
                                        bookmark.Text = rs1.Fields["DOC006"].Value.ToString();
                                        break;
                                    case "DOC007":
                                        bookmark.Text = rs1.Fields["DOC007"].Value.ToString();
                                        break;
                                    case "DOC008":
                                        bookmark.Text = rs1.Fields["DOC008"].Value.ToString();
                                        break;
                                    case "DOC009":
                                        bookmark.Text = rs1.Fields["DOC009"].Value.ToString();
                                        break;
                                    case "DOC010":
                                        bookmark.Text = rs1.Fields["DOC010"].Value.ToString();
                                        break;
                                    case "DOC011":
                                        bookmark.Text = rs1.Fields["DOC011"].Value.ToString();
                                        break;
                                    case "DOC012":
                                        bookmark.Text = rs1.Fields["DOC012"].Value.ToString();
                                        break;
                                    case "DOC013":
                                        bookmark.Text = rs1.Fields["DOC013"].Value.ToString();
                                        break;
                                    case "DOC014":
                                        bookmark.Text = rs1.Fields["DOC014"].Value.ToString();
                                        break;
                                    case "DOC015":
                                        bookmark.Text = rs1.Fields["DOC015"].Value.ToString();
                                        break;
                                    case "DOC016":
                                        bookmark.Text = rs1.Fields["DOC016"].Value.ToString();
                                        break;
                                    case "DOC017":
                                        bookmark.Text = rs1.Fields["DOC017"].Value.ToString();
                                        break;
                                    case "DOC018":
                                        bookmark.Text = rs1.Fields["DOC018"].Value.ToString();
                                        break;
                                    case "DOC019":
                                        bookmark.Text = rs1.Fields["DOC019"].Value.ToString();
                                        break;
                                    case "DOC020":
                                        bookmark.Text = rs1.Fields["DOC020"].Value.ToString();
                                        break;
                                   
                                }
                            }
                            Console.WriteLine("Document al forn... un segon!");
                            doc.Print();
                        }
                        //doc.Save(strFileTarget, Aspose.Words.SaveFormat.Pdf);

                        rs1.Close();
                        conn.Close();
 }
                
                else {

                    Console.WriteLine("falten paràmetres..");
              }
}
              catch (Exception e) {
           
                Console.WriteLine(e.Message);
                //var message = new UTF8Encoding(true).GetBytes(e.Message);
                //fileLog.Write(message, 0, message.Length);
                //fileLog.Flush();

               
            }

            finally {
                //var message = new UTF8Encoding(true).GetBytes("ok");
                //fileLog.Write(message, 0, message.Length);
                //fileLog.Flush();
              
            }
                      
                
          
        
                
        }
    }
}
 

