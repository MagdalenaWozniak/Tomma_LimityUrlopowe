using CDNBase;
using CDNPrac;
using OP_KALB2Lib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace LimityUrlopowe
{
    class Program
    {

        public static string Operator = "ADMIN";
        public static string Haslo = "NE!w0rk12";
        public static string Firma = "ERP-TOMMA-KADRY"; //"ERP_TOMMA_KADRY_TEST";
        //  public static string Baza = "[CDN_PREZENTACJA_KH_1].[CDN].";
        protected static CDNBase.Application Application = null;
        protected static ILogin Login = null;


        static void Main(string[] args)
        {


            TextWriter tw2 = new StreamWriter(@"C:\OneSupport\LimityUrlopowe\log.txt", append: true);
            tw2.WriteLine();
            tw2.WriteLine();
            tw2.WriteLine();
            tw2.WriteLine();
            tw2.WriteLine("Hello " + DateTime.Now.ToString("f"));
            tw2.WriteLine();
            tw2.WriteLine();

            WykonajProgram(tw2);

            Firma = "ERP-BONUS2-KADRY"; //"ERP_BONUS2_KADRY_TEST";

            WykonajProgram(tw2);








                tw2.Close();
            }

        static void WykonajProgram(TextWriter tw2)
        {
            
            try
            {
                LogowanieAutomatyczne();
            }
            catch (COMException eo)
            {
                if (eo.InnerException != null)
                    tw2.WriteLine("Błąd logowania: " + eo.InnerException.Message);
                else
                    tw2.WriteLine("Błąd logowania: " + eo.Message);
                return;
            }
            catch (Exception ei)
            {
                if (ei.InnerException != null)
                    tw2.WriteLine("Błąd logowania: " + ei.InnerException.Message);
                else
                    tw2.WriteLine("Błąd logowania: " + ei.Message);
                //MessageBox.Show("Zalogowano ", "Komunikat",
                // MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {


                CDNBase.AdoSession Sesja = Login.CreateSession();

                /*string idQuery = "SELECT PRa_PraId FROM cdn.PracKod join CDN.PracEtaty on PRE_PraId = Pra_PraID join CDN.PracIdx on Pri_PraID = Pra_PraID " +

                                                     " where PRa_Archiwalny = 0 and ISNULL(Pre_ZatrudnionyDo, convert(datetime,'1900-12-31', 121))=convert(datetime,'2999-12-31', 121) " +
                                                     " and Pri_Typ = 10  and ISNULL(PRI_DataOd, convert(datetime,'1900-12-31', 121)) <= GETDATE() and ISNULL(PRI_DataDo, convert(datetime,'1900-12-31', 121)) >= GETDATE()" +
                                                     " group by PRA_PraId";*/

                string idQuery = "SELECT PRa_PraId FROM cdn.PracKod join CDN.PracEtaty on PRE_PraId = Pra_PraID join CDN.PracIdx on Pri_PraID = Pra_PraID " +

                                                    " where  PRI_Typ = 10  AND PRI_DataDo >= Convert(DATETIME, GETDATE(), 121) AND PRE_ZatrudnionyDo >= Convert(DATETIME, GETDATE(), 121)  AND PRE_TypPracownika = 0 AND PRI_Archiwalny< 1" +
                                                    " and Convert(DATETIME, GETDATE(),121) >= PRE_DataOd   and Convert(DATETIME, GETDATE(),121) <= PRE_DataDo "+
                                                    " group by PRA_PraId";




               
                object dummy = Type.Missing;

                var idList = Sesja.Connection.Execute(idQuery, out dummy, 0);

                idList.MoveFirst();
                while(!idList.EOF)
                {
                    IAdoSession session2 = Login.CreateSession();

                    IPracownik pracownik = (IPracownik)session2.CreateObject("CDN.Pracownicy", idList.Fields[0].Value.ToString());


                    ISerwisLimitow serwisLimitow = (ISerwisLimitow)session2.CreateObject("CDNKalb.SerwisLimitow");

                    string czyLimitQuery = "if exists(SELECT * FROM cdn.PracLimit WHERE PLN_Rok = 2024 and PLN_PraId = 1 and PLN_LnbId = 1) SELECT 'TAK' else SELECT 'NIE' ";
                    var czyLimit = session2.Connection.Execute(czyLimitQuery, out dummy, 0);
                    if(!czyLimit.EOF)
                    {
                        if(czyLimit.Fields[0].Value.ToString()=="NIE")
                        {
                            tw2.WriteLine("Dodanie limitu pracownikowi: " + pracownik.ImieNazwisko);
                            serwisLimitow.DodajLimit(pracownik.ID, 1, DateTime.Today.Year, 1);
                            
                        }
                        else
                        {
                            tw2.WriteLine("Aktualizacja limitu pracownikowi: " + pracownik.ImieNazwisko);
                            serwisLimitow.SprawdzLimity(pracownik, DateTime.Today.Year);
                        }
                    }


                    session2.Save();

                    idList.MoveNext();
                }

               /* ICollection pracownicy = Sesja.CreateObject("CDN.Pracownicy", null);
                List<int> idList = pracownicy.Cast<IPracownik>().ToList().Select(x => x.ID).ToList();

                

                foreach (int id in idList)

                {
                    IAdoSession session2 = Login.CreateSession();

                    IPracownik pracownik = (IPracownik)session2.CreateObject("CDN.Pracownicy", id.ToString());
                    try
                    {
                        object dummy = Type.Missing;
                        //tworzymy nową sesję i w jej obrębie pracownika oraz jego limity
                        string czyEtatWuery = "if exists(SELECT 1 FROM cdn.PracKod join CDN.PracEtaty on PRE_PraId = Pra_PraID join CDN.PracIdx on Pri_PraID = Pra_PraID " +

                                                     " where PRa_Archiwalny = 0 and ISNULL(Pre_ZatrudnionyDo, CAST('1900-12-31' as datetime))=CAST('2999-12-31' as datetime) and PRA_PraId = " + pracownik.ID +
                                                     " and Pri_Typ = 10  and ISNULL(PRI_DataOd, CAST('1900-01-01' as datetime)) <= GETDATE() and ISNULL(PRI_DataDo, CAST('1900-01-01' as datetime)) >= GETDATE()) " +
                                                     "SELECT 'TAK' else SELECT 'NIE'";
                        //tw2.WriteLine(czyEtatWuery);
                        var czyEtat=session2.Connection.Execute(czyEtatWuery, out dummy, 0);

                        if (!czyEtat.EOF)
                        {
                            tw2.WriteLine("Wynik Query: " + czyEtat.Fields[0].Value.ToString() +" dla pracownika "+pracownik.ImieNazwisko);
                            if (czyEtat.Fields[0].Value.ToString()=="TAK")
                            {


                                ISerwisLimitow serwisLimitow = (ISerwisLimitow)session2.CreateObject("CDNKalb.SerwisLimitow");

                                Console.WriteLine("Dodanie limitów pracownikowi: " + pracownik.ImieNazwisko);

                                foreach (IPracLimityElem pracLimityElem in pracownik.Limity)

                                {

                                    serwisLimitow.SprawdzLimit(pracLimityElem);

                                    //ileLimit++;

                                }


                            }
                        }

                        session2.Save();

                        //oznaczymy obiekty powstałe w obrębie session2 że mogą być zwolnione

                        session2.Dispose();
                    }
                    catch(Exception ex)
                    {
                        tw2.WriteLine("Błąd aktualizacji limitu pracownika " + pracownik.ImieNazwisko + " :" + ex.Message);
                        Console.WriteLine("Błąd aktualizacji limitu pracownika " + pracownik.ImieNazwisko + " :" + ex.Message);
                        session2.Dispose();
                    }

                }*/


                /*
                foreach (IPracownik pracownik in pracownicy)

                {
                    
                    foreach (IPracLimityElem item in pracownik.Limity)

                    {
                       
                        Sesja = Login.CreateSession();

                        var serwisLimitow = (ISerwisLimitow)Sesja.CreateObject("CDNKalb.SerwisLimitow", null);

                        serwisLimitow.SprawdzLimit(item);
                        //serwisLimitow.
                        Sesja.Save();
                        Sesja.Dispose();
                      

                    }


                }
                */
                tw2.WriteLine("Zakończono przetwarzanie pracowników w " + Firma);
                tw2.WriteLine();
            }
            catch(Exception en)
            {
                tw2.WriteLine("Błąd przeglądu pracowników: " + en.Message);
            }

            //dodawanieDekretu(dataTable, Sesja);
            WyjscieZApki("done", tw2);



            //zapisujemy
            //Sesja.Save();



        }
        

            static void LogowanieAutomatyczne()
            {
                object[] hPar = new object[] {
                         0,  0,   0,  0,  0,   0,  0,    1,   0,   0,   0,   0,   0,   0,  0,  0, 0 };    // do jakich modułów się logujemy
                /* Kolejno: KP, KH, KHP, ST, FA, MAG, PK, PKXL, CRM, ANL, DET, BIU, SRW, ODB, KB, KBP, HAP*/
                // katalog, gdzie jest zainstalowana Optima (bez ustawienia tej zmiennej nie zadziała, chyba że program odpalimy z katalogu O!)
                System.Environment.CurrentDirectory = @"C:\Program Files (x86)\Comarch ERP Optima";
                // tworzymy nowy obiekt apliakcji
                Application = new CDNBase.Application();
                //Application.HASPServerCommandLine = "41941-FE3B2-002::5000139271";
                // blokujemy
                Application.LockApp(1);//(513, 5000, null, null, null, null);
                                       // logujemy się do podanej Firmy, na danego operatora, do podanych modułów
                Login = Application.Login(vUser: Operator, vChk: Haslo, vFirm: Firma, _ModulKP: hPar[0], _ModulKH: hPar[1], _ModulKHP: hPar[2], _ModulST: hPar[3], _ModulFA: hPar[4], _ModulMAG: hPar[5], _ModulPK: hPar[6], _ModulPKXL: hPar[7], _ModulCRM: hPar[8], _ModulANL: hPar[9], _ModulDET: hPar[10], _ModulBIU: hPar[11], _ModulSRW: hPar[12], _ModulOBD: hPar[13], _ModulKB: hPar[14], _ModulKBP: hPar[15], _ModulHAP: hPar[16]);

            }

            static void Wylogowanie()
            {
                // niszczymy Login
                Login = null;
                // odblokowanie (wylogowanie) O!
                Application.UnlockApp();
                // niszczymy obiekt Aplikacji
                Application = null;
            }

            static void WyjscieZApki(string a, TextWriter tw2)
            {
                Wylogowanie();
                if (a == "done")
                {
                    //tw2.WriteLine("Zakończono przetwarzanie pracowników w " + Firma);
                }
                else
                {
                    tw2.WriteLine("Generowanie limitów urlopowych zakończyło się niepowodzeniem.");
                }

                return;
            }
        }
}
