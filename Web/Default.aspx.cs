using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.OleDb;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Runtime.Remoting.Messaging;
using System.Collections;

public partial class _Default : System.Web.UI.Page
{

    public OleDbCommand Connection_Maker(string path, string select_from_page) //kod tekrarı olmaması için bu fonksiyonu yazdım.
    {
        string dosya = path;
        string connString = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 12.0;",
            Server.MapPath(dosya));

        OleDbConnection conn = new OleDbConnection(connString);
        conn.Open();
        OleDbCommand cmd = new OleDbCommand("SELECT * FROM [" + select_from_page + "$]", conn);

        return cmd;
    }
   
    public double Vergisiz_fiyat_hesapla(int dolar_fiyat, double kur)
    {
        double vergisiz_fiyat = dolar_fiyat * kur;
        return vergisiz_fiyat;
    }
    public double Normal_otv_hesapla(int dolar_fiyat, double kur)
    {
        double vergisiz_fiyat = Vergisiz_fiyat_hesapla(dolar_fiyat, kur);
        double otv_miktar_tl;
        if (vergisiz_fiyat < 70000) { otv_miktar_tl = vergisiz_fiyat * 0.45; }
        else if (vergisiz_fiyat >= 70000 && vergisiz_fiyat < 120000) { otv_miktar_tl = vergisiz_fiyat * 0.50; }
        else { otv_miktar_tl = vergisiz_fiyat * 0.60; }

        return otv_miktar_tl;
    }
    public double Normal_kdv_hesapla(int dolar_fiyat, double kur)
    {
        double otv_miktar = Normal_otv_hesapla(dolar_fiyat, kur);
        double kdv_miktar_tl = otv_miktar * 0.18;
        return kdv_miktar_tl;

    }

    public double Normal_toplam_fiyat_hesapla(int dolar_fiyat, double kur)
    {
        double toplam_fiyat = Vergisiz_fiyat_hesapla(dolar_fiyat, kur) + Normal_otv_hesapla(dolar_fiyat, kur) + Normal_kdv_hesapla(dolar_fiyat, kur);
        return toplam_fiyat;
    }
    public double Engelli_kdv_hesapla(int dolar_fiyat, double kur)
    {
        double vergisiz_fiyat = Vergisiz_fiyat_hesapla(dolar_fiyat, kur);
        double engelli_kdv_miktar_tl = vergisiz_fiyat * 0.18;

        return engelli_kdv_miktar_tl; // Fonksiyonun 0 döndürmesi demek engelli bireyin bu aracı alamadığı anlamına gelir.;
    }

    public double Engelli_toplam_fiyat_hesapla(int dolar_fiyat, double kur)
    {
        double engelli_toplam_fiyat = Vergisiz_fiyat_hesapla(dolar_fiyat, kur) + Engelli_kdv_hesapla(dolar_fiyat, kur);

        if (engelli_toplam_fiyat < 303200) { return engelli_toplam_fiyat; }
        return 0; // Fonksiyonun 0 döndürmesi demek engelli bireyin bu aracı alamadığı anlamına gelir.;
    }
    protected void Page_Load(object sender, EventArgs e)
    {
        OleDbCommand cmd2 = Connection_Maker("~/App_Data/WVGolf_Fiyat.xlsx", "Sheet1");
        OleDbDataReader reader2 = cmd2.ExecuteReader();

        var yıllar = new List<string>();
        var fiyatlar = new List<string>();
        var fiyatlar_int = new List<int>(); //fiatlar üzerinde hesaplamalar yapmak için int listeye attım.
        int val6;
        var tarihler = new List<string>(); //excel den oku ve bu listeye tarihleri at.
        var kurlar = new List<string>(); //excel den oku ve bu listeye kurları at.
        var kurlar_double = new List<double>(); //hesaplamalar için kurları double listede tut.
        var tarih_yillar = new List<string>(); // wvgolf_fiyatları dosyasından yılları bu listeye at
        
        double val3;

        while (reader2.Read())
        {
            string val4 = reader2[0].ToString();

            string val5 = reader2[1].ToString();

            if (val5 == "")
            {
                val5 = "10,000"; 


            }

            Regex reg = new Regex("[$,]");
            string price_formatted = reg.Replace(val5, string.Empty);
            val6 = int.Parse(price_formatted); //dolar ücretleri formatlayıp fiyat listeme attım.
            yıllar.Add(val4);
            fiyatlar.Add(val5);
            fiyatlar_int.Add(val6);

        }
        reader2.Close(); //fiyatlar excel dosyasıyla işim bitti. reader' ı kapattım.
        OleDbCommand cmd = Connection_Maker("~/App_Data/dolar-kurları.xlsx", "EVDS");
        OleDbDataReader reader = cmd.ExecuteReader();


        DataTable tarih_kur_table = new DataTable(); //tarih ve kurları bir DataTable yaptım.
        tarih_kur_table.Columns.Add("id", typeof(string));
        tarih_kur_table.TableName = "tarihler ve kurlar";
        tarih_kur_table.Columns.Add("tarih", typeof(string));
        tarih_kur_table.Columns.Add("kur ($)", typeof(string));
        
        int id2 = 0;

        while (reader.Read())
        {
            
            string val1 = reader[0].ToString();
            if (val1 != string.Empty) { 
                string tarih_yil = val1.Substring(6);
                tarih_yillar.Add(tarih_yil); //tarihlerin yıl kısmını bir listede topladım.
                
                string val2 = reader[1].ToString();
            
                string onceki_kur;
            if (val2 == "")
            {
                if (val1 == "18-04-2010") // ilk tarih için bir sonraki gümün kurunu default olarak atadım.
                {
                    val2 = "1,4685";
                        
                    tarih_kur_table.Rows.Add(id2,val1, val2);
                    kurlar.Add(val2);
                    double ilk_kur_double = Convert.ToDouble(val2);
                    kurlar_double.Add(ilk_kur_double);
                    tarihler.Add(val1);
                    tarih_yillar.Add(val1.Substring(6));
                    
                    id2++;
                    continue;
                        
                }

                // aşağıdaki satırda boş olan kurlara bir önceki günün kur değerini atadım.
                onceki_kur = kurlar.TakeWhile(x => x != val2).DefaultIfEmpty(kurlar[kurlar.Count - 1]).LastOrDefault();
                val2 = onceki_kur;

            }
            val3 = Convert.ToDouble(val2); 

            tarihler.Add(val1);
            kurlar.Add(val2); //Bu liste string tipinde çünkü gridview'da kullanılacak. 
                                //Çoğu listenin hem string hem de double veya int olarak tutmanın nedeni de bu.
            
            kurlar_double.Add(val3);
            
            
            string id_str2 = id2.ToString();
            tarih_kur_table.Rows.Add(id2,val1, val2);
            id2++;
            }
            else
            {
                break;
            }
        }
        

        string[] yıllar_asarray = yıllar.ToArray(); //yılları  array yap.
        int[] fiyatlar_int_asarray = fiyatlar_int.ToArray(); //yıla göre fiyat ataması için bir array yaptım.
        string[] tarih_yillar_asarray = tarih_yillar.ToArray();
        var yila_gore_araba = new List<int>(); //yıla göre araba fiyatları
        
        double[] kurlar_double_asarray = kurlar_double.ToArray(); 
        int araba_fiyati;
        
        for (int i = 0; i < tarih_yillar_asarray.Length; i++) //her tarih için araba fiyatını belirler.
        {

            if (tarih_yillar_asarray[i] == "2010") { araba_fiyati = 14850; yila_gore_araba.Add(araba_fiyati); }
            else if (tarih_yillar_asarray[i] == "2011") { araba_fiyati = 19030; yila_gore_araba.Add(araba_fiyati); }
            else if (tarih_yillar_asarray[i] == "2012") { araba_fiyati = 21890; yila_gore_araba.Add(araba_fiyati); }
            else if (tarih_yillar_asarray[i] == "2013") { araba_fiyati = 24860; yila_gore_araba.Add(araba_fiyati); }
            else if (tarih_yillar_asarray[i] == "2014") { araba_fiyati = 35420; yila_gore_araba.Add(araba_fiyati); }
            else if (tarih_yillar_asarray[i] == "2015") { araba_fiyati = 41910; yila_gore_araba.Add(araba_fiyati); }
            else if (tarih_yillar_asarray[i] == "2016") { araba_fiyati = 45430; yila_gore_araba.Add(araba_fiyati); }
            else if (tarih_yillar_asarray[i] == "2017") { araba_fiyati = 49720; yila_gore_araba.Add(araba_fiyati); }
            else if (tarih_yillar_asarray[i] == "2018") { araba_fiyati = 54120; yila_gore_araba.Add(araba_fiyati); }
            else if (tarih_yillar_asarray[i] == "2019") { araba_fiyati = 57310; yila_gore_araba.Add(araba_fiyati); }
            else if (tarih_yillar_asarray[i] == "2020") { araba_fiyati = 57990; yila_gore_araba.Add(araba_fiyati); }
        }

        int[] yila_gore_araba_asarray = yila_gore_araba.ToArray(); //yıla göre araba fiyatları array tipinde
       

        
        double vergisiz_fiyat, normal_otv, normal_kdv, normal_top, engelli_kdv, engelli_top;
            
        DataTable fiyatlar_listesi_table = new DataTable(); //bütün fiyat bileşenlerini tek datatable'a ekledim.
        fiyatlar_listesi_table.Columns.Add("id", typeof(string));
        fiyatlar_listesi_table.TableName = "fiyatlar_listesi_tablosu";
        fiyatlar_listesi_table.Columns.Add("vergisiz fiyat (tl)", typeof(string));
        fiyatlar_listesi_table.Columns.Add("normal otv miktarı (tl)", typeof(string));
        fiyatlar_listesi_table.Columns.Add("normal kdv miktarı (tl)", typeof(string));
        fiyatlar_listesi_table.Columns.Add("normal fiyat toplam (tl)", typeof(string));
        fiyatlar_listesi_table.Columns.Add("engelli kdv miktarı (tl)", typeof(string));
        fiyatlar_listesi_table.Columns.Add("engelli fiyat toplam (tl)", typeof(string));
        
        for (int a=0; a< yila_gore_araba_asarray.Length; a++)
            {
            for (int b = 0; b < kurlar_double_asarray.Length; b++)
            {
                if (a == b) { 
                    int usd_fiyat = yila_gore_araba_asarray[a];
                    double usd_kur = kurlar_double_asarray[b];
                    vergisiz_fiyat = Vergisiz_fiyat_hesapla(usd_fiyat, usd_kur);
                    normal_otv = Normal_otv_hesapla(usd_fiyat, usd_kur);
                    normal_kdv = Normal_kdv_hesapla(usd_fiyat, usd_kur);
                    normal_top = Normal_toplam_fiyat_hesapla(usd_fiyat, usd_kur);
                    engelli_kdv = Engelli_kdv_hesapla(usd_fiyat, usd_kur);
                    engelli_top = Engelli_toplam_fiyat_hesapla(usd_fiyat, usd_kur);
                    int id = a;
                    
                    string vergisiz_str = vergisiz_fiyat.ToString("#.####"); //double değerin virgülden sonraki 4 basamağını aldım.
                    string norm_otv_str = normal_otv.ToString("#.####"); //double değerin virgülden sonraki 4 basamağını aldım.
                    string norm_kdv_str = normal_kdv.ToString("#.####"); //double değerin virgülden sonraki 4 basamağını aldım.
                    string norm_top_str = normal_top.ToString("#.####"); //double değerin virgülden sonraki 4 basamağını aldım.
                    string eng_kdv_str = engelli_kdv.ToString("#.####"); //double değerin virgülden sonraki 4 basamağını aldım.
                    string eng_top_str = engelli_top.ToString("#.####"); //double değerin virgülden sonraki 4 basamağını aldım.
                    string id_str = id.ToString();
                    
                    fiyatlar_listesi_table.Rows.Add(id_str, vergisiz_str, norm_otv_str, norm_kdv_str, norm_top_str, eng_kdv_str, eng_top_str);
                    }

                }
            }
        tarih_kur_table.Constraints.Add("pk", tarih_kur_table.Columns[0], true); //tarih_kur ve fiyatlar_listesi_table ı tek tablo yapmak için pk ekleyip merge ettim.
        fiyatlar_listesi_table.Constraints.Add("pk", fiyatlar_listesi_table.Columns[0], true); //tarih_kur ve fiyatlar_listesi_table ı tek tablo yapmak için pk ekleyip merge ettim.
        tarih_kur_table.Merge(fiyatlar_listesi_table);
        tarih_kur_table.PrimaryKey = null;
        tarih_kur_table.Columns.Remove("id"); //id sütununu birleştirilmiş tablodan kaldırdım.

        
        reader.Close();
            
            gv.DataSource = tarih_kur_table; //birleştirilmiş tabloyu gridview ile ekrana bastırdım.
            gv.DataBind();
            this.gv.Visible = true;
            
            

    }
}
