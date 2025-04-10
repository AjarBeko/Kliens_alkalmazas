using Hotcakes.CommerceDTO.v1.Client;
using Hotcakes.CommerceDTO.v1;
using Hotcakes.CommerceDTO.v1.Orders;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Hotcakes.Web;
using Novacode;
using System.IO;
using BarcodeLib;
using Hotcakes.Web.Barcodes;


namespace Szamla_Kliensalkalmazas
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private static Api apiHivas()
        {
            string url = "http://rendfejl10000.northeurope.cloudapp.azure.com:8080";
            string kulcs = "1-35939070-0c15-468a-b98c-8da71a3e96ca";
            Api proxy = new Api(url, kulcs);
            return proxy;
        }


        //RENDELÉSEK LEKÉRÉSE
        private void button1_Click(object sender, EventArgs e)
        {
            Api proxy = apiHivas();

            var response = proxy.OrdersFindAll();

            if (response == null || response.Content == null || response.Content.Count == 0)
            {
                MessageBox.Show("Nem sikerült lekérni a rendeléseket vagy nincs adat.");
                return;
            }

            DataTable tabla = new DataTable();

            tabla.Columns.Add("OrderNumber");
            tabla.Columns.Add("OrderBvin");
            tabla.Columns.Add("OrderDate");             
            tabla.Columns.Add("UserEmail");
            tabla.Columns.Add("TotalGrand");
            tabla.Columns.Add("BillingName");
            tabla.Columns.Add("BillingStreet");
            tabla.Columns.Add("BillingCity");
            tabla.Columns.Add("ShippingName");
            tabla.Columns.Add("ShippingStreet");
            tabla.Columns.Add("ShippingCity");

            foreach (var order in response.Content)
            {
                var row = tabla.NewRow();

                row["OrderNumber"] = order.OrderNumber;
                row["OrderBvin"] = order.bvin;
                row["OrderDate"] = order.TimeOfOrderUtc.ToLocalTime().ToString("yyyy.MM.dd HH:mm");  // ⬅️ dátum
                row["UserEmail"] = order.UserEmail;
                row["TotalGrand"] = order.TotalGrand.ToString("0.00");

                if (order.BillingAddress != null)
                {
                    row["BillingName"] = $"{order.BillingAddress.FirstName} {order.BillingAddress.LastName}";
                    row["BillingStreet"] = order.BillingAddress.Line1;
                    row["BillingCity"] = order.BillingAddress.City;
                }

                if (order.ShippingAddress != null)
                {
                    row["ShippingName"] = $"{order.ShippingAddress.FirstName} {order.ShippingAddress.LastName}";
                    row["ShippingStreet"] = order.ShippingAddress.Line1;
                    row["ShippingCity"] = order.ShippingAddress.City;
                }

                tabla.Rows.Add(row);
            }

            dataGridView1.DataSource = tabla;
            dataGridView1.Columns["OrderBvin"].Visible = false;

            textBox1.Text = tabla.Rows.Count.ToString();
        }

        //TERMÉKEKHEZ TARTOZÓ RENDELÉSEK LEKÉRÉSE

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("Előbb válassz ki egy rendelést!");
                return;
            }

            //OrderBvin lekérése
            string orderBvin = dataGridView1.SelectedRows[0].Cells["OrderBvin"].Value.ToString();

            Api proxy = apiHivas();

            var response = proxy.OrdersFind(orderBvin);

            if (response == null || response.Content == null)
            {
                MessageBox.Show("Nem sikerült lekérni a rendelés részleteit.");
                return;
            }

            var order = response.Content;

            //Termékek táblázat létrehozása
            DataTable termekTabla = new DataTable();
            termekTabla.Columns.Add("Termék neve");
            termekTabla.Columns.Add("Mennyiség");
            termekTabla.Columns.Add("Egységár");
            termekTabla.Columns.Add("Bruttó összesen");

            foreach (var item in order.Items)
            {
                DataRow row = termekTabla.NewRow();
                row["Termék neve"] = item.ProductName;
                row["Mennyiség"] = item.Quantity;
                row["Egységár"] = item.BasePricePerItem.ToString("0.00");
                row["Bruttó összesen"] = item.LineTotal.ToString("0.00");
                termekTabla.Rows.Add(row);
            }

            dataGridView2.DataSource = termekTabla;
        }

        //SZÁMLA GENERÁLÁSA

        private void GeneralSzamla(OrderDTO order)
        {
            string sablonPath = Path.Combine(Application.StartupPath, "PixelPress_SzamlaSablon.docx");
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string kimenetiPath = Path.Combine(desktopPath, $"Szamla_{order.OrderNumber}.docx");

            var doc = DocX.Load(sablonPath);

            //Random sorszám mezők
            Random rnd = new Random();
            doc.ReplaceText("{{randomszam1}}", rnd.Next(100000, 999999).ToString());
            doc.ReplaceText("{{randomszam2}}", rnd.Next(100000, 999999).ToString());
            doc.ReplaceText("{{randomszam3}}", rnd.Next(100000, 999999).ToString());

            //Dátumok
            DateTime kelte = order.TimeOfOrderUtc.ToLocalTime();
            DateTime teljesites = kelte.AddDays(2);
            doc.ReplaceText("{{OrderDate}}", kelte.ToString("yyyy.MM.dd"));
            doc.ReplaceText("{{OrderDate2}}", teljesites.ToString("yyyy.MM.dd"));

            //Számlázási adatok
            doc.ReplaceText("{{BillingName}}", $"{order.BillingAddress.FirstName} {order.BillingAddress.LastName}");
            doc.ReplaceText("{{BillingCity}}", order.BillingAddress.City);
            doc.ReplaceText("{{BillingStreet}}", order.BillingAddress.Line1);
            doc.ReplaceText("{{Iranyito}}", order.BillingAddress.PostalCode ?? "0000");

            //Táblázat létrehozása
            var table = doc.AddTable(order.Items.Count + 1, 7);
            table.Design = TableDesign.TableGrid;

            table.Rows[0].Cells[0].Paragraphs[0].Append("Megnevezés").Bold();
            table.Rows[0].Cells[1].Paragraphs[0].Append("Termékkód").Bold();
            table.Rows[0].Cells[2].Paragraphs[0].Append("Egységár").Bold();
            table.Rows[0].Cells[3].Paragraphs[0].Append("Mennyiség").Bold();
            table.Rows[0].Cells[4].Paragraphs[0].Append("ÁFA").Bold();
            table.Rows[0].Cells[5].Paragraphs[0].Append("Nettó").Bold();
            table.Rows[0].Cells[6].Paragraphs[0].Append("Bruttó").Bold();

            decimal osszNetto = 0, osszAfa = 0, osszBrutto = 0;

            for (int i = 0; i < order.Items.Count; i++)
            {
                var item = order.Items[i];
                decimal unitBrutto = item.BasePricePerItem;
                decimal unitNetto = unitBrutto / 1.27m;
                decimal unitAfa = unitBrutto - unitNetto;

                decimal netto = unitNetto * item.Quantity;
                decimal afa = unitAfa * item.Quantity;
                decimal brutto = unitBrutto * item.Quantity;

                table.Rows[i + 1].Cells[0].Paragraphs[0].Append(item.ProductName);
                table.Rows[i + 1].Cells[1].Paragraphs[0].Append("termek" + (i + 1));
                table.Rows[i + 1].Cells[2].Paragraphs[0].Append(unitBrutto.ToString("0.00"));
                table.Rows[i + 1].Cells[3].Paragraphs[0].Append(item.Quantity.ToString());
                table.Rows[i + 1].Cells[4].Paragraphs[0].Append("27%");
                table.Rows[i + 1].Cells[5].Paragraphs[0].Append(netto.ToString("0.00"));
                table.Rows[i + 1].Cells[6].Paragraphs[0].Append(brutto.ToString("0.00"));

                osszNetto += netto;
                osszAfa += afa;
                osszBrutto += brutto;
            }

            // ⬇️ Táblázat beszúrása
            var p = doc.Paragraphs.FirstOrDefault(x => x.Text.Contains("{{BeszurasHelye}}"));
            if (p != null)
            {
                p.ReplaceText("{{BeszurasHelye}}", "");
                p.InsertTableAfterSelf(table);
            }
            else
            {
                doc.InsertParagraph().InsertTableAfterSelf(table);
            }

            //Szállítási díj
            decimal szallitasDij = 1000;
            decimal totalBruttoSzallitassal = osszBrutto + szallitasDij;

            //Összegzések
            doc.ReplaceText("{{TotelGrandNetto}}", osszNetto.ToString("0.00"));
            doc.ReplaceText("{{TotelGrandAFA}}", osszAfa.ToString("0.00"));
            doc.ReplaceText("{{Szallitas}}", szallitasDij.ToString("0.00"));
            doc.ReplaceText("{{TotalGrand}}", totalBruttoSzallitassal.ToString("0.00"));
            doc.ReplaceText("{{TotelGrand}}", totalBruttoSzallitassal.ToString("0.00"));

            doc.SaveAs(kimenetiPath);
            MessageBox.Show("Számla elkészült: " + kimenetiPath);
        }



        //SZÁMLA ELKÉSZÍTÉSE
        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("Válassz ki egy rendelést a számlához!");
                return;
            }

            string orderBvin = dataGridView1.SelectedRows[0].Cells["OrderBvin"].Value.ToString();
            Api proxy = apiHivas();
            var response = proxy.OrdersFind(orderBvin);

            if (response == null || response.Content == null)
            {
                MessageBox.Show("Hiba a rendelés lekérésekor.");
                return;
            }

            GeneralSzamla(response.Content);
        }

        //CIMKE GENERÁLÁSA

        private void GeneralCimke(OrderDTO order)
        {
            string sablonPath = Path.Combine(Application.StartupPath, "PixelPress_SzallitasiCimke.docx");
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string kimenetiPath = Path.Combine(desktopPath, $"Cimke_{order.OrderNumber}.docx");

            int mennyiseg = order.Items.Sum(x => x.Quantity);
            decimal vegosszeg = order.TotalGrand;

            string random1 = new Random().Next(1000000, 9999999).ToString();
            string random2 = new Random().Next(100000, 999999).ToString();
            int sulySzorzat = new Random().Next(100, 999);
            int suly = mennyiseg * sulySzorzat;

            string iranyitoszam = order.BillingAddress?.PostalCode ?? "0000";
            string telefon = string.IsNullOrWhiteSpace(order.BillingAddress?.Phone)
                             ? random1
                             : order.BillingAddress.Phone;

            //Vonalkód generálás
            string kod = $"PP-{order.OrderNumber}-{new Random().Next(1000, 9999)}";
            Barcode b = new Barcode();
            System.Drawing.Image vonalkodKep = b.Encode(TYPE.CODE128, kod, Color.Black, Color.White, 300, 100);
            string kepPath = Path.Combine(Path.GetTempPath(), $"vonalkod_{order.OrderNumber}.png");
            vonalkodKep.Save(kepPath, System.Drawing.Imaging.ImageFormat.Png);

            var doc = DocX.Load(sablonPath);

            //Mezők cseréje
            doc.ReplaceText("{{BillingName}}", $"{order.BillingAddress.FirstName} {order.BillingAddress.LastName}");
            doc.ReplaceText("{{BillingStreet}}", order.BillingAddress.Line1);
            doc.ReplaceText("{{Billing City}}", order.BillingAddress.City);
            doc.ReplaceText("{{Iranyito}}", iranyitoszam);
            doc.ReplaceText("{{Telefon}}", telefon);
            doc.ReplaceText("{{Random1}}", random1);
            doc.ReplaceText("{{Random2}}", random2);
            doc.ReplaceText("{{Suly}}", suly.ToString());
            doc.ReplaceText("{{TotalGrand}}", vegosszeg.ToString("0.00"));
            doc.ReplaceText("{{Mennyiseg}}", mennyiseg.ToString());

            //Vonalkép beszúrása
            var kepHely = doc.Paragraphs.FirstOrDefault(x => x.Text.Contains("{{VonalkodHelye}}"));
            if (kepHely != null)
            {
                kepHely.ReplaceText("{{VonalkodHelye}}", "");
                var image = doc.AddImage(kepPath);
                var picture = image.CreatePicture(100, 300);
                kepHely.AppendPicture(picture).Alignment = Alignment.center;
            }

            //Mentés
            doc.SaveAs(kimenetiPath);
            MessageBox.Show("Címke generálva: " + kimenetiPath);

        }
        //CÍMKE ELKÉSZÍTÉSE

        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("Előbb válassz ki egy rendelést!");
                return;
            }

            string orderBvin = dataGridView1.SelectedRows[0].Cells["OrderBvin"].Value.ToString();
            Api proxy = apiHivas();
            var response = proxy.OrdersFind(orderBvin);

            if (response == null || response.Content == null)
            {
                MessageBox.Show("Hiba a rendelés lekérésekor.");
                return;
            }

            GeneralCimke(response.Content);
        }




    }
}
