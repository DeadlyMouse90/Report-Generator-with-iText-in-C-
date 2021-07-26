using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Diagnostics;



namespace GerandoRelatorioPDF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void BtnGerarRelatorio_Click(object sender, EventArgs e)
        {
            RelatorioPDF();
        }

        private void RelatorioPDF()
        {
            Document document = new Document(PageSize.A4);
            document.SetMargins(30, 20, 20, 20);
            string filePath = Directory.GetCurrentDirectory() + "\\Relatorio.pdf";
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Create));

            iTextSharp.text.Font.FontFamily family = new iTextSharp.text.Font.FontFamily();
            family = iTextSharp.text.Font.FontFamily.HELVETICA;
            iTextSharp.text.Font font = new iTextSharp.text.Font(family, 8, (int)System.Drawing.FontStyle.Bold);

            Eventos ev = new Eventos(font);
            writer.PageEvent = ev;

            document.Open();
            PdfPTable table = new PdfPTable(5);

            string titulo = "";
            iTextSharp.text.Font fontetitulo = FontFactory.GetFont(BaseFont.HELVETICA, 14);
            Paragraph title = new Paragraph(titulo, fontetitulo);
            title.Alignment = Element.ALIGN_CENTER;
            title.Add("Relatório de Matriculas processadas");
            document.Add(title);

            iTextSharp.text.Font fonte = FontFactory.GetFont(BaseFont.HELVETICA, 10, iTextSharp.text.Font.BOLD);
            Paragraph coluna1 = new Paragraph("#", fonte);
            Paragraph coluna2 = new Paragraph("Data/Hora", fonte);
            Paragraph coluna3 = new Paragraph("Cliente", fonte);
            Paragraph coluna4 = new Paragraph("Matricula", fonte);
            Paragraph coluna5 = new Paragraph("Usuario", fonte);

            coluna1.Alignment = Element.ALIGN_LEFT;
            coluna2.Alignment = Element.ALIGN_LEFT;
            coluna3.Alignment = Element.ALIGN_LEFT;
            coluna4.Alignment = Element.ALIGN_LEFT;
            coluna5.Alignment = Element.ALIGN_LEFT;

            var cell1 = new PdfPCell();
            var cell2 = new PdfPCell();
            var cell3 = new PdfPCell();
            var cell4 = new PdfPCell();
            var cell5 = new PdfPCell();

            cell1.AddElement(coluna1);
            cell2.AddElement(coluna2);
            cell3.AddElement(coluna3);
            cell4.AddElement(coluna4);
            cell5.AddElement(coluna5);

            cell1.Border = 2;
            cell2.Border = 2;
            cell3.Border = 2;
            cell4.Border = 2;
            cell5.Border = 2;

            table.AddCell(cell1);
            table.AddCell(cell2);
            table.AddCell(cell3);
            table.AddCell(cell4);
            table.AddCell(cell5);

            List<DadosEquipamentos> dados = new List<DadosEquipamentos>();

            for(int x=1; x <= 150; x++)
            {
                DadosEquipamentos dado = new DadosEquipamentos();
                dado.Num = x.ToString();
                dado.Data = "20-06-2021";
                dado.Cliente = "Rogerio";
                dado.Matricula = "ADBC";
                dado.Usuario = "Ernane";
                dados.Add(dado);
            }

            iTextSharp.text.Font fonte2 = FontFactory.GetFont(BaseFont.HELVETICA, 8);

            foreach(var dado in dados)
            {
                Phrase num = new Phrase(dado.Num, fonte2);
                var cell = new PdfPCell(num);
                cell.HorizontalAlignment = 0;
                cell.Border = 0;
                table.AddCell(cell);

                Phrase data = new Phrase(dado.Data, fonte2);
                cell = new PdfPCell(data);
                cell.HorizontalAlignment = 0;
                cell.Border = 0;
                table.AddCell(cell);

                Phrase cliente = new Phrase(dado.Cliente, fonte2);
                cell = new PdfPCell(cliente);
                cell.HorizontalAlignment = 0;
                cell.Border = 0;
                table.AddCell(cell);

                Phrase matricula = new Phrase(dado.Matricula, fonte2);
                cell = new PdfPCell(matricula);
                cell.HorizontalAlignment = 0;
                cell.Border = 0;
                table.AddCell(cell);

                Phrase usuario = new Phrase(dado.Usuario, fonte2);
                cell = new PdfPCell(usuario);
                cell.HorizontalAlignment = 0;
                cell.Border = 0;
                table.AddCell(cell);
            }


            byte[] imageStream = getChartImage();
            iTextSharp.text.Image picture = iTextSharp.text.Image.GetInstance(imageStream);
            picture.ScalePercent(65f);
            picture.Alignment = 1;
            document.Add(picture);
            document.Add(table);

            document.Close();
            Process.Start(filePath);
        }

        private byte[] getChartImage()
        {
            using(var chartimage=new MemoryStream())
            {
                chart1.SaveImage(chartimage, System.Drawing.Imaging.ImageFormat.Png);
                return chartimage.GetBuffer();
            }
        }

        public class DadosEquipamentos
        {
            public string Num;
            public string Data;
            public string Cliente;
            public string Matricula;
            public string Usuario;
        }

    }

    class Eventos : PdfPageEventHelper
    {
        public iTextSharp.text.Font fonte { get; set; }

        public Eventos(iTextSharp.text.Font fonte_)
        {
            fonte = fonte_;
        }

        public override void OnStartPage(PdfWriter writer, Document document)
        {
            string imagem = Directory.GetCurrentDirectory() + "\\logobeaver.png";
            iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(imagem);
            image.ScalePercent(15f);
            image.Alignment = 4;
            image.Alignment = 0;
            document.Add(image);
            Paragraph ph = new Paragraph();

            PdfContentByte cb = writer.DirectContent;

            BaseFont font;

            font = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);

            cb.SetFontAndSize(font, 8);

            string operador = "ernane";
            string data = "29/06/2021";
            string coleta = "Relatório mês junho";

            cb.ShowTextAligned(Element.ALIGN_RIGHT, coleta, document.Right, document.Top, 0);
            cb.ShowTextAligned(Element.ALIGN_RIGHT, data, document.Right, document.Top, 0);
            cb.ShowTextAligned(Element.ALIGN_RIGHT, operador, document.Right, document.Top, 0);

            iTextSharp.text.pdf.draw.VerticalPositionMark separator = new iTextSharp.text.pdf.draw.LineSeparator();

            ph.Add(separator);

            ph.Add(new Chunk("\n"));

            document.Add(ph);

        }

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            PdfContentByte cb = writer.DirectContent;

            BaseFont font;

            font = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);

            cb.SetFontAndSize(font, 8);

            string texto = "Página: " + writer.PageNumber.ToString();

            cb.ShowTextAligned(Element.ALIGN_RIGHT, texto, document.Right, document.Bottom - 5, 0);
        }


    }
}
