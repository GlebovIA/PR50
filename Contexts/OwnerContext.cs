using Microsoft.Office.Interop.Word;
using PR50.Models;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;
namespace PR50.Contexts
{
    public class OwnerContext : Owner
    {
        public OwnerContext(string surname, string name, string lastname, int numberRoom) : base(surname, name, lastname, numberRoom) { }
        public static List<OwnerContext> AllOwners()
        {
            List<OwnerContext> allOwners = new List<OwnerContext>();
            allOwners.Add(new OwnerContext("Иванова", "Елена", "Петровна", 1));
            allOwners.Add(new OwnerContext("Смирнов", "Алексей", "Владимирович", 2));
            allOwners.Add(new OwnerContext("Кузнецова", "Анна", "Сергеевна", 3));
            allOwners.Add(new OwnerContext("Павлов", "Дмитрий", "Александрович", 3));
            allOwners.Add(new OwnerContext("Михайлова", "Ольга", "Ивановна", 4));
            allOwners.Add(new OwnerContext("Козлов", "Артем", "Олегович", 5));
            allOwners.Add(new OwnerContext("Соколова", "Наталья", "Викторовна", 6));
            allOwners.Add(new OwnerContext("Лебедев", "Игорь", "Андреевич", 6));
            allOwners.Add(new OwnerContext("Федорова", "Екатерина", "Дмитриевна", 7));
            allOwners.Add(new OwnerContext("Александров", "Андрей", "Игоревич", 7));
            allOwners.Add(new OwnerContext("Степанова", "Оксана", "Николаевна", 8));
            allOwners.Add(new OwnerContext("Никитин", "Сергей", "Васильевич", 9));
            allOwners.Add(new OwnerContext("Ковалева", "Мария", "Александровна", 10));
            allOwners.Add(new OwnerContext("Фролов", "Павел", "Михайлович", 11));
            allOwners.Add(new OwnerContext("Белова", "Елена", "Александровна", 12));
            allOwners.Add(new OwnerContext("Поляков", "Илья", "Данилович", 13));
            allOwners.Add(new OwnerContext("Гаврилова", "Анастасия", "Валерьевна", 14));
            allOwners.Add(new OwnerContext("Орлов", "Денис", "Владимирович", 15));
            allOwners.Add(new OwnerContext("Киселева", "Алина", "Сергеевна", 16));
            allOwners.Add(new OwnerContext("Ткаченко", "Артем", "Викторович", 16));
            allOwners.Add(new OwnerContext("Романова", "Валерия", "Павловна", 16));
            allOwners.Add(new OwnerContext("Максимов", "Александр", "Юрьевич", 17));
            allOwners.Add(new OwnerContext("Сидорова", "Евгения", "Игоревна", 17));
            allOwners.Add(new OwnerContext("Антонов", "Никита", "Алексеевич", 18));
            allOwners.Add(new OwnerContext("Дмитриева", "Юлия", "Владимировна", 19));
            return allOwners;
        }
        public static void Report(string fileName)
        {
            Word.Application app = new Word.Application();
            Word.Document document = app.Documents.Add();
            Word.Paragraph header = document.Paragraphs.Add();
            header.Range.Font.Size = 16;
            header.Range.Text = "Список жильцов дома";
            header.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            header.Range.ParagraphFormat.SpaceAfter = 0;
            header.Range.Font.Bold = 1;
            header.Range.InsertParagraphAfter();

            Word.Paragraph address = document.Paragraphs.Add();
            address.Range.Font.Size = 14;
            address.Range.Text = "по адресу: г. Пермь, ул. Луначарского, д. 24";
            address.Range.ParagraphFormat.SpaceAfter = 20;
            address.Range.Font.Bold = 0;
            address.Range.InsertParagraphAfter();

            Word.Paragraph count = document.Paragraphs.Add();
            count.Range.Font.Size = 14;
            count.Range.Text = $"Всего жильцов: {AllOwners().Count}";
            count.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            count.Range.ParagraphFormat.SpaceAfter = 0;
            count.Range.InsertParagraphAfter();

            Word.Paragraph table = document.Paragraphs.Add();
            Word.Table payments = document.Tables.Add(table.Range, AllOwners().OrderBy(x => x.NumberRoom).ToArray().Last().NumberRoom + 1, 4);
            payments.Borders.InsideLineStyle = payments.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            payments.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            Cell("№", payments.Cell(1, 1).Range);
            Cell("Фамилия", payments.Cell(1, 2).Range);
            Cell("Имя", payments.Cell(1, 3).Range);
            Cell("Отчество", payments.Cell(1, 4).Range);

            string surnames = "";
            string names = "";
            string lastnames = "";
            int cell = 2;
            List<OwnerContext> All = AllOwners().OrderBy(x => x.NumberRoom).ToList();
            int currentRoom = All[0].NumberRoom;
            for (int i = 0; i < All.Count; i++)
            {
                if (All[i].NumberRoom != currentRoom && i != 0)
                {
                    Cell((currentRoom).ToString(), payments.Cell(cell, 1).Range);
                    Cell(surnames, payments.Cell(cell, 2).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                    Cell(names, payments.Cell(cell, 3).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                    Cell(lastnames, payments.Cell(cell, 4).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                    currentRoom = All[i].NumberRoom;
                    surnames = "";
                    names = "";
                    lastnames = "";
                    cell++;
                }
                if (surnames != "")
                {
                    surnames += "\n";
                    names += "\n";
                    lastnames += "\n";
                }
                surnames += All[i].Surname;
                names += All[i].Name;
                lastnames += All[i].Lastname;
                if (i == All.Count - 1)
                {
                    Cell((currentRoom).ToString(), payments.Cell(cell, 1).Range);
                    Cell(surnames, payments.Cell(cell, 2).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                    Cell(names, payments.Cell(cell, 3).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                    Cell(lastnames, payments.Cell(cell, 4).Range, Word.WdParagraphAlignment.wdAlignParagraphLeft);
                }

            }
            document.SaveAs2(fileName);
            document.Close();
            app.Quit();
        }
        /// <summary>
        /// Добавление текста в ячейку
        /// </summary>
        /// <param name="text">Текст в ячейке</param>
        /// <param name="Cell">Ячейка</param>
        /// <param name="Alignment">Положение в ячейке</param>
        public static void Cell(string text, Word.Range Cell, WdParagraphAlignment Alignment = WdParagraphAlignment.wdAlignParagraphCenter)
        {
            Cell.Text = text;
            Cell.ParagraphFormat.Alignment = Alignment;
        }
    }
}
