using ClosedXML.Excel;

using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;

using System.Text;

namespace CreateCalendar
{
    internal class Program
    {
        static void Main(string[] args)
        {
            
            using (FileStream fileStream = new FileStream(@"data\teknologisk.xlsx",
                FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var wb = new XLWorkbook(fileStream);
                var ws = wb.Worksheet("Underviserplan - frem i tid");

                int row = 2;
                var calendar = new Calendar(); 
                calendar.Name = "Undervisning";
                calendar.Properties.Add(new CalendarProperty("X-WR-CALNAME", "Undervisning"));
                while (!string.IsNullOrEmpty(ws.Cell(row, 1).Value.ToString()))
                {
                    var e = new CalendarEvent();
                    var start = ws.Cell(row, 5).GetValue<DateTime>();
                    var end = ws.Cell(row, 6).GetValue<DateTime>();
                    e.Start = new CalDateTime(start);
                    e.End = new CalDateTime(end);
                    e.Name = ws.Cell(row, 8).Value.ToString();
                    e.Location = ws.Cell(row, 14).Value.ToString();
                    e.Description = ws.Cell(row, 10).Value.ToString();
                    calendar.Events.Add(e);
                    Console.WriteLine(e.Description);
                    row++;
                }
                
                var serializer = new CalendarSerializer();
                var serializedCalendar = serializer.SerializeToString(calendar);

                File.WriteAllText("data/calendar.ics", serializedCalendar, new UTF8Encoding(false));
            }
        }
    }
}
