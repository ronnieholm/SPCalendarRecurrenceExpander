using System;
using System.Linq;
using System.Security;
using Microsoft.SharePoint.Client;
using Holm.SPCalendarRecurrenceExpander;

namespace SPCalendarRecurrenceExpander.CSharpExample {
    class Appointment {
        public int Id { get; set; }
        public string Title { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        // add any custom columns here
    }

    class Program {
        static void Main(string[] args) {
            var web = "https://<tenant>.sharepoint.com/<subsite>";
            var username = "<username@domain>";
            var password = "<password>";
            var calendarTitle = "<list title>";

            var ctx = new ClientContext(web);
            var securePassword = new SecureString();
            password.ToList().ForEach(securePassword.AppendChar);
            ctx.Credentials = new SharePointOnlineCredentials(username, securePassword);
            var calendar = ctx.Web.Lists.GetByTitle(calendarTitle);
            ctx.Load(ctx.Web.RegionalSettings.TimeZone);
            var tz = ctx.Web.RegionalSettings.TimeZone;
            ctx.ExecuteQuery();

            var query = new CamlQuery();
            var items = calendar.GetItems(query);
            ctx.Load(items);
            ctx.ExecuteQuery();

            var collapsedAppointments = items.ToList().Select(i => i.FieldValues).ToList();
            var expander = new CalendarRecurrenceExpander(tz.Information.Bias, tz.Information.DaylightBias);
            var recurrenceInstances = expander.Expand(collapsedAppointments);

            Func<RecurrenceInstance, Appointment> toDomainObject = (ri => {
                var a = collapsedAppointments.First(i => int.Parse(i["ID"].ToString()) == ri.Id);
                return new Appointment {
                    Id = ri.Id,
                    Title = (string) a["Title"],
                    Start = ri.Start,
                    End = ri.End
                };
            });

            var expandedAppointments = recurrenceInstances.Select(toDomainObject).ToList();
            expandedAppointments
                .ForEach(a => 
                    Console.WriteLine(
                        string.Format("{0} {1} {2} {3} {4} {5}", 
                            a.Id, a.Title, a.Start.ToShortDateString(), a.Start.ToShortTimeString(),
                            a.End.ToShortDateString(), a.End.ToShortTimeString())));
        }
    }
}
