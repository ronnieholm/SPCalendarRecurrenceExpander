open System
open System.Security
open Microsoft.SharePoint.Client
open Holm.SPCalendarRecurrenceExpander

type ScheduleAppointment =
    { Id: int
      Title: string
      Start: DateTime
      End: DateTime
      (* add any custom columns here *) }

module Program =
    [<EntryPoint>]
    let main _ =
        let web = "https://<tenant>.sharepoint.com/<subsite>"
        let username = "<username@domain>"
        let password = "<password>"
        let listTitle = "<list title>"

        let ctx = new ClientContext(web)
        let securePassword = new SecureString()
        password.ToCharArray() |> Seq.iter (securePassword.AppendChar)
        ctx.Credentials <- SharePointOnlineCredentials(username, securePassword)
        let calendar = ctx.Web.Lists.GetByTitle(listTitle)
        ctx.Load(ctx.Web.RegionalSettings.TimeZone)
        let tz = ctx.Web.RegionalSettings.TimeZone
        ctx.ExecuteQuery()

        let query = CamlQuery()
        let items = calendar.GetItems(query)
        ctx.Load(items)
        ctx.ExecuteQuery()

        let collapsedAppointments = ResizeArray<_>(items |> Seq.map (fun i -> i.FieldValues))
        let expander = CalendarRecurrenceExpander(tz.Information.Bias, tz.Information.DaylightBias)
        let recurrenceInstances = expander.Expand(collapsedAppointments)

        let toDomainObject (ri: RecurrenceInstance) = 
            let a = 
                collapsedAppointments 
                |> Seq.find (fun i -> i.["ID"] |> string |> Int32.Parse = ri.Id)
            { Id = ri.Id
              Title = a.["Title"] |> string
              Start = ri.Start
              End = ri.End }
        
        let expandedAppointments = recurrenceInstances |> Seq.map toDomainObject       
        expandedAppointments 
        |> Seq.iter (fun a -> 
            printfn "%d %s %s %s" a.Id a.Title (a.Start.ToShortDateString() + " " + a.Start.ToShortTimeString()) (a.End.ToShortDateString() + " " + a.End.ToShortTimeString()))
        0