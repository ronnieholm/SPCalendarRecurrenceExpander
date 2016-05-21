namespace Holm.SPCalendarRecurrenceExpander

open System
open System.Collections.Generic

type CalendarRecurrenceExpander(tzBias: int, tzDaylightBias: int) =
    let parseDateTime dt = dt |> string |> DateTime.Parse        

    let toLocalTime dt =
        (parseDateTime dt).AddMinutes(float(-tzBias - tzDaylightBias))

    let timeZoneCorrect (a: Dictionary<string, obj>) =
        // all day events already have local EventDate and EndDate which is
        // 12:00 AM and 11.59 PM, respectively
        if (a.["fAllDayEvent"] |> string |> bool.Parse) 
        then a
        else
            a.["EventDate"] <- toLocalTime (a.["EventDate"])
            a.["EndDate"] <- toLocalTime (a.["EndDate"])
            if a.["RecurrenceID"] |> string <> "" 
            then a.["RecurrenceID"] <- toLocalTime (a.["RecurrenceID"])
            a

    /// Dumps the internal structure of appointments for debugging purposes.
    member __.Dump(appointments: ResizeArray<Dictionary<string, obj>>) =
        let keys = 
            ["ID"; "EventDate"; "EndDate"; "Duration"; "fRecurrence"; "EventType"; 
                "MasterSeriesItemID"; "RecurrenceID"; "RecurrenceData"; "fAllDayEvent"]
        appointments 
        |> Seq.map (fun a -> 
            keys |> List.map (fun k -> 
                (k, if a.ContainsKey(k) then Some (a.[k]) else None)))
        |> Seq.toList
        |> sprintf "%A"                    

    /// Expands appointments into recurrence instances.
    member __.Expand(appointments: ResizeArray<Dictionary<string, obj>>) =       
        let tzCorrectedAppointments =
            appointments
            |> Seq.toList
            |> List.map timeZoneCorrect
            |> List.map (fun a -> Parser().Parse(a))        

        let regularAppointments =
            tzCorrectedAppointments
            |> List.filter (fun a ->
                match a.Recurrence with
                | DeletedRecurrenceInstance _
                | ModifiedRecurrenceInstance(_, _) -> false
                | _ -> true)

        let deletedInstancesByMasterSeriesItemId =
            tzCorrectedAppointments
            |> List.filter (fun a ->
                match a.Recurrence with
                | DeletedRecurrenceInstance _ -> true                
                | _ -> false)
            |> Seq.groupBy (fun a -> 
                a.Recurrence 
                |> function 
                    | DeletedRecurrenceInstance(masterSeriesItemId, _) -> masterSeriesItemId
                    | _ -> failwith "Should never happen")

        let modifiedInstanceByMasterSeriesItemId =           
            tzCorrectedAppointments
            |> List.filter (fun a ->
                match a.Recurrence with
                | ModifiedRecurrenceInstance _ -> true                
                | _ -> false)
            |> Seq.groupBy (fun a ->
                a.Recurrence
                |> function
                    | ModifiedRecurrenceInstance(masterSeriesItemId, _) -> masterSeriesItemId
                    | _ -> failwith "Should never happen")
            |> Seq.toList

        let expandedAppointments =
            regularAppointments
            |> List.map (fun a -> 
                let deletedInstancesForAppointment = 
                    match deletedInstancesByMasterSeriesItemId |> Seq.tryFind (fun (id, _) -> id = a.Id) with
                    | Some (_, instances) -> instances |> Seq.toList
                    | None -> []

                let modifiedInstancesForAppointment =
                    match modifiedInstanceByMasterSeriesItemId |> Seq.tryFind (fun (id, _) -> id = a.Id) with
                    | Some (_, instances) -> instances |> Seq.toList
                    | None -> []

                Compiler().Compile(a, deletedInstancesForAppointment, modifiedInstancesForAppointment))
            |> Seq.concat

        ResizeArray<_>(expandedAppointments)
