namespace Holm.SPCalendarRecurrenceExpander

open System
open System.Globalization
open System.Collections.Generic
open System.Text.RegularExpressions

type KindOfDayQualifier =
    | First
    | Second
    | Third
    | Fourth
    | Last

type KindOfDay =
    | Day
    | Weekday
    | WeekendDay
    | DayOfWeek of DayOfWeek

type EndRange =
    // SharePoint defines "no explicit end range" as 999 repeat instances
    | NoExplicitEndRange
    | RepeatInstances of int
    | WindowEnd of DateTime

type DailyPattern =
    | EveryNthDay of int
    | EveryWeekDay

type WeeklyPattern =
    | EveryNthWeekOnDays of int * Set<DayOfWeek>

type MonthlyPattern =
    | EveryNthDayOfEveryMthMonth of int * int
    | EveryQualifierOfKindOfDayEveryMthMonth of KindOfDayQualifier * KindOfDay * int

type YearlyPattern =
    | EveryNthDayOfEveryMMonth of int * int
    | EveryQualifierOfKindOfDayMMonth of KindOfDayQualifier * KindOfDay * int

type MasterSeriesItemId = int

type Recurrence =
    | NoRecurrence
    | UnknownRecurrence
    | DeletedRecurrenceInstance of MasterSeriesItemId
    | RecurreceExceptionInstance of MasterSeriesItemId * DateTime
    | Daily of DailyPattern * EndRange
    | Weekly of WeeklyPattern * EndRange
    | Monthly of MonthlyPattern * EndRange
    | Yearly of YearlyPattern * EndRange

type Appointment =
    { Id: int
      Start: DateTime
      End: DateTime
      Duration: int64
      Recurrence: Recurrence }

type Parser() =
    let groupAsString (m: Match) (g: string) = m.Groups.[g].Value
    let groupAsInt (m: Match) (g: string) = (groupAsString m g) |> int

    let groupAsDateTime (m: Match) (g: string) = 
        let s = groupAsString m g
        DateTime.Parse(s, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind)

    let (|DailyEveryNthDay|_|) s =
        let re = Regex("<daily dayFrequency=\"(?<dayFrequency>(\d+))\" />")
        let m = re.Match(s)
        if m.Success then Some (groupAsInt m "dayFrequency")
        else None        

    let (|DailyEveryWeekDay|_|) s =
        let re = Regex("<daily weekday=\"TRUE\" />")
        if re.Match(s).Success then Some(0) else None

    let (|WeeklyEveryNthWeekOnDays|_|) s =
        let re = Regex("<repeat><weekly (?<su>(su=\"TRUE\")?).?(?<mo>(mo=\"TRUE\")?).?(?<tu>(tu=\"TRUE\")?).?(?<we>(we=\"TRUE\")?).?(?<th>(th=\"TRUE\")?).?(?<fr>(fr=\"TRUE\")?).?(?<sa>(sa=\"TRUE\")?).?weekFrequency=\"(?<weekFrequency>(\d+))\" /></repeat>")
        let m = re.Match(s)
        
        if m.Success then
            let g (k: string) = m.Groups.[k].Length > 0
            let days = List<DayOfWeek>()
            if g "su" then days.Add(DayOfWeek.Sunday)
            if g "mo" then days.Add(DayOfWeek.Monday)
            if g "tu" then days.Add(DayOfWeek.Tuesday)
            if g "we" then days.Add(DayOfWeek.Wednesday)
            if g "th" then days.Add(DayOfWeek.Thursday)
            if g "fr" then days.Add(DayOfWeek.Friday)
            if g "sa" then days.Add(DayOfWeek.Saturday) 
            Some (groupAsInt m "weekFrequency", days |> Set.ofSeq)
        else None

    let (|MonthlyEveryNthDayOfEveryMthMonth|_|) s =
        let re = Regex("<repeat><monthly monthFrequency=\"(?<monthFrequency>(\d+))\" day=\"(?<day>(\d+))\" /></repeat>")
        let m = re.Match(s)
        if m.Success then Some (groupAsInt m "day", groupAsInt m "monthFrequency")
        else None

    let (|KindOfDayQualifier|_|) s =
        match s with
        | "first" -> Some First
        | "second" -> Some Second
        | "third" -> Some Third
        | "fourth" -> Some Fourth
        | "last" -> Some Last
        | _ -> None

    let (|KindOfDay|_|) s =
        match s with
        | "day" -> Some Day
        | "weekday" -> Some Weekday
        | "weekend_day" -> Some WeekendDay
        | "su" -> Some(DayOfWeek(DayOfWeek.Sunday))
        | "mo" -> Some(DayOfWeek(DayOfWeek.Monday))
        | "tu" -> Some(DayOfWeek(DayOfWeek.Thursday))
        | "we" -> Some(DayOfWeek(DayOfWeek.Wednesday))
        | "th" -> Some(DayOfWeek(DayOfWeek.Thursday))
        | "fr" -> Some(DayOfWeek(DayOfWeek.Friday))
        | "sa" -> Some(DayOfWeek(DayOfWeek.Saturday))
        | _ -> None

    let (|MonthlyEveryQualifierOfKindOfDayEveryMthMonth|_|) s =
        let re = Regex("<repeat><monthlyByDay (?<kindOfDay>(day|weekday|weekend_day|su|mo|tu|we|th|fr|sa))=\"TRUE\" weekdayOfMonth=\"(?<kindOfDayQualifier>(first|second|third|fourth|last))\" monthFrequency=\"(?<monthFrequency>(\d+))\" /></repeat>")
        let m = re.Match(s)
        if m.Success then 
            let kindOfDayQualifier = 
                match groupAsString m "kindOfDayQualifier" with
                | KindOfDayQualifier q -> q
                | _ -> failwithf "Unable to parse kindOfDayQualifier: %s" (groupAsString m "kindOfDayQualifier")

            let kindOfDay = 
                match groupAsString m "kindOfDay" with
                | KindOfDay d -> d
                | _ -> failwithf "Unable to parse kindOfDay: %s" (groupAsString m "kindOfDay")

            Some (kindOfDayQualifier, kindOfDay, groupAsInt m "monthFrequency")
        else None

    let (|YearlyEveryNthDayOfEveryMMonth|_|) s =
        let re = Regex("<repeat><yearly yearFrequency=\"1\" month=\"(?<month>(\d+))\" day=\"(?<day>(\d+))\" /></repeat>")
        let m = re.Match(s)
        if m.Success then Some(groupAsInt m "day", groupAsInt m "month")
        else None

    let (|YearlyEveryQualifierOfKindOfDayNMonth|_|) s =
        let re = Regex("<repeat><yearlyByDay yearFrequency=\"1\" (?<kindOfDay>(day|weekday|weekend_day|su|mo|tu|we|th|fr|sa))=\"TRUE\" weekdayOfMonth=\"(?<kindOfDayQualifier>(first|second|third|fourth|last))\" month=\"(?<month>(\d+))\" /></repeat>")
        let m = re.Match(s)
        if m.Success then
            let kindOfDayQualifier = 
                match groupAsString m "kindOfDayQualifier" with
                | KindOfDayQualifier q -> q
                | _ -> failwithf "Unable to parse kindOfDayQualifier: %s" (groupAsString m "kindOfDayQualifier")

            let kindOfDay = 
                match groupAsString m "kindOfDay" with
                | KindOfDay d -> d
                | _ -> failwithf "Unable to parse kindOfDay: %s" (groupAsString m "kindOfDay")

            Some (kindOfDayQualifier, kindOfDay, groupAsInt m "month")
        else None

    let (|NoExplicitEndRange|_|) s =
        let re = Regex("<repeatForever>FALSE</repeatForever>")
        if re.Match(s).Success then Some(0) else None

    let (|RepeatInstances|_|) s =
        let re = Regex("<repeatInstances>(?<repeatInstances>(\d+))</repeatInstances>")
        let m = re.Match(s)
        if m.Success then Some(groupAsInt m "repeatInstances")
        else None

    let (|WindowEnd|_|) s =
        let re = Regex("<windowEnd>(?<windowEnd>(.*?))</windowEnd>")
        let m = re.Match(s)
        if m.Success then Some(groupAsDateTime m "windowEnd")
        else None       

    let parseRecurrence(d: Dictionary<string, obj>) =
        if (d.["fRecurrence"] :?> bool) then 
            if (d.["EventType"] :?> int) = 3 then
                DeletedRecurrenceInstance(d.["MasterSeriesItemID"] :?> int)
            else if (d.["EventType"] :?> int) = 4 then
                RecurreceExceptionInstance(d.["MasterSeriesItemID"] :?> int, d.["RecurrenceID"] |> string |> DateTime.Parse )
            else
                let rd = d.["RecurrenceData"] |> string
            
                let endRange =
                    match rd with
                    | NoExplicitEndRange _ -> NoExplicitEndRange
                    | RepeatInstances n -> RepeatInstances n
                    | WindowEnd dt -> 
                        // WindowEnd contains both a date and a time component expressed
                        // in the same timezone as the start and end dates of the event.
                        // WindowEnd is equal to end date and time of the last recurrence
                        // event except when end is at midnight (00:00am) in which case 
                        // WindowEnd is 24 hours ahead.
                        if dt.Hour = 0 && dt.Minute = 0 
                        then WindowEnd dt
                        else WindowEnd (d.["EndDate"] |> string |> DateTime.Parse)
                    | _ -> failwith "Unable to parse EndRange"

                match rd with
                | DailyEveryNthDay n -> Daily(EveryNthDay n, endRange)
                | DailyEveryWeekDay _ -> Daily(EveryWeekDay, endRange)
                | WeeklyEveryNthWeekOnDays (n, days) -> Weekly(EveryNthWeekOnDays(n, days), endRange)
                | MonthlyEveryNthDayOfEveryMthMonth (d, m) -> Monthly(MonthlyPattern.EveryNthDayOfEveryMthMonth(d, m), endRange)
                | MonthlyEveryQualifierOfKindOfDayEveryMthMonth (q, k, n) -> Monthly(EveryQualifierOfKindOfDayEveryMthMonth(q, k, n), endRange)
                | YearlyEveryNthDayOfEveryMMonth (n, m) -> Yearly(EveryNthDayOfEveryMMonth(n, m), endRange)
                | YearlyEveryQualifierOfKindOfDayNMonth (q, k, n) -> Yearly(EveryQualifierOfKindOfDayMMonth(q, k, n), endRange)
                | _ -> UnknownRecurrence
        else NoRecurrence

    let parse (a: Dictionary<string, obj>) =
        { Id = a.["ID"] |> string |> Int32.Parse
          Start = a.["EventDate"] |> string |> DateTime.Parse
          End = a.["EndDate"] |> string |> DateTime.Parse
          Duration = a.["Duration"] |> string |> Int64.Parse
          Recurrence = parseRecurrence a }

    member __.Parse(appointment: Dictionary<string, obj>) =
        parse appointment