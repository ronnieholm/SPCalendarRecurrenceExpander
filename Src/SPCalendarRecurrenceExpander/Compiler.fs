namespace Holm.SPCalendarRecurrenceExpander

open System

type RecurrenceInstance =
    { Id: int
      Start: DateTime
      End: DateTime }

[<AutoOpen>]
module DateGenerators =
    type Dow = DayOfWeek
    let weekDays = [Dow.Monday; Dow.Tuesday; Dow.Wednesday; Dow.Thursday; Dow.Friday] // todo: use Set instead?
    let weekendDays = [Dow.Saturday; Dow.Sunday] 
    let allDays = weekDays @ weekendDays

    let rec genDateRange (first: DateTime) (last: DateTime) = seq {
        if first.Day > last.Day then
            yield first
            yield! genDateRange (first.AddDays(-1.)) last    
        elif first.Day < last.Day then
            yield first
            yield! genDateRange (first.AddDays(1.)) last
        elif first.Day = last.Day then
            yield first }

    let computeDaysHistogram (first: DateTime) (last: DateTime) =
        if first.Month <> last.Month then failwith "First month not equal to last month"
        genDateRange first last
        |> Seq.groupBy (fun dt -> dt.DayOfWeek)
        |> Seq.map (fun (dow, dates) -> (dow, dates |> Seq.length))
        |> Map.ofSeq

    let sumHistogramByDays (h: Map<DayOfWeek, int>) (d: DayOfWeek list) =
        let frequencyBy key =
            match h.TryFind key with
            | Some v -> v
            | None -> 0
        d |> List.fold (fun acc key -> acc + frequencyBy key) 0

    let genMonthHistograms (days: DayOfWeek list) year month =
        let daysInMonth = DateTime.DaysInMonth(year, month)
        [1..daysInMonth]
        |> List.map (fun d ->
            d, computeDaysHistogram (DateTime(year, month, 1)) (DateTime(year, month, d)))
        |> List.map (fun (d, h) -> d, sumHistogramByDays h days)

    let daysCountFinder n hist = 
        hist |> List.find (fun (_, h) -> h = n)

    let lastDaysFinder hist =
        let day = hist |> List.maxBy snd |> snd
        hist |> List.find (fun (_, h) -> h = day)
    
    let lastDaysUpperLimitFinder n hist =
        let day = hist |> List.maxBy snd |> snd
        let day' = min n day
        hist |> List.find (fun (_, h) -> h = day')

    let daysPatternOfEveryMthMonth finderFn days m first (dt: DateTime) =
        if first then
            // recurrence instance in partial month?
            let next = dt.AddDays(1.)
            let histograms = genMonthHistograms days (next.Year) (next.Month)
            let recurrence = histograms |> finderFn |> fun (day, _) -> DateTime(next.Year, next.Month, day, next.Hour, next.Minute, next.Second)
            
            if recurrence >= next then recurrence
            else
                let skippedAhead = 
                    DateTime(next.Year, next.Month, 1, next.Hour, next.Minute, next.Second).AddMonths(m)
                let histograms' = genMonthHistograms days (skippedAhead.Year) (skippedAhead.Month)
                histograms' |> finderFn |> fun (day, _) -> DateTime(skippedAhead.Year, skippedAhead.Month, day, skippedAhead.Hour, skippedAhead.Minute, skippedAhead.Second)
        else
            // given we've exchausted previous (partial) month and there can be one
            // and only one recurrence instance per month, we can safely skip ahead.
            let daysInMonth = DateTime.DaysInMonth(dt.Year, dt.Month)
            let skippedAhead = 
                DateTime(dt.Year, dt.Month, daysInMonth, dt.Hour, dt.Minute, dt.Second)
                    .AddDays(1.)
                    .AddMonths(m - 1)

            let histograms = genMonthHistograms days (skippedAhead.Year) (skippedAhead.Month)
            histograms |> finderFn |> fun (day, _) -> DateTime(skippedAhead.Year, skippedAhead.Month, day, skippedAhead.Hour, skippedAhead.Minute, skippedAhead.Second)

    let nextNthDaysPatternOfEveryMthMonth days n m first (dt: DateTime) =
        daysPatternOfEveryMthMonth (daysCountFinder n) days m first dt

    let lastDaysPatternOfEveryMthMonth (days: DayOfWeek list) m first (dt: DateTime) =        
        daysPatternOfEveryMthMonth lastDaysFinder days m first dt

    let nextNthDay n first (dt: DateTime) =
        if first then dt.AddDays(1.)
        else dt.AddDays(float n)

    let nextWeekDay _ (dt: DateTime) =
        let next = dt.AddDays(1.) 
        genDateRange next (DateTime.MaxValue) 
        |> Seq.find(fun d -> weekDays |> List.exists((=) d.DayOfWeek))

    let rec nextNthWeekOnDaysNext n daysOfWeek first (dt: DateTime) =
        let next = dt.AddDays(1.)
        if daysOfWeek |> Set.contains next.DayOfWeek 
        then next
        else
            if next.AddDays(1.).DayOfWeek = DayOfWeek.Monday
            then nextNthWeekOnDaysNext n daysOfWeek first (next.AddDays(float(7 * (n - 1))))
            else nextNthWeekOnDaysNext n daysOfWeek first next

    let daysPatternOfEveryMMonth finderFn days m first (dt: DateTime) =
        if first then
            let next = dt.AddDays(1.)
            let histograms = genMonthHistograms days (next.Year) (m)
            let recurrence = histograms |> finderFn |> fun (day, _) -> DateTime(next.Year, m, day, next.Hour, next.Minute, next.Second)

            // recurrence instance in partial year?
            if recurrence >= next then recurrence
            else               
                let next2 = DateTime(next.Year + 1, m, 1, next.Hour, next.Minute, next.Second)
                genMonthHistograms days next.Year next.Month
                |> finderFn
                |> fun (day, _) -> DateTime(next2.Year, next2.Month, day, next2.Hour, next2.Minute, next2.Second)
        else
            let finder histograms = histograms |> finderFn
            daysPatternOfEveryMthMonth finder days 12 first dt

    let nextNthDaysPatternOfEveryMMonth days n m first (dt: DateTime) =
        daysPatternOfEveryMMonth (daysCountFinder n) days m first dt

    let lastDaysPatternOfEveryMMonth days m first (dt: DateTime) =
        daysPatternOfEveryMMonth lastDaysFinder days m first dt

    // cannot be replaced with call to nextNthDaysPatternOfEveryMthMonth because
    // special care is needed to "round down" result to last day of month if
    // n greater than days in month.
    let nextNthDayOfEveryMthMonth n m first (dt: DateTime) =
        daysPatternOfEveryMthMonth (lastDaysUpperLimitFinder n) allDays m first dt

    let nextNthDayOfEveryMMonth n m first (dt: DateTime) =
        daysPatternOfEveryMMonth (lastDaysUpperLimitFinder n) allDays m first dt

type Compiler() =
    // we potentially need to pass in multiple appointments because each appointment 
    // isn't necessarily self-contained. A recurrence appointment might have deleted 
    // appointments which are appointments in themselves but with a special event type.
    member __.Compile(a: Appointment, deletedRecurrences: Appointment list, recurreceExceptions: Appointment list): seq<RecurrenceInstance> =
        // SharePoint allows events with start = end => 0 duration
        if a.Start >= a.End then failwith "a.start and a.end"

        let gen nextFn =
            let template = { Id = a.Id; Start = DateTime.MinValue; End = DateTime.MinValue }

            // given that start date may be a valid recurrence instance and generators
            // start by incrementing time pointer by one day, we set initial start
            // date back by one.
            (a.Start.AddDays(-1.), nextFn)
            |> Seq.unfold (fun (dt, nextFn) ->
                let next = if a.Start.AddDays(-1.) = dt then nextFn true dt else nextFn false dt
                if next > a.End then None
                // we cannot increment next here because some generators
                // require knowledge about previous recurrence instance
                // date to be able to properly skip ahead.
                else Some (next, (next, nextFn)))
            |> Seq.map (fun dt -> { template with Start = dt; End = dt.AddSeconds(float a.Duration) })

        let recurrences =
            match a.Recurrence with
            | DeletedRecurrenceInstance _ -> Seq.empty
            | RecurreceExceptionInstance _ -> Seq.empty
            | NoRecurrence -> seq { yield { Id = a.Id; Start = a.Start; End = a.Start.AddSeconds(float a.Duration) } }
            | Daily(p, _) ->
                match p with
                | EveryNthDay n -> gen (nextNthDay n)
                | EveryWeekDay -> gen nextWeekDay
            | Weekly(p, _) ->
                match p with
                | EveryNthWeekOnDays (n, daysOfWeek) -> gen (nextNthWeekOnDaysNext n daysOfWeek)
            | Monthly(p, _) ->
                match p with
                | EveryNthDayOfEveryMthMonth (n, m) -> gen (nextNthDayOfEveryMthMonth n m)
                | EveryQualifierOfKindOfDayEveryMthMonth (qualifier, kindOfDay, skipMonths) ->
                    match qualifier with                   
                    | First ->
                        match kindOfDay with
                        | Day -> gen (nextNthDaysPatternOfEveryMthMonth allDays 1 skipMonths)
                        | Weekday -> gen (nextNthDaysPatternOfEveryMthMonth weekDays 1 skipMonths)
                        | WeekendDay -> gen (nextNthDaysPatternOfEveryMthMonth weekendDays 1 skipMonths)
                        | DayOfWeek d -> gen (nextNthDaysPatternOfEveryMthMonth [d] 1 skipMonths)
                    | Second ->
                        match kindOfDay with
                        | Day -> gen (nextNthDaysPatternOfEveryMthMonth allDays 2 skipMonths)
                        | Weekday -> gen (nextNthDaysPatternOfEveryMthMonth weekDays 2 skipMonths)
                        | WeekendDay -> gen (nextNthDaysPatternOfEveryMthMonth weekendDays 2 skipMonths)
                        | DayOfWeek d -> gen (nextNthDaysPatternOfEveryMthMonth [d] 2 skipMonths)
                    | Third ->
                        match kindOfDay with
                        | Day -> gen (nextNthDaysPatternOfEveryMthMonth allDays 3 skipMonths)
                        | Weekday -> gen (nextNthDaysPatternOfEveryMthMonth weekDays 3 skipMonths)
                        | WeekendDay -> gen (nextNthDaysPatternOfEveryMthMonth weekendDays 3 skipMonths)
                        | DayOfWeek d -> gen (nextNthDaysPatternOfEveryMthMonth [d] 3 skipMonths)
                    | Fourth ->
                        match kindOfDay with
                        | Day -> gen (nextNthDaysPatternOfEveryMthMonth allDays 4 skipMonths)
                        | Weekday -> gen (nextNthDaysPatternOfEveryMthMonth weekDays 4 skipMonths)
                        | WeekendDay -> gen (nextNthDaysPatternOfEveryMthMonth weekendDays 4 skipMonths)
                        | DayOfWeek d -> gen (nextNthDaysPatternOfEveryMthMonth [d] 4 skipMonths)
                    | Last ->
                        match kindOfDay with
                        | Day -> gen (lastDaysPatternOfEveryMthMonth allDays skipMonths)
                        | Weekday -> gen (lastDaysPatternOfEveryMthMonth weekDays skipMonths)
                        | WeekendDay -> gen (lastDaysPatternOfEveryMthMonth weekendDays skipMonths)
                        | DayOfWeek d -> gen (lastDaysPatternOfEveryMthMonth [d] skipMonths)
            | Yearly (p, _) ->
                match p with
                | EveryNthDayOfEveryMMonth (n, m) -> gen (nextNthDayOfEveryMMonth n m)
                | EveryQualifierOfKindOfDayMMonth (qualifier, kindOfDay, month) ->
                    match qualifier with
                    | First -> 
                        match kindOfDay with
                        | Day -> gen (nextNthDaysPatternOfEveryMMonth allDays 1 month)
                        | Weekday -> gen (nextNthDaysPatternOfEveryMMonth weekDays 1 month)
                        | WeekendDay -> gen (nextNthDaysPatternOfEveryMMonth weekendDays 1 month)
                        | DayOfWeek d -> gen (nextNthDaysPatternOfEveryMMonth [d] 1 month)
                    | Second ->
                        match kindOfDay with
                        | Day -> gen (nextNthDaysPatternOfEveryMMonth allDays 2 month)
                        | Weekday -> gen (nextNthDaysPatternOfEveryMMonth weekDays 2 month)
                        | WeekendDay -> gen (nextNthDaysPatternOfEveryMMonth weekendDays 2 month)
                        | DayOfWeek d -> gen (nextNthDaysPatternOfEveryMMonth [d] 2 month)
                    | Third ->
                        match kindOfDay with
                        | Day -> gen (nextNthDaysPatternOfEveryMMonth allDays 3 month)
                        | Weekday -> gen (nextNthDaysPatternOfEveryMMonth weekDays 3 month)
                        | WeekendDay -> gen (nextNthDaysPatternOfEveryMMonth weekendDays 3 month)
                        | DayOfWeek d -> gen (nextNthDaysPatternOfEveryMMonth [d] 3 month)
                    | Fourth ->
                        match kindOfDay with
                        | Day -> gen (nextNthDaysPatternOfEveryMMonth allDays 4 month)
                        | Weekday -> gen (nextNthDaysPatternOfEveryMMonth weekDays 4 month)
                        | WeekendDay -> gen (nextNthDaysPatternOfEveryMMonth weekendDays 4 month)
                        | DayOfWeek d -> gen (nextNthDaysPatternOfEveryMMonth [d] 4 month)
                    | Last ->
                        match kindOfDay with
                        | Day -> gen (lastDaysPatternOfEveryMMonth allDays month)
                        | Weekday -> gen (lastDaysPatternOfEveryMMonth weekDays month)
                        | WeekendDay -> gen (lastDaysPatternOfEveryMMonth weekendDays month)
                        | DayOfWeek d -> gen (lastDaysPatternOfEveryMMonth [d] month)
            | UnknownRecurrence -> failwith "Unknown recurrence"

        recurrences
        |> Seq.filter (fun r -> 
            deletedRecurrences 
            |> Seq.exists(fun a -> 
                // todo: include check on id as well?
                a.Start = r.Start && a.End = r.End)
            |> not)
        |> Seq.map (fun r ->
            let re = 
                recurreceExceptions 
                |> Seq.filter(fun a -> 
                    let (mid, dt) =
                        match a.Recurrence with
                        | RecurreceExceptionInstance(masterSeriesItemId, originalStartDateTime) -> (masterSeriesItemId, originalStartDateTime)
                        | _ -> failwith "Should never happen"                        
                    r.Start = dt && r.Id = mid)
                |> Seq.toList

            if re |> List.length = 1
            then { r with Start = re.Head.Start; End = re.Head.Start.AddSeconds(re.Head.Duration |> float) }
            else r)