module SPCalendarRecurrenceExpander.Tests.CompilerTests

open System
open Xunit
open Swensen.Unquote
open Holm.SPCalendarRecurrenceExpander

let dt year month day = DateTime(year, month, day)
let dt2 year month day hour minute = DateTime(year, month, day, hour, minute, 0)
let tm h m = TimeSpan(h, m, 0)

let basic = 
    { Id = 123
      Start = dt2 2014 8 1 11 30
      End = DateTime.MaxValue
      Duration = 7200L
      Recurrence = UnknownRecurrence }

let sut = Compiler()

let compareDate (i: RecurrenceInstance) (dt: DateTime) =
    i.Start.Year = dt.Year && i.Start.Month = dt.Month && i.Start.Day = dt.Day &&
    i.End.Year = dt.Year && i.End.Month = dt.Month && i.End.Day = dt.Day

let compareTime (i: RecurrenceInstance) (tm1: TimeSpan) (tm2: TimeSpan) =
    i.Start.Hour = tm1.Hours && 
    i.Start.Minute = tm1.Minutes &&
    i.End.Hour = tm2.Hours &&
    i.End.Minute = tm2.Minutes

let compareMultiple (instances: seq<RecurrenceInstance>) (dateTimes: DateTime list) ts1 ts2 = 
    Seq.zip instances dateTimes |> Seq.forall (fun (i, dt) -> 
        compareDate i dt && compareTime i ts1 ts2)

[<Fact>]
let ``daily every n'th day``() =
    // start date is by definition always included in recurrences
    // a better test would cross month-boundaries
    let a = { basic with Start = dt2 2014 8 1 11 30; End = dt2 2017 4 25 13 30; Recurrence = Daily(EveryNthDay 1, ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r ([1..3] |> List.map(fun d -> dt 2014 8 d)) (tm 11 30) (tm 13 30) @> 

    let a = { basic with Start = dt2 2014 8 1 11 30; End = dt2 2017 4 25 13 30; Recurrence = Daily(EveryNthDay 3, ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r ([1..3..7] |> List.map(fun d -> dt 2014 8 d)) (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``daily every week day``() =
    let a1 = { basic with Start = dt2 2014 8 1 11 30; End = dt2 2018 5 30 13 30; Recurrence = Daily(EveryWeekDay, ImplicitEnd) }
    let r1 = sut.Compile(a1, [], [])
    test <@ compareMultiple r1 ([1;4;5;6;7;8] |> List.map(fun d -> dt 2014 8 d)) (tm 11 30) (tm 13 30) @>

    // start date isn't a recurrence instance
    let a2 = { basic with Start = dt2 2014 8 2 11 30; End = dt2 2018 5 30 13 30; Recurrence = Daily(EveryWeekDay, ImplicitEnd) }
    let r2 = sut.Compile(a2, [], [])
    test <@ compareMultiple r2 ([4;5;6;7;8;11] |> List.map(fun d -> dt 2014 8 d)) (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``weekly every n'th day on days``() =
    let a1 = { basic with Start = dt2 2014 8 2 11 30; End = dt2 2020 12 17 13 30; Recurrence = Weekly (EveryNthWeekOnDays (1, set [DayOfWeek.Tuesday; DayOfWeek.Thursday; DayOfWeek.Saturday]),ImplicitEnd) }
    let r1 = sut.Compile(a1, [], []) 
    test <@ compareMultiple r1 ([2;5;7;9;12;14;16;19;21;23] |> List.map(fun d -> dt 2014 8 d)) (tm 11 30) (tm 13 30) @>
    
    // start date isn't necessarily the day of the first Recurrence
    let a2 = { basic with Start = dt2 2014 9 22 11 30; End = dt2 2020 12 17 13 30; Recurrence = Weekly (EveryNthWeekOnDays (1, set [DayOfWeek.Sunday; DayOfWeek.Thursday]), RepeatInstances 10) }
    let r2 = sut.Compile(a2, [], [])     
    test <@ compareMultiple r2 [(dt 2014 9 25); (dt 2014 9 28); (dt 2014 10 2); (dt 2014 10 5); (dt 2014 10 9); (dt 2014 10 12); (dt 2014 10 16); (dt 2014 10 19); (dt 2014 10 23); (dt 2014 10 26)] (tm 11 30) (tm 13 30) @>

    // example of when the week of start date doesn't contain Thursday. One hit must be made before skipping weeks
    let a3 = { basic with Start = dt2 2014 8 6 11 30; End = dt2 2072 1 7 13 30; Recurrence = Weekly (EveryNthWeekOnDays (3, set [DayOfWeek.Wednesday; DayOfWeek.Thursday]),ImplicitEnd) }
    let r3 = sut.Compile(a3, [], [])
    test <@ compareMultiple r3 [(dt 2014 8 6); (dt 2014 8 7); (dt 2014 8 27); (dt 2014 8 28); (dt 2014 9 17); (dt 2014 9 18)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every n'th day of every m'th month``() =
    // if a month has less days than n, SharePoint default to last day of the month
    let a1 = { basic with Start = dt2 2014 8 31 11 30; End = dt2 2097 10 12 13 30; Recurrence = Monthly (EveryNthDayOfEveryMthMonth (31, 1), ImplicitEnd) }
    let r1 = sut.Compile(a1, [], [])
    test <@ compareMultiple r1 [(dt 2014 8 31); (dt 2014 9 30); (dt 2014 10 31); (dt 2014 11 30); (dt 2014 12 31)] (tm 11 30) (tm 13 30) @>

    let a2 = { basic with Start = dt2 2014 8 12 11 30; End = dt2 2165 5 12 13 30; Recurrence = Monthly (EveryNthDayOfEveryMthMonth (12, 3), ImplicitEnd) }
    let r2 = sut.Compile(a2, [], [])
    test <@ compareMultiple r2 [(dt 2014 8 12); (dt 2014 11 12); (dt 2015 2 12)] (tm 11 30) (tm 13 30) @>

    // start is date not part of recurrence
    let a3 = { basic with Start = dt2 2014 8 10 11 30; End = dt2 2165 5 12 13 30; Recurrence = Monthly (EveryNthDayOfEveryMthMonth (12, 3), ImplicitEnd) }
    let r3 = sut.Compile(a3, [], [])
    test <@ compareMultiple r3 [(dt 2014 8 12); (dt 2014 11 12); (dt 2015 2 12)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every first day of the month``() =
    let a = { basic with Start = dt2 2014 8 1 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (First, Day, 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 1); (dt 2014 9 1); (dt 2014 10 1)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 1 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (First, Day, 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 1); (dt 2014 11 1); (dt 2015 2 1)] (tm 11 30) (tm 13 30) @>

    // start is date not part of recurrence. Notice how recurrences jump even before finding the first one
    let a = { basic with Start = dt2 2014 9 22 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (First, Day, 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 12 1); (dt 2015 3 1); (dt 2015 6 1)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every first weekday of the month``() =
    let a = { basic with Start = dt2 2014 8 1 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (First, Weekday, 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 1); (dt 2014 9 1); (dt 2014 10 1)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 1 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (First, Weekday, 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 1); (dt 2014 11 3); (dt 2015 2 2)] (tm 11 30) (tm 13 30) @>

    // start date before only recurrence date of month
    let a = { basic with Start = dt2 2014 11 01 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (First, Weekday, 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 11 03); (dt 2014 12 01); (dt 2015 01 01)] (tm 11 30) (tm 13 30) @>

    // start date after only recurrence date of month
    let a = { basic with Start = dt2 2014 7 28 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (First, Weekday, 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 10 1); (dt 2015 1 1); (dt 2015 4 1)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every first weekend day of the month``() =
    let a = { basic with Start = dt2 2014 8 2 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (First, WeekendDay, 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 2); (dt 2014 9 6); (dt 2014 10 4)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 2 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (First, WeekendDay, 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 2); (dt 2014 11 1); (dt 2015 2 1)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every first specific day of the month``() =
    let a = { basic with Start = dt2 2014 8 6 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (First, DayOfWeek(DayOfWeek.Wednesday), 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 6); (dt 2014 9 3); (dt 2014 10 1)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 6 11 30;  End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (First, DayOfWeek(DayOfWeek.Wednesday), 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])    
    test <@ compareMultiple r [(dt 2014 8 6); (dt 2014 11 5); (dt 2015 2 4)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every second day of the month``() =
    let a = { basic with Start = dt2 2014 8 2 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Second, Day, 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 2); (dt 2014 9 2); (dt 2014 10 2)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 2 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Second, Day, 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 2); (dt 2014 11 2); (dt 2015 2 2)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every second weekday of the month``() =
    let a = { basic with Start = dt2 2014 8 4 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Second, Weekday, 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 4); (dt 2014 9 2); (dt 2014 10 2)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 4 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Second, Weekday, 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 4); (dt 2014 11 4); (dt 2015 2 3)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every second weekend day of the month``() =
    let a = { basic with Start = dt2 2014 8 3 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Second, WeekendDay, 1), ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 3); (dt 2014 9 7); (dt 2014 10 5)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 3 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Second, WeekendDay, 3), ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 3); (dt 2014 11 2); (dt 2015 2 7)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every second specific day of the month``() =
    let a = { basic with Start = dt2 2014 8 13 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Second, DayOfWeek(DayOfWeek.Wednesday), 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 13); (dt 2014 9 10); (dt 2014 10 8)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 13 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Second, DayOfWeek(DayOfWeek.Wednesday), 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 13); (dt 2014 11 12); (dt 2015 2 11)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every third day of the month``() =
    let a = { basic with Start = dt2 2014 8 3 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Third, Day, 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 3); (dt 2014 9 3); (dt 2014 10 3)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 3 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Third, Day, 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 3); (dt 2014 11 3); (dt 2015 2 3)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every third weekday of the month``() =
    let a = { basic with Start = dt2 2014 8 5 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Third, Weekday, 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 5); (dt 2014 9 3); (dt 2014 10 3)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 5 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Third, Weekday, 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 5); (dt 2014 11 5); (dt 2015 2 4)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every third weekend day of the month``() =
    let a = { basic with Start = dt2 2014 8 9 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Third, WeekendDay, 1), ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 9); (dt 2014 9 13); (dt 2014 10 11)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 9 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Third, WeekendDay, 3), ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 9); (dt 2014 11 8); (dt 2015 2 8)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every third specific day of the month``() =
    let a = { basic with Start = dt2 2014 8 20 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Third, DayOfWeek(DayOfWeek.Wednesday), 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 20); (dt 2014 9 17); (dt 2014 10 15)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 20 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Third, DayOfWeek(DayOfWeek.Wednesday), 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 20); (dt 2014 11 19); (dt 2015 2 18)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every fourth day of the month``() =
    let a = { basic with Start = dt2 2014 8 4 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Fourth, Day, 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 4); (dt 2014 9 4); (dt 2014 10 4)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 4 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Fourth, Day, 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 4); (dt 2014 11 4); (dt 2015 2 4)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every fourth weekday of the month``() =
    let a = { basic with Start = dt2 2014 8 6 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Fourth, Weekday, 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 6); (dt 2014 9 4); (dt 2014 10 6)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 6 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Fourth, Weekday, 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 6); (dt 2014 11 6); (dt 2015 2 5)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every fourth weekend day of the month``() =
    let a = { basic with Start = dt2 2014 8 10 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Fourth, WeekendDay, 1), ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 10); (dt 2014 9 14); (dt 2014 10 12)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 10 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Fourth, WeekendDay, 3), ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 10); (dt 2014 11 9); (dt 2015 2 14)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every fourth specific day of the month``() =
    let a = { basic with Start = dt2 2014 8 27 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Fourth, DayOfWeek(DayOfWeek.Wednesday), 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 27); (dt 2014 9 24); (dt 2014 10 22)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 27 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Fourth, DayOfWeek(DayOfWeek.Wednesday), 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 27); (dt 2014 11 26); (dt 2015 2 25)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every last day of the month``() =
    let a1 = { basic with Start = dt2 2014 8 31 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Last, Day, 1),ImplicitEnd) }
    let r1 = sut.Compile(a1, [], [])
    test <@ compareMultiple r1 [(dt 2014 8 31); (dt 2014 9 30); (dt 2014 10 31)] (tm 11 30) (tm 13 30) @>

    let a2 = { basic with Start = dt2 2014 8 31 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Last, Day, 3),ImplicitEnd) }
    let r2 = sut.Compile(a2, [], [])
    test <@ compareMultiple r2 [(dt 2014 8 31); (dt 2014 11 30); (dt 2015 2 28)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every last weekday of the month``() =
    let a = { basic with Start = dt2 2014 8 29 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Last, Weekday, 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 29); (dt 2014 9 30); (dt 2014 10 31)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 29 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Last, Weekday, 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 29); (dt 2014 11 28); (dt 2015 2 27)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every last weekend day of the month``() =
    let a = { basic with Start = dt2 2014 8 31 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Last, WeekendDay, 1), ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 31); (dt 2014 9 28); (dt 2014 10 26)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 31 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Last, WeekendDay, 3), ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 31); (dt 2014 11 30); (dt 2015 2 28)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``monthly every last specific day of the month``() =
    let a = { basic with Start = dt2 2014 8 27 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Last, DayOfWeek(DayOfWeek.Wednesday), 1),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 27); (dt 2014 9 24); (dt 2014 10 29)] (tm 11 30) (tm 13 30) @>

    let a = { basic with Start = dt2 2014 8 27 11 30; End = dt2 2020 1 1 13 30; Recurrence = Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Last, DayOfWeek(DayOfWeek.Wednesday), 3),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2014 8 27); (dt 2014 11 26); (dt 2015 2 25)] (tm 11 30) (tm 13 30) @>

// yearly are mostly special cases on months with skipmonths = 12
[<Fact>]
let ``yearly every n'th day of every m month``() =
    // when n is larger than days in month, select last day of month
    let a = { basic with Start = dt2 2015 2 28 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryNthDayOfEveryMMonth (31, 2), ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 28); (dt 2016 2 29); (dt 2017 2 28)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every first day of specific month``() =
    let a1 = { basic with Start = dt2 2015 2 1 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (First, Day, 2),ImplicitEnd) }
    let r1 = sut.Compile(a1, [], [])
    test <@ compareMultiple r1 [(dt 2015 2 1); (dt 2016 2 1); (dt 2017 2 1)] (tm 11 30) (tm 13 30) @>

    // start date is in previous year before recurrence instance
    let a2 = { basic with Start = dt2 2014 7 12 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (First, Day, 2),ImplicitEnd) }
    let r2 = sut.Compile(a2, [], [])
    test <@ compareMultiple r2 [(dt 2015 2 1); (dt 2016 2 1); (dt 2017 2 1)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every first weekday of specific month``() =
    let a = { basic with Start = dt2 2015 2 2 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (First, Weekday, 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 2); (dt 2016 2 1); (dt 2017 2 1)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every first weekend day of specific month``() =
    let a1 = { basic with Start = dt2 2015 2 1 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (First, WeekendDay, 2),ImplicitEnd) }
    let r1 = sut.Compile(a1, [], [])
    test <@ compareMultiple r1 [(dt 2015 2 1); (dt 2016 2 6); (dt 2017 2 4)] (tm 11 30) (tm 13 30) @>

    // start date before first instance
    let a1 = { basic with Start = dt2 2014 9 25 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (First, WeekendDay, 10),ImplicitEnd) }
    let r1 = sut.Compile(a1, [], [])
    test <@ compareMultiple r1 [(dt 2014 10 4); (dt 2015 10 3); (dt 2016 10 1)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every first specific day of specific month``() =
    let a = { basic with Start = dt2 2015 2 4 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (First, DayOfWeek(DayOfWeek.Wednesday), 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 4); (dt 2016 2 3); (dt 2017 2 1)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every second day of specific month``() =
    let a = { basic with Start = dt2 2015 2 2 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Second, Day, 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 2); (dt 2016 2 2); (dt 2017 2 2)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every second weekday of specific month``() =
    let a = { basic with Start = dt2 2015 2 3 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Second, Weekday, 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 3); (dt 2016 2 2); (dt 2017 2 2)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every second weekend day of specific month``() =
    let a = { basic with Start = dt2 2015 2 7 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Second, WeekendDay, 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 7); (dt 2016 2 7); (dt 2017 2 5)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every second specific day of specific month``() =
    let a = { basic with Start = dt2 2015 2 11 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Second, DayOfWeek(DayOfWeek.Wednesday), 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 11); (dt 2016 2 10); (dt 2017 2 8)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every third day of specific month``() =
    let a = { basic with Start = dt2 2015 2 3 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Third, Day, 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 3); (dt 2016 2 3); (dt 2017 2 3)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every third weekday of specific month``() =
    let a = { basic with Start = dt2 2015 2 4 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Third, Weekday, 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 4); (dt 2016 2 3); (dt 2017 2 3)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every third weekend day of specific month``() =
    let a = { basic with Start = dt2 2015 2 8 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Third, WeekendDay, 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 8); (dt 2016 2 13); (dt 2017 2 11)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every third specific day of specific month``() =
    let a = { basic with Start = dt2 2015 2 18 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Third, DayOfWeek(DayOfWeek.Wednesday), 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 18); (dt 2016 2 17); (dt 2017 2 15)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every fourth day of specific month``() =
    let a = { basic with Start = dt2 2015 2 4 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Fourth, Day, 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 4); (dt 2016 2 4); (dt 2017 2 4)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every fourth weekday of specific month``() =
    let a = { basic with Start = dt2 2015 2 5 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Fourth, Weekday, 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 5); (dt 2016 2 4); (dt 2017 2 6)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every fourth weekend day of specific month``() =
    let a = { basic with Start = dt2 2015 2 14 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Fourth, WeekendDay, 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 14); (dt 2016 2 14); (dt 2017 2 12)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every fourth specific day of specific month``() =
    let a = { basic with Start = dt2 2015 2 25 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Fourth, DayOfWeek(DayOfWeek.Wednesday), 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 25); (dt 2016 2 24); (dt 2017 2 22)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every last day of specific month``() =
    let a = { basic with Start = dt2 2015 2 28 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Last, Day, 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 28); (dt 2016 2 29); (dt 2017 2 28)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every last weekday of specific month``() =
    let a = { basic with Start = dt2 2015 2 27 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Last, Weekday, 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 27); (dt 2016 2 29); (dt 2017 2 28)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every last weekend day of specific month``() =
    let a = { basic with Start = dt2 2015 2 28 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Last, WeekendDay, 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 28); (dt 2016 2 28); (dt 2017 2 26)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``yearly every last specific day of specific month``() =
    let a = { basic with Start = dt2 2015 2 25 11 30; End = dt2 2020 1 1 13 30; Recurrence = Yearly (EveryQualifierOfKindOfDayMMonth (Last, DayOfWeek(DayOfWeek.Wednesday), 2),ImplicitEnd) }
    let r = sut.Compile(a, [], [])
    test <@ compareMultiple r [(dt 2015 2 25); (dt 2016 2 24); (dt 2017 2 22)] (tm 11 30) (tm 13 30) @>

[<Fact>]
let ``daily recurrence with deleted instance``() = 
    let recurrence =
        { Id = 5
          Start = dt2 2014 8 16 09 00
          End = dt2 2014 8 20 10 30
          Duration = 5400L
          Recurrence = Daily (EveryNthDay 1, RepeatInstances 5) }      
    
    let deletedOccurance = 
        { Id = 6
          Start = dt2 2014 8 18 09 00
          End = dt2 2014 8 18 10 30
          Duration = 5400L
          Recurrence = DeletedRecurrenceInstance(5) }      
   
    let output = sut.Compile((recurrence, [deletedOccurance], []))       
    test <@ output |> Seq.length = 4 @>
    test <@ output |> Seq.filter (fun r -> r.Start.Day = 16) |> Seq.length = 1 @>
    test <@ output |> Seq.filter (fun r -> r.Start.Day = 17) |> Seq.length = 1 @>
    test <@ output |> Seq.filter (fun r -> r.Start.Day = 18) |> Seq.length = 0 @>
    test <@ output |> Seq.filter (fun r -> r.Start.Day = 19) |> Seq.length = 1 @>
    test <@ output |> Seq.filter (fun r -> r.Start.Day = 20) |> Seq.length = 1 @>

[<Fact>]
let ``daily recurrence with recurrence exception``() =
    let recurrence =
        { Id = 5
          Start = dt2 2014 8 16 09 00
          End = dt2 2014 8 20 10 30
          Duration = 5400L
          Recurrence = Daily (EveryNthDay 1, RepeatInstances 5) }      

    let recurrenceException = 
        { Id = 6
          Start = dt2 2014 8 18 10 00
          End = dt2 2014 8 18 11 00
          Duration = 3600L
          Recurrence = ModifiedRecurreceInstance(5, dt2 2014 8 18 09 00) }  
          
    let output = sut.Compile(recurrence, [], [recurrenceException]) |> Seq.toList

    test <@ output |> List.length = 5 @>
    test <@ output.[0].Start = dt2 2014 8 16 09 00 && output.[0].End = dt2 2014 8 16 10 30 @>
    test <@ output.[1].Start = dt2 2014 8 17 09 00 && output.[1].End = dt2 2014 8 17 10 30 @>
    test <@ output.[2].Start = dt2 2014 8 18 10 00 && output.[2].End = dt2 2014 8 18 11 00 @>  // <-- recurrence exception
    test <@ output.[3].Start = dt2 2014 8 19 09 00 && output.[3].End = dt2 2014 8 19 10 30 @>
    test <@ output.[4].Start = dt2 2014 8 20 09 00 && output.[4].End = dt2 2014 8 20 10 30 @>