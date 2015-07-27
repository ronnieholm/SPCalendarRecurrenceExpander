module SPCalendarRecurrenceExpander.Tests.ParserTests

open System
open System.Collections.Generic
open Xunit
open Swensen.Unquote
open Holm.SPCalendarRecurrenceExpander

let sut = Parser()

[<Fact>]
let ``parse with WindowEnd event ending at midnight``() =
    let a = Dictionary<string, obj>()
    a.Add("ID", 15)
    a.Add("EventDate", DateTime(2014, 8, 25, 0, 0, 0))
    a.Add("EndDate", DateTime(2014, 8, 29, 0, 0, 0))
    a.Add("Duration", "0")
    a.Add("fRecurrence", true)
    a.Add("EventType", 1)
    a.Add("RecurrenceData", """<recurrence><rule><firstDayOfWeek>mo</firstDayOfWeek><repeat><daily dayFrequency="1" /></repeat><windowEnd>2014-08-30T00:00:00Z</windowEnd></rule></recurrence>""")

    let p = sut.Parse(a)
    test <@ p.Recurrence = Daily (EveryNthDay 1, ExplicitEnd(DateTime(2014, 8, 30, 0, 0, 0))) @>

    // ignores the windowEnd date in place of endDate during parsing
    a.["EndDate"] <- DateTime(2014, 8, 29, 1, 0, 0)
    a.["RecurrenceData"] <- """<recurrence><rule><firstDayOfWeek>mo</firstDayOfWeek><repeat><daily dayFrequency="1" /></repeat><windowEnd>2014-08-29T08:00:00Z</windowEnd></rule></recurrence>"""
    let p = sut.Parse(a)
    test <@ p.Recurrence = Daily (EveryNthDay 1, ExplicitEnd(DateTime(2014, 8, 29, 1, 0, 0))) @>