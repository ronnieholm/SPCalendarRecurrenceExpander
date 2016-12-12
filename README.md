SPCalendarRecurrenceExpander
============================

SPCalendarRecurrenceExpander turns each SharePoint calendar recurrence
event into a series of individual events, taking into account
recurrence exceptions.

When to use it
--------------

SharePoint 2007, 2010, 2013, and Online comes with a calendar list
type and web forms for creating either single events or recurrence
events following a large number of patterns.

On-prem SharePoint comes with CAML query support for programmatic
expansion of recurrence events (though the feature is known to be
buggy). With SharePoint Online, however, the expansion feature has
been disabled. Instead, the only out-of-the-box expansion option is to
reverse engineer the internal and undocumented CalendarService.ashx
web service used by the calendar web part.

SPCalendarRecurrenceExpander implements event recurrence expansion by
working directly with the underlying calendar list items, i.e., it
only requires access to the calendar app entries.

Use cases for SPCalendarRecurrenceExpander involve creating <a
href="http://fullcalendar.io">custom views</a> on top of calendars,
either presenting events from a single calendar or aggregating events
across any number of calendars (the built-in SharePoint calendar
supports aggregating up to four calendars). Another use case would be
exposing the expanded recurrence events via a web service for
JavaScript consumption.

How to get it
-------------

Download the
[package](https://www.nuget.org/packages/SPCalendarRecurrenceExpander)
from NuGet:

    Install-Package SPCalendarRecurrenceExpander

The NuGet package contains a .NET 4.5 assembly for use with SharePoint
Online. For other .NET runtime versions (supporting older versions of
SharePoint), currently you'd have to build the library yourself.

SPCalendarRecurrenceExpander is written in F# which means your project
will have to reference fsharp.core.dll to consume the library. If you
included installed Visual Studio with F# support, fsharp.core.dll is
already installed. Otherwise, you can get the assembly by installing
[this](https://www.nuget.org/packages/FSharp.Core) NuGet package.

How to use it
-------------

The
[Examples](https://github.com/ronnieholm/SPCalendarRecurrenceExpander/tree/master/Examples)
folder contains complete C# and F# examples.

Here's an abbreviated example that makes use of the SharePoint CSOM
API to read all calendar list items. These are then fed into the
expander which returns a list of recurrence instances:

```cs
class Appointment {
    public int Id { get; set; }
    public string Title { get; set; }
    public DateTime Start { get; set; }
    public DateTime End { get; set; }
    // add any custom columns here
}

class Program {
    static void Main(string[] args) {
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
        var expander = new CalendarRecurrenceExpander(
            tz.Information.Bias, 
            tz.Information.DaylightBias);
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
    }
}
```

Supported platforms
-------------------

SPCalendarRecurrenceExpander doesn't depend on any SharePoint assembly
and thus no specific SharePoint version. Provided you can access the
raw calendar list items, the library will work. The library doesn't
work with SharePoint's OData web service because it doesn't expose
each item's FieldValues collection wherein the calendar metadata is
stored.

How it works
------------

When a user creates a recurrence event through the user interface,
SharePoint
[transforms](http://aspnetguru.wordpress.com/2007/06/01/understanding-the-sharepoint-calendar-and-how-to-export-it-to-ical-format)
the event into a set of key/value properties and uses an XML-based
domain specific language to describe recurrences.

SPCalendarRecurrenceExpander consists of a parser for these key/value
properties and the recurrence description language. The output of the
parser is a syntax tree representing the recurrence.

For instance, here's the output for a weekly recurrence event,
repeating every week on Sundays and Thursdays for ten instances:

```fs
Weekly (EveryNthWeekOnDays (1, set [DayOfWeek.Sunday; DayOfWeek.Thursday]), RepeatInstances 10)
```

Another example is monthly recurrendes every third weekend day of the
month, every second month for 999 instances (SharePoint's default
number of instances when a user doesn't explicitly specify an end
time):

```fs
Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Third, WeekendDay, 2), NoExplicitEndRange)
```

These syntax trees show two of about 50 recurrence patterns supported
by SharePoint. Each of these patterns is fed to a recurrence compiler
which "executes" the recurrence program, effectively returning
recurrence instances. Recurrence exceptions, such as deleted or
modified instances, are special types of events which replace regular
recurrence instances.

Please let me know if you find this package helpful.

See also
--------

[Internet Calendaring and Scheduling Core Object Specification](https://www.ietf.org/rfc/rfc2445.txt)
[Data elements and interchange formats – Information interchange – Representation of dates and times](https://en.wikipedia.org/wiki/ISO_8601)
