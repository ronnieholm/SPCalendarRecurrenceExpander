SPCalendarRecurrenceExpander
============================

SPCalendarRecurrenceExpander turns each SharePoint calendar
recurrence event into a series of individual events, taking into
account recurrence exceptions.

When to use it
--------------

SharePoint 2007, 2010, 2013, and Online comes with a calendar list
type and web forms for creating either single events or recurrence
events following a large number of patterns.

For on-prem SharePoint, Microsoft added special CAML query support for
programmatic expansion of recurrence events (although the feature is
buggy). With SharePoint Online, on the other hand, recurrence
expansion through CAML query is unsupported. You could reverse
engineer the internal and undocumented CalendarService.ashx used by
the calendar views.

Or you could use SPCalendarRecurrenceExpander which implements event
recurrence expansion by working with the underlying calendar list
items directly, i.e., it depends only on standard list item access.

Use cases for SPCalendarRecurrenceExpander involve creating custom
views on top of calendars, either presenting events from a single
calendar or aggregating events across any number of calendars. The
built-in SharePoint calendar supports aggregating only up to four
calendars whereas SPCalendarRecurrenceExpander has no upper
limit. Another use case would be exposing the expanded recurrence
events via a web service for JavaScript consumption.

How to get it
-------------

Download the
[package](https://www.nuget.org/packages/SPCalendarRecurrenceExpander)
from NuGet.

The NuGet package contains a .NET 4.5 assembly as that's required
for use with SharePoint Online. For support of other .NET runtime
versions (for older SharePoint versions), currently you'd have to
build the library yourself.

SPCalendarRecurrenceExpander is written in F# which means your
project will have to reference fsharp.core.dll to consume the
library. The fsharp.core.dll assembly is installed by Visual
Studio if you include F# language support or you can get the
assembly by installing
[this](https://www.nuget.org/packages/FSharp.Core.Microsoft.Signed/)
NuGet package.

How to use it
-------------

The
[Examples](https://github.com/ronnieholm/SPCalendarRecurrenceExpander/tree/master/Examples)
folder contains complete C# and F# examples.

Here's an abbreviated example that uses the SharePoint CSOM API
to read all calendar list items. These are then fed into the
expander which returns a list of recurrence instances which are
then merged with the original appointments to produce a final
list of expanded appointments:

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
    }
}
```

Supported platforms
-------------------

SPCalendarRecurrenceExpander doesn't depend on any SharePoint
assembly and thus no specific SharePoint version. Provided you
can access the raw calendar list items, the library will work.

How it works
------------

When a user creates a recurrence event through the user
interface, SharePoint
[transforms](http://aspnetguru.wordpress.com/2007/06/01/understanding-the-sharepoint-calendar-and-how-to-export-it-to-ical-format)
the event into a set of key/value properties and uses an
XML-based domain specific language for describing recurrences.

SPCalendarRecurrenceExpander consists of a parser for these
key/value properties and the recurrence description language. The
output of the parser is a syntax tree describing the recurrence. 

For instance, here's the output for a weekly recurrence event,
repeating every week on Sundays and Thursdays for ten instances:

    Weekly (EveryNthWeekOnDays (1, set [DayOfWeek.Sunday; DayOfWeek.Thursday]), RepeatInstances 10)

Another example is monthly recurrendes every third weekend day of
the month, every second month for 999 instances (SharePoint
default when user doesn't explicitly specify an end time):

    Monthly (EveryQualifierOfKindOfDayEveryMthMonth (Third, WeekendDay, 2), NoExplicitEndRange)

These syntax trees show two of about 50 recurrence patterns
supported by SharePoint. Each of these patterns is fed to a
compiler which executes the language, effectively returning
recurrence instances. Recurrence exceptions, such as deleted or
modified instances, are special types of events that replace
regular recurrence instances.