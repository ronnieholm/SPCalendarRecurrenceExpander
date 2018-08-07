TODO
====

* Adopt terminology of [ical specification](https://tools.ietf.org/html/rfc5545) in code. See section 3.8.5.3. Recurrence rules for a text-based serialization format.
* Implement custom mapper such that not only Start date and End date of recurring appointment gets returned. Some inspiration may be found [here](https://github.com/ronnieholm/SPCalendarRecurrenceExpander/issues/16).
* To work around time zones and Daylight Saving Times, perhapes use [IANA time zones database](https://www.iana.org/time-zones). See [this](https://www.youtube.com/watch?v=Vwd3pduVGKY) talk for inspiration.
