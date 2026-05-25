using System;
using System.Collections.Generic;
using System.Linq;


class CalculatedRecord
{
    // Identity
    public string EmpCode { get; set; }
    public string FullName { get; set; }
    public string Department { get; set; }
    public string Date { get; set; }

    // Raw punches
    public string ClockIn { get; set; }
    public string ClockOut { get; set; }
    public string BreakOut { get; set; }
    public string BreakIn { get; set; }

    // Computed attendance values (in minutes internally)
    public double TotalWorkedMinutes { get; set; }
    public double BreakMinutes { get; set; }
    public double NetWorkedMinutes { get; set; }
    public double RegularMinutes { get; set; }
    public double TotalOTMinutes { get; set; }
    public double OT1Minutes { get; set; }
    public double OT2Minutes { get; set; }
    public double OT3Minutes { get; set; }
    public double EarlyInMinutes { get; set; }
    public double LateOutMinutes { get; set; }
    public double LateMinutes { get; set; }
    public double EarlyOutMinutes { get; set; }
    public bool IsAbsent { get; set; }
    public bool IsUnscheduled { get; set; }

    // Display helpers — hours rounded to 2 decimal places
    public double RegularH => Math.Round(RegularMinutes / 60.0, 2);
    public double TotalOTH => Math.Round(TotalOTMinutes / 60.0, 2);
    public double OT1H => Math.Round(OT1Minutes / 60.0, 2);
    public double OT2H => Math.Round(OT2Minutes / 60.0, 2);
    public double OT3H => Math.Round(OT3Minutes / 60.0, 2);
    public double EarlyInH => Math.Round(EarlyInMinutes / 60.0, 2);
    public double LateOutH => Math.Round(LateOutMinutes / 60.0, 2);
    public double LateH => Math.Round(LateMinutes / 60.0, 2);
    public double EarlyOutH => Math.Round(EarlyOutMinutes / 60.0, 2);
    public double NetWorkedH => Math.Round(NetWorkedMinutes / 60.0, 2);
}

// ══════════════════════════════════════════════════════════════
//  ATTENDANCE CALCULATOR
// ══════════════════════════════════════════════════════════════
static class AttendanceCalculator
{
    // ── Main entry point ──────────────────────────────────────
    public static List<CalculatedRecord> Calculate(
        List<PunchRecord> punches,
        TimetableStore store)
    {
        var results = new List<CalculatedRecord>();

        var groups = punches.GroupBy(p => new
        {
            p.EmpCode,
            Date = p.PunchTime?.Length >= 10
                ? p.PunchTime.Substring(0, 10)
                : p.PunchTime
        });

        foreach (var g in groups)
        {
            var first = g.First();
            var timetable = store.GetForDepartment(first.Department ?? "");
            var rec = CalculateDay(g.Key.EmpCode, g.Key.Date,
                                         first, g.ToList(), timetable);
            results.Add(rec);
        }

        return results
            .OrderBy(r => r.Department)
            .ThenBy(r => r.EmpCode)
            .ThenBy(r => r.Date)
            .ToList();
    }

    // ── Calculate one employee's day ──────────────────────────
    static CalculatedRecord CalculateDay(
        string empCode,
        string date,
        PunchRecord first,
        List<PunchRecord> dayPunches,
        Timetable t)
    {
        var rec = new CalculatedRecord
        {
            EmpCode = empCode,
            FullName = FormatName(first.LastName, first.FirstName),
            Department = first.Department ?? "",
            Date = date
        };

        // Sort all punches chronologically
        var times = dayPunches
            .Select(p => p.PunchTime)
            .Where(p => !string.IsNullOrEmpty(p))
            .OrderBy(p => p)
            .Select(p => ParseDateTime(p))
            .Where(p => p.HasValue)
            .Select(p => p.Value)
            .ToList();

        if (times.Count == 0)
        {
            rec.IsAbsent = true;
            return rec;
        }

        var clockIn = times[0];
        var clockOut = times.Count >= 2 ? times[^1] : (TimeSpan?)null;
        var breakOut = times.Count >= 4 ? times[1] : (TimeSpan?)null;
        var breakIn = times.Count >= 4 ? times[^2] : (TimeSpan?)null;

        rec.ClockIn = FormatHHmm(clockIn);
        rec.ClockOut = clockOut.HasValue ? FormatHHmm(clockOut.Value) : "";
        rec.BreakOut = breakOut.HasValue ? FormatHHmm(breakOut.Value) : "";
        rec.BreakIn = breakIn.HasValue ? FormatHHmm(breakIn.Value) : "";

        // No timetable assigned — just record raw times
        if (t == null)
        {
            if (clockOut.HasValue)
            {
                rec.TotalWorkedMinutes = (clockOut.Value - clockIn).TotalMinutes;
                rec.NetWorkedMinutes = rec.TotalWorkedMinutes;
                rec.RegularMinutes = rec.NetWorkedMinutes;
            }
            return rec;
        }

        var scheduledIn = ParseHHmm(t.CheckIn);
        var scheduledOut = ParseHHmm(t.CheckOut);
        var checkInStart = ParseHHmm(t.CheckInStart);
        var checkInEnd = ParseHHmm(t.CheckInEnd);
        var regularMins = t.RegularHours * 60.0;

        // ── Gross worked time ─────────────────────────────────
        double grossMins = 0;
        if (clockOut.HasValue)
            grossMins = (clockOut.Value - clockIn).TotalMinutes;
        rec.TotalWorkedMinutes = grossMins;

        // ── Break deduction ───────────────────────────────────
        double breakMins = 0;
        foreach (var br in t.Breaks)
        {
            if (br.PunchRequired)
            {
                // Only deduct if employee punched break out and in
                if (breakOut.HasValue && breakIn.HasValue)
                {
                    var actualBreak = (breakIn.Value - breakOut.Value).TotalMinutes;
                    breakMins += Math.Min(Math.Max(actualBreak, 0), br.Duration);
                }
            }
            else
            {
                // Auto-deduct if employee worked through the break window
                var brStart = ParseHHmm(br.StartTime);
                var brEnd = ParseHHmm(br.EndTime);
                if (clockOut.HasValue &&
                    clockIn < brEnd &&
                    clockOut.Value > brStart)
                    breakMins += br.Duration;
            }
        }
        rec.BreakMinutes = breakMins;
        rec.NetWorkedMinutes = Math.Max(grossMins - breakMins, 0);

        // ── Late In ───────────────────────────────────────────
        var lateGrace = scheduledIn.Add(TimeSpan.FromMinutes(t.AllowLateIn));
        if (clockIn > lateGrace)
            rec.LateMinutes = (clockIn - scheduledIn).TotalMinutes;

        // ── Early Out ─────────────────────────────────────────
        if (clockOut.HasValue)
        {
            var earlyGrace = scheduledOut.Subtract(TimeSpan.FromMinutes(t.AllowEarlyOut));
            if (clockOut.Value < earlyGrace)
                rec.EarlyOutMinutes = (scheduledOut - clockOut.Value).TotalMinutes;
        }

        // ── Early In — pre-shift OT ───────────────────────────
        if (t.EarlyInEnabled && clockIn < scheduledIn)
        {
            var earlyMins = (scheduledIn - clockIn).TotalMinutes;
            if (earlyMins >= t.EarlyInMinimum)
                rec.EarlyInMinutes = earlyMins;
            else if (t.EarlyInCountMinimum)
                rec.EarlyInMinutes = t.EarlyInMinimum; // count minimum if enabled
        }

        // ── Late Out — post-shift OT ──────────────────────────
        if (t.LateOutEnabled && clockOut.HasValue && clockOut.Value > scheduledOut)
        {
            var lateMins = (clockOut.Value - scheduledOut).TotalMinutes;
            if (lateMins >= t.LateOutMinimum)
                rec.LateOutMinutes = lateMins;
            else if (t.LateOutCountMinimum)
                rec.LateOutMinutes = t.LateOutMinimum;
        }

        // ── Regular hours ─────────────────────────────────────
        // Regular = net worked, capped at regularHours/day
        rec.RegularMinutes = Math.Min(
            Math.Max(rec.NetWorkedMinutes, 0),
            regularMins);

        // ── Total OT ─────────────────────────────────────────
        double totalOT = Math.Max(rec.NetWorkedMinutes - regularMins, 0);

        // Add early-in to OT bucket if configured
        if (t.EarlyInEnabled &&
           (t.EarlyInAssignTo == "Normal OT" || t.EarlyInAssignTo.StartsWith("OT")))
            totalOT += rec.EarlyInMinutes;

        // Add late-out to OT bucket if configured
        if (t.LateOutEnabled &&
           (t.LateOutAssignTo == "Normal OT" || t.LateOutAssignTo.StartsWith("OT")))
            totalOT += rec.LateOutMinutes;

        // Cap at max OT duration
        if (t.MaxOTEnabled)
            totalOT = Math.Min(totalOT, t.MaxOTMinutes);

        rec.TotalOTMinutes = totalOT;

        // ── OT tier breakdown ─────────────────────────────────
        if (t.OvertimeEnabled && t.OvertimeRules.Count > 0)
        {
            double totalWorkedHours = rec.NetWorkedMinutes / 60.0;

            foreach (var rule in t.OvertimeRules.OrderBy(r => r.HoursFrom))
            {
                if (totalWorkedHours <= rule.HoursFrom) break;

                double tierMins = (Math.Min(totalWorkedHours, rule.HoursTo)
                                  - rule.HoursFrom) * 60.0;
                tierMins = Math.Max(tierMins, 0);

                switch (rule.Name)
                {
                    case "OT1": rec.OT1Minutes = tierMins; break;
                    case "OT2": rec.OT2Minutes = tierMins; break;
                    case "OT3": rec.OT3Minutes = tierMins; break;
                }
            }
        }

        // ── Unscheduled flag ──────────────────────────────────
        if (clockIn < checkInStart || clockIn > checkInEnd)
            rec.IsUnscheduled = true;

        return rec;
    }

    // ── Helpers ───────────────────────────────────────────────
    public static TimeSpan ParseHHmm(string s)
    {
        if (string.IsNullOrEmpty(s)) return TimeSpan.Zero;
        s = s.Trim();
        var parts = s.Split(':');
        if (parts.Length >= 2 &&
            int.TryParse(parts[0], out int h) &&
            int.TryParse(parts[1], out int m))
            return new TimeSpan(h, m, 0);
        return TimeSpan.Zero;
    }

    static TimeSpan? ParseDateTime(string dt)
    {
        // "yyyy-MM-dd HH:mm:ss"
        if (string.IsNullOrEmpty(dt) || dt.Length < 16) return null;
        return ParseHHmm(dt.Substring(11, 5));
    }

    public static string FormatHHmm(TimeSpan t)
        => $"{(int)t.TotalHours:D2}:{t.Minutes:D2}";

    static string FormatName(string last, string first)
    {
        last = last?.Trim() ?? "";
        first = first?.Trim() ?? "";
        if (string.IsNullOrEmpty(last) && string.IsNullOrEmpty(first)) return "";
        if (string.IsNullOrEmpty(last)) return first;
        if (string.IsNullOrEmpty(first)) return last;
        return $"{last}, {first}";
    }
}
