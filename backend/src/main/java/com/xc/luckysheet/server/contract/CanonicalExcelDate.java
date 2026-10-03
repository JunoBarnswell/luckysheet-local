package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.JsonNode;
import com.xc.luckysheet.server.service.ServiceException;
import java.time.*;
import java.time.format.DateTimeFormatter;

/** UTC calendar semantics shared with formula-engine/excel-date.ts. */
public final class CanonicalExcelDate {
    private static final long DAY = 86400000;
    private static final DateTimeFormatter ISO = DateTimeFormatter.ofPattern("uuuu-MM-dd'T'HH:mm:ss.SSS'Z'");
    private CanonicalExcelDate() { }
    private static long epoch(String system) {
        if (!java.util.Set.of("1900", "1904").contains(system)) throw ServiceException.validation("Fill dateSystem must be 1900 or 1904");
        return LocalDate.of("1904".equals(system) ? 1904 : 1899, "1904".equals(system) ? 1 : 12, "1904".equals(system) ? 1 : 31).atStartOfDay().toInstant(ZoneOffset.UTC).toEpochMilli();
    }
    public static double serial(JsonNode value, String system) {
        if (value.isNumber()) { fromSerial(value.asDouble(), system); return value.asDouble(); }
        if (!value.isTextual() || !value.asText().matches("[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}(\\.[0-9]{1,3})?Z")) throw ServiceException.validation("Date fill requires a canonical UTC timestamp or finite serial");
        try { return toSerial(LocalDateTime.parse(value.asText().substring(0, value.asText().length() - 1), DateTimeFormatter.ISO_LOCAL_DATE_TIME), system); }
        catch (DateTimeException error) { throw ServiceException.validation("Date fill seed is outside the supported calendar"); }
    }
    public static LocalDateTime fromSerial(double serial, String system) {
        if (!Double.isFinite(serial) || "1900".equals(system) && serial == 60) throw ServiceException.validation("Invalid Excel date serial or leap-day sentinel");
        double whole = Math.floor(serial), adjusted = "1900".equals(system) && whole > 60 ? whole - 1 : whole;
        double millis = epoch(system) + adjusted * DAY + Math.round((serial - whole) * DAY);
        if (!Double.isFinite(millis) || Math.abs(millis) > 8640000000000000d) throw ServiceException.validation("Excel date is outside the calendar budget");
        return checked(LocalDateTime.ofInstant(Instant.ofEpochMilli((long) millis), ZoneOffset.UTC));
    }
    private static LocalDateTime checked(LocalDateTime date) {
        if (date.getYear() < 1 || date.getYear() > 9999) throw ServiceException.validation("Excel date must use year 1 through 9999");
        return date;
    }
    private static double toSerial(LocalDateTime date, String system) {
        double serial = (checked(date).toInstant(ZoneOffset.UTC).toEpochMilli() - epoch(system)) / (double) DAY;
        return "1900".equals(system) && serial >= 60 ? serial + 1 : serial;
    }
    public static String iso(double serial, String system) { return ISO.format(fromSerial(serial, system)); }
    public static double shift(double serial, double amount, String unit, String system) {
        if (!Double.isFinite(amount) || Math.abs(amount) > 3652058) throw ServiceException.validation("Date fill step exceeds the supported calendar");
        var date = fromSerial(serial, system);
        try {
            if ("day".equals(unit)) return toSerial(LocalDateTime.ofInstant(Instant.ofEpochMilli((long) (date.toInstant(ZoneOffset.UTC).toEpochMilli() + amount * DAY)), ZoneOffset.UTC), system);
            if ("weekday".equals(unit)) {
                long remaining = (long) Math.abs(amount); int direction = amount < 0 ? -1 : 1;
                while (remaining > 0 && date.getDayOfWeek().getValue() > 5) { date = date.plusDays(direction); if (date.getDayOfWeek().getValue() <= 5) remaining--; }
                date = date.plusDays(remaining / 5 * 7 * direction); remaining %= 5;
                while (remaining > 0) { date = date.plusDays(direction); if (date.getDayOfWeek().getValue() <= 5) remaining--; }
                return toSerial(date, system);
            }
            if (!java.util.Set.of("month", "year").contains(unit) || amount != Math.rint(amount)) throw ServiceException.validation("Date fill calendar unit/step is invalid");
            var target = date.plusMonths((long) amount * ("year".equals(unit) ? 12 : 1));
            checked(target.withDayOfMonth(1).plusMonths(1));
            return toSerial(target, system);
        } catch (DateTimeException | ArithmeticException error) { throw ServiceException.validation("Date fill result is outside the supported calendar"); }
    }
}
