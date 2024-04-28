export namespace Date {
    export type Action = "add" | "subtract";
    export type Unit = (
        "year" | "years" | "y" |
        "month" | "months" | "M" |
        "week" | "weeks" | "w" |
        "day" | "days" | "d" |
        "hour" | "hours" | "h" |
        "minute" | "minutes" | "m" |
        "second" | "seconds" | "s" |
        "millisecond" | "milliseconds" | "ms"
    );
}