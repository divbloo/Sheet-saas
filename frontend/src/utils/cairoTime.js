import {
  CAIRO_TIMEZONE,
  FIRST_CONFIRMATION_EDIT_END_MINUTE,
  FIRST_CONFIRMATION_EDIT_START_MINUTE,
} from "../spreadsheetConfig";

const padDatePart = (value) => String(value).padStart(2, "0");

export const getCairoDateTimeParts = (date = new Date()) => {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: CAIRO_TIMEZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).formatToParts(date);

  return Object.fromEntries(parts.map((part) => [part.type, part.value]));
};

export const formatCairoDateTimeInput = (date = new Date()) => {
  const parts = getCairoDateTimeParts(date);
  const hour = String(Number(parts.hour || 0) % 24);

  return [
    parts.year,
    padDatePart(parts.month),
    padDatePart(parts.day),
  ].join("-") + "T" + [padDatePart(hour), padDatePart(parts.minute)].join(":");
};

export const getDefaultTaxExportPeriod = () => {
  const now = new Date();
  const parts = getCairoDateTimeParts(now);
  const date = [parts.year, padDatePart(parts.month), padDatePart(parts.day)].join("-");

  return {
    from: date + "T00:00",
    to: formatCairoDateTimeInput(now),
  };
};

export const getCairoDateDaysAgo = (daysAgo) => {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: CAIRO_TIMEZONE,
    year: "numeric",
    month: "numeric",
    day: "numeric",
  }).formatToParts(new Date());
  const year = Number(parts.find((part) => part.type === "year")?.value || 1970);
  const month = Number(parts.find((part) => part.type === "month")?.value || 1);
  const day = Number(parts.find((part) => part.type === "day")?.value || 1);
  const date = new Date(year, month - 1, day, 12, 0, 0, 0);

  date.setDate(date.getDate() - daysAgo);
  return date;
};

export const cairoTimeMinutes = () => {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: CAIRO_TIMEZONE,
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).formatToParts(new Date());
  const hour = Number(parts.find((part) => part.type === "hour")?.value || 0) % 24;
  const minute = Number(parts.find((part) => part.type === "minute")?.value || 0);

  return (hour * 60) + minute;
};

export const isFirstConfirmationEditOpen = () => {
  const minutes = cairoTimeMinutes();

  return minutes >= FIRST_CONFIRMATION_EDIT_START_MINUTE && minutes <= FIRST_CONFIRMATION_EDIT_END_MINUTE;
};
