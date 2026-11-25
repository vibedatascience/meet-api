// Edit this file to change availability. Times are in HOST_TZ.
export const config = {
  hostName: "Rahul Chaudhary",
  hostEmail: "rahul.chaudhary@outlook.com",
  hostTz: "America/New_York",

  eventTypes: {
    "15": { minutes: 15, title: "15 min intro" },
    "30": { minutes: 30, title: "30 min chat" },
    "60": { minutes: 60, title: "60 min deep dive" },
  },

  // weekday: 1 = Monday ... 7 = Sunday. Each entry is a list of [start, end] windows.
  hours: {
    1: [["10:00", "17:00"]],
    2: [["10:00", "17:00"]],
    3: [["10:00", "17:00"]],
    4: [["10:00", "17:00"]],
    5: [["10:00", "16:00"]],
  },

  bufferMinutes: 10,      // gap kept before and after every existing event
  minNoticeHours: 12,     // earliest bookable time from now
  maxDaysAhead: 30,       // how far out guests can book
  maxPerDay: 4,           // cap on bookings made through this page per day
  holidayCalendars: ["US Holidays"], // whole days blocked
  blockingShowAs:["busy", "tentative", "oof", "workingElsewhere"],

  allowedOrigins: [
    "https://rahulch.site",
    "https://www.rahulch.site",
    "http://localhost:8000",
    "http://127.0.0.1:8000",
  ],
};
