# meet-api

Booking API behind https://rahulch.site/meet/. Reads free/busy from Outlook (Microsoft Graph via the Maton gateway) and creates the event with the guest as attendee, so Outlook sends the invite.

- `GET /api/config`, `GET /api/slots?type=15|30|60`, `POST /api/book`, `GET|POST /cancel?t=...`
- Availability rules: `src/config.js`
- Env: `MATON_KEY`, `SIGNING_SECRET`, `PUBLIC_URL`, optional `MEETING_LINK`
