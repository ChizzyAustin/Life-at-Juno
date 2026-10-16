# Life at Juno

The community events calendar for Juno Karimu Ikotun (Sabo, Yaba) and Juno Emily Akinola (Akoka, Yaba): what's on today, this week and this month.

Served with GitHub Pages from `index.html` on the `main` branch.

Calendar by Chidiverse Consulting Ltd.

## Event check-in

`checkin/` is the room-number check-in page members open from the QR code at each event (`checkin/?b=karimu` or `checkin/?b=emily` preselects the building). It picks today's event from the calendar and posts each check-in to a Google Sheet through the Apps Script web app in `apps-script/Code.gs`. Paste the web app URL into `CHECKIN_URL` in `checkin/index.html`.
