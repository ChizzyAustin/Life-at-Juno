# Life at Juno

The community events calendar for Juno Karimu Ikotun (Sabo, Yaba) and Juno Emily Akinola (Akoka, Yaba): what's on today, this week and this month.

Served with GitHub Pages from `index.html` on the `main` branch.

Calendar by Chidiverse Consulting Ltd.

## Event check-in

`checkin/` is the room-number check-in page members open from the QR code at each event (`checkin/?b=karimu` or `checkin/?b=emily` preselects the building). It picks today's event from the calendar and posts each check-in to a Google Sheet through the Apps Script web app in `apps-script/Code.gs`. Paste the web app URL into `CHECKIN_URL` in `checkin/index.html`.

## RSVP

`rsvp/` is the RSVP page for Juno at Five. Residents give their building and room; guests give an organisation and, if they like, a phone number. `rsvp/?g=sankore` opens the guest version for the Sankore team (Sankore at 16 & Juno at 5). Replies go to the "RSVPs" tab of the same Google Sheet through the same Apps Script web app as check-in (`apps-script/Code.gs`). A second reply from the same person updates their row. After pasting the updated `Code.gs`, redeploy with Deploy > Manage deployments > Edit > Version: New version so the web app URL stays the same.
