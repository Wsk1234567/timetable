# My Timetable

Interactive MMU timetable with browser-local activities, weekly recurrence, and
optional Google Sheet sync across devices.

## Google Sheet sync setup

1. Create a Google Cloud project.
2. Enable **Google Drive API** and **Google Sheets API**.
3. Configure the Google Auth Platform as an external app and add the owner as a
   test user.
4. Create a **Web application** OAuth client with these JavaScript origins:
   - `https://wsk1234567.github.io`
   - `https://my-time-canvas-1123.electburr17.chatgpt.site`
   - `http://localhost:4173`
5. Put the public OAuth client ID in `dist/config.js`. Never add a client secret.

The app requests only `drive.file` and `spreadsheets`. It creates a visible
Google Sheet named `我的时间表资料`, keeps an append-only `Activity Log`, and
continues using `localStorage` as its offline cache.
