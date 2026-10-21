# PhotoNotes

Phone web app for property walk-through photos. `index.html` is the whole app (GitHub Pages). Tapping Next Photo resizes the photo, stores it on the phone, and clears the screen immediately. End Session and New Unit use that same save when a category, a note, or a photo is still on the screen, then leave. A queue in IndexedDB uploads each item until the Apps Script confirms that client id, so a photo is not dropped when the inspector leaves the property or loses signal.

The Apps Script source is `google-apps-script.js`, matching the live web app (Version 18). It is pasted into the Sheet; it is not deployed from this repo.

## Update Apps Script (same URL)

1. Open the walk-through Google Sheet → **Extensions → Apps Script**.
2. Replace the project contents with `google-apps-script.js`.
3. `CONFIG` is set for the live sheet:
   - Rows tab: `PhotoNotes` (pinned so the old Inspections tab is not used).
   - Follow-ups: `FollowUp`.
   - Categories: `Categories`, or unique values from the rows tab.
   - Row headers: Property, Unit, Location, User, Category, Notes, Photo link, Date, Time. User, Notes, and Photo link map to inspector, note, and photoUrl. The script fills Time when it saves the row.
   - Drive: a `PhotoNotes` folder, then a folder per property and a subfolder per unit.
4. **Deploy → Manage deployments → Edit** (pencil) on the existing deployment → **Version: New version** → Deploy. That keeps the web app URL. A new deployment would change it. Skip this if the project already is Version 18.

The phone keeps posting to the existing `/exec` URL. If a deployment answers `Unknown action` for `submitPhotoRow`, the queue falls back to `uploadPhoto` then `submitRows` and keeps retrying.

## What Jeff does not need to change

The page, the Settings screen, and the saved web app URL stay as they are. End Session does not delete photos still waiting on the phone. With an empty note and no photo, End Session and New Unit do not add a row.
