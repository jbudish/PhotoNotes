# PhotoNotes

Phone web app for property walk-through photos. `index.html` is the whole app (GitHub Pages). Tapping Next Photo resizes the photo, stores it on the phone, and clears the screen immediately. End Session and New Unit use that same save when a category, a note, or a photo is still on the screen, then leave. A queue in IndexedDB uploads each item until the Apps Script confirms that client id, so a photo is not dropped when the inspector leaves the property or loses signal.

The Apps Script source is `google-apps-script.js`. It is pasted into the Sheet; it is not deployed from this repo.

## Update Apps Script (same URL)

1. Open the walk-through Google Sheet → **Extensions → Apps Script**.
2. Replace the project contents with `google-apps-script.js`.
3. `CONFIG` is set for the live sheet:
   - Rows tab: `PhotoNotes` (pinned so the old Inspections tab is not used).
   - Follow-ups: `FollowUp`.
   - Categories: `Categories`, or unique values from the rows tab.
   - Row headers: Property, Unit, Location, User, Category, Notes, Photo link, Date, Time. User, Notes, and Photo link map to inspector, note, and photoUrl. The script fills Time when it saves the row.
   - Drive: new photos are files directly in folder id `1ejeYoPismA9hAEC6G7V11CqvAGhPHyu-` (`CONFIG.PHOTO_NOTES_FOLDER_ID`; Script Property `photoNotesFolderId` overrides it). The name `PhotoNotes` is used only if that id cannot be opened. The file name is `Property_Unit_Category_0000001.jpg`. No property, unit, or category folders. The first time `photoNotesCounter` is unset, it starts at the highest `*_#######.jpg` number already in that folder.
4. **Deploy → Manage deployments → Edit** (pencil) on the existing deployment → **Version: New version** → Deploy. That keeps the web app URL. A new deployment would change it.

The phone keeps posting to the existing `/exec` URL. If a deployment answers `Unknown action` for `submitPhotoRow`, the queue falls back to `uploadPhoto` then `submitRows` and keeps retrying.

End Session, after saving any open note, waits for this phone’s photos to upload and then sends `endSession` with those photo ids only. The script matches the typed property to a workbook in the Property Campaigns folder. Small spelling differences still match when one workbook is clearly closest (`Millie Benneke`, `Milly Beneke`, `Twin pins`). A weak name can match from the street number and street name, or from the city when that only breaks a tie. A confident match moves that session’s rows onto the workbook’s Photonotes tab and moves that session’s photo files into `[property folder]/Photo Notes` when a property folder exists. Nothing is created under Photo Notes. The file’s property segment is renamed to the matched workbook name (the ` - MB m/yy` suffix dropped). Photos already sitting somewhere else under that property folder are left where they are.

If the match is not confident, nothing moves. The phone shows **Did you mean…?** with the closest workbooks, a search of every workbook in Property Campaigns, and **New property / keep in PhotoNotes**. Tapping a workbook files the session there and stores that typed name as an alias on the hidden `PropertyAliases` tab, so the next inspector who types it skips the prompt. New property stores nothing and leaves the photos on the master tab. Offline, the filing request waits on the phone and sends later.

Redeploy the existing web app (do not create a new deployment) after replacing the script. Scopes are the ones photo upload already uses: `https://www.googleapis.com/auth/spreadsheets` and `https://www.googleapis.com/auth/drive`.

## What Jeff does not need to change

The page, the Settings screen, and the saved web app URL stay as they are. End Session does not delete photos still waiting on the phone. With an empty note and no photo, End Session and New Unit do not add a row.
