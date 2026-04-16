# Cutlist Optimizer

Web-based planning tool for furniture projects imported from Fusion 360 OBJ files.

## Units + Orientation

- **Design import units:** `mm`, `cm`, `m`, `in`, or `ft` (+ optional scale multiplier).
- **Part outputs:** metric (`mm`).
- **Stock buying inputs:** imperial (`quarters`, width in inches, length in feet).

Part orientation is fixed as requested:
- **X axis** = longest part dimension (grain assumed along X).
- **Y axis** = middle dimension.
- **Z axis** = shortest dimension (thickness axis).

## Raw, Net, Rough Metrics

- **Raw**: dimensions from the imported OBJ bounding box before allowances.
- **Net**: finished target dimensions after orientation mapping (X longest, Z shortest).
- **Rough**: cut dimensions used for optimization, calculated as:
  - `rough length = net length + length allowance`
  - `rough width = net width + width allowance`
  - `rough thickness = net thickness + thickness allowance`

## Additional Scale Multiplier

- Applied after unit conversion.
- Final scaling formula:
  - `scaled_mm = source_value * unit_factor_to_mm * additional_scale_multiplier`
- Use `1` for normal exports.
- Use another value only if your source export is consistently scaled wrong.

## Planning vs Lumber Yard

These two workflows share milling allowances, kerf, and pricing but use different board sources:

| | Plan Stock | Lumber Yard Recalculate |
|---|---|---|
| **Board source** | Auto-generated catalog within your width/length range | Boards you enter in the Inventory table |
| **Purpose** | Budget estimate before visiting the yard | Exact cut plan for boards you've found |
| **Catalog range** | Controlled by width/length min–max on Planning tab | Not used |
| **Milling allowances** | Shared — configured on Planning tab | Shared — same values |
| **Max planer width** | Biases toward narrower stock | Same bias + Workshop warnings |
| **Quantities** | Unlimited (hypothetical) | Respects inventory quantities |

**Settings that apply to Planning only** (Planning Stock Catalog section):
- Width min/max and Length min/max — define the hypothetical board range for Plan Stock.

**Settings on the Lumber Yard tab:**
- **Max planer width (in)** — your planer's physical capacity. Boards wider than this are flagged in the Workshop guide to be ripped into strips before planing. Both Planning and Lumber Yard prefer boards within this width. Set to `0` to disable.

**Shared settings (configured once on the Planning tab, apply to both):**
- Milling allowances (thickness, width, length, board-end trim, rip margin)
- Saw kerf
- Price per board foot

## Features

1. **Part extraction + planning**
   - Parses OBJ groups/objects into individual parts.
   - Computes per-part dimensions.
   - Applies editable milling allowances (project-specific defaults):
     - thickness: `3.2 mm`
     - width: `3.2 mm`
     - length: `25.4 mm`
     - board-end trim reserve: `50.8 mm`
     - rip margin per cut: `1.6 mm`
   - **Max planer width (in):** boards wider than this are flagged in the workshop guide with rip-to-strips instructions. The planner and lumber yard recalculation bias board selection toward widths ≤ this limit; wider boards are still used when no narrow board fits a blank. Set to `0` to disable.
   - **Curved part detection:** parts whose vertex cloud contains more than ~15 % "interior" vertices (not near any bounding-box face) are flagged with a ⌒ curved badge in the parts table. The bounding box is still the correct blank size, but the badge reminds you that shaping will be required after rough milling.
   - Assigns stock thickness in **quarters**, with lamination layers when needed.

2. **Thickness + grain controls**
   - Default grain direction per project (`Long` lock or `Free`).
   - Per-part 3-way grain direction: **Long** (grain along longest axis, no rotation), **Mid** (grain along middle axis — blank pre-rotated before nesting), **Free** (nesting may rotate freely).
   - Per-part thickness override (`Auto` or specific quarter).
   - Global override apply-to-all.

3. **Enhanced nesting solver**
   - MaxRects-style nesting heuristic.
   - Respects grain direction: `Long` and `Mid` both lock nesting rotation; `Mid` pre-rotates the blank before placement.
   - Packs rough blanks, including lamination layer expansion.
   - Board diagrams displayed horizontally — parts labelled left-to-right.
   - All boards in a result set share a proportional scale (widest board = 100% height).

4. **Workshop cut guide**
   - Activated automatically after Plan Stock or Recalculate.
   - Prefers Lumber Yard Recalculate result (real boards); falls back to Plan Stock.
   - Per-board cards showing: board description, parts table (rough + net dims + grain), recommended cut sequence, final milling reference.
   - **🖨 Print / Save PDF** button exports the guide as a PDF — one page per board plus a final consolidated schedule page. Each page footer shows the project name and print date/time.
   - **Cut sequence** covers the full standard milling workflow per board:
     - Inspect for defects / warp
     - Face joint (jointer) → plane to rough thickness (planer)
     - Re-saw suggestion when stock is significantly thicker than needed
     - Joint one edge (jointer) → trim ends (miter saw)
     - Cross-cut sections (miter saw) → rip blanks to width (table saw), ordered left-to-right
     - Lamination note when layers must be glued up
   - **Final milling reference** table shows net target dimensions (T/W/L) per part.
   - **Panel glue-up**: parts whose rough width exceeds the widest available board are automatically split into the minimum number of strips needed. Each strip gets `3.2 mm` of extra width for edge-jointing the glue faces. Strip blanks are nested onto boards like any other blank; the workshop guide shows which strips are on which board and provides edge-joint, dry-fit, and glue-up instructions.
   - **Consolidated Mill Schedule** card (bottom of the tab and last PDF page) sequences all operations across every board to minimise planer height changes, table saw fence moves, and miter saw stop adjustments. Phases: Inspect → Face joint → (Re-saw) → Plane by thickness group (thickest first) → Joint edge → Trim ends → Cross-cut by length (longest first) → Rip by width (widest first) → (Panel glue-up) → (Thickness lamination glue-up).

5. **Lumber-yard recalculation**
   - Re-optimizes against real inventory rows.
   - Optional infinite quantity mode (inventory rows define sizes only).
   - Reports unmet parts and additional stock suggestions.
   - Boards exceeding the planning **Length max** are flagged for yard pre-cuts.
   - **Recalculate** button activates as soon as parts are analyzed and inventory is configured (infinite mode or at least one row) — does not require Plan Stock to be run first.

5. **Costing**
   - Total board feet calculation.
   - Project-level price-per-board-foot.
   - Estimated lumber cost totals and per stock size line.

6. **3D model preview**
  - Touch-friendly orbit/pan/zoom + reset view.
  - Optimized layout for iPhone-sized screens.
   - Refreshes after model load and analyze.

7. **Planner UI behavior**
   - Planner / Lumber Yard / Admin / Instructions tab layout.
   - Top-right sync icon (red = disconnected, green = connected to Firebase).
   - Parts table sorted alphabetically by default.
   - Click any displayed parts column header to sort.
   - Raw dimensions are hidden from table columns; hover the part name to see raw X/Y/Z.

## Storage Backend

Projects are stored in **Firebase Cloud Storage** (Firestore). Users must sign in via the avatar button to save and load projects.

Saved projects include the **final inventory result** (Lumber Yard Recalculate output) so the Workshop tab restores correctly on load. `freeRects` data is stripped before saving to keep document size small.

## Firebase Setup Steps

1. Create a Firebase project in the Firebase console.
2. Add a **Web App** to the project and copy:
   - `apiKey`
   - `authDomain`
   - `projectId`
   - `appId`
   - optional: `storageBucket`, `messagingSenderId`
3. Paste these values into `DEFAULT_FIREBASE_CONFIG` at the top of `app.js`.
4. Enable **Firestore Database** in **Native mode**.
5. Enable **Authentication → Email/Password** sign-in method.
6. Add your host to **Authentication → Settings → Authorized domains**:
   - include `localhost` (for local testing)
   - include your production domain when deployed
7. Paste the Firestore Security Rules below into your project.
8. Use the avatar menu in the app to create the first user account, then promote it to Admin via the Firestore Console.

### Recommended Firestore Security Rules

These rules allow admins to read all projects (needed for the Admin tab's Users & Projects panel).

```txt
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {

    function isSignedIn() {
      return request.auth != null;
    }

    function isOwner(uid) {
      return isSignedIn() && request.auth.uid == uid;
    }

    function userDoc() {
      return get(/databases/$(database)/documents/users/$(request.auth.uid)).data;
    }

    function isAdmin() {
      return isSignedIn() && userDoc().role == 'admin';
    }

    // User profile documents — role field is write-protected from clients
    match /users/{uid} {
      allow read:   if isOwner(uid) || isAdmin();
      allow create: if isOwner(uid)
                    && request.resource.data.role == 'standard'
                    && request.resource.data.email == request.auth.token.email;
      allow update: if isAdmin();
      allow delete: if isAdmin();
    }

    // Projects — owned by creating user; admins can read all
    match /projects/{projectId} {
      allow read:           if isAdmin()
                            || (isSignedIn() && request.auth.uid == resource.data.ownerUid);
      allow update, delete: if isSignedIn() && request.auth.uid == resource.data.ownerUid;
      allow create:         if isSignedIn() && request.auth.uid == request.resource.data.ownerUid;
    }
  }
}
```

Notes:
- The app writes `ownerUid` on each project document.
- The `create` rule for `/users/{uid}` prevents any client from self-assigning the `admin` role.
- If you already have data without `ownerUid`, add it before enforcing strict rules.

## Firebase App Check

App Check prevents unauthorized clients from using your Firebase backend.

### Web setup (reCAPTCHA v3)

1. Firebase Console → **App Check** → select your web app → choose **reCAPTCHA v3**.
2. [Google reCAPTCHA Admin](https://www.google.com/recaptcha/admin) → **+ Create** → type **Score based (v3)**.
   - Add your domains: `yourusername.github.io` and `localhost`.
   - Copy the **Site Key**.
3. In `app.js`, paste the site key into `RECAPTCHA_SITE_KEY`:
   ```js
   const RECAPTCHA_SITE_KEY = "6Lc...your-key-here";
   ```
4. Firebase Console → App Check → your app → **overflow menu → Monitor**.
   Watch traffic for ~1 week, then switch to **Enforce** once legitimate traffic is confirmed.

### Cordova (future)
Replace `ReCaptchaV3Provider` with the native Play Integrity (Android) or App Attest (iOS) provider in the Cordova build. Leave `RECAPTCHA_SITE_KEY` empty in native builds.

### Development / local testing
Leave `RECAPTCHA_SITE_KEY = ""` to skip App Check entirely during development. To test with App Check locally, add a debug token: Firebase Console → App Check → your app → **Add debug token**, then set `self.FIREBASE_APPCHECK_DEBUG_TOKEN = "your-debug-token"` before the Firebase scripts load in `index.html`.

## User Accounts + Roles

The app supports two user types:

| Role | Created via | Access |
|------|-------------|--------|
| **Standard** | Web app sign-up | All tabs except Admin |
| **Admin** | Firebase Console (manual promotion) | All tabs including Admin |

### Signing in

Click the **avatar button** (upper-right corner) to open the account menu. From there you can:
- **Sign In** — log in with an existing email/password account
- **Create Account** — register a new Standard account
- **Forgot password?** — receive a password reset email
- **Sign Out** — log out

### Creating an Admin account

All accounts created through the web app are Standard users. To promote a user to Admin:

1. Sign in to the **Firebase Console** → your project → **Firestore Database**.
2. Browse to the `users` collection.
3. Find the document whose `email` matches the user you want to promote.
   - If no document exists yet, have that user sign in once to trigger document creation.
4. Click the document → click the `role` field → change `"standard"` to `"admin"` → **Update**.
5. The change takes effect the next time that user signs in (or on their next page load).

> **Security note:** The Firestore rules prevent any client from writing their own `role` field.
> Only the Firebase Console (service-account level) can promote users to Admin.

## Hosting on GitHub Pages

The app has no build step, so GitHub Pages serves it directly from the repository.

### Steps

1. Push the repository to GitHub (if not already there):
   ```bash
   git remote add origin https://github.com/<your-username>/<repo-name>.git
   git push -u origin main
   ```
2. In your GitHub repository, go to **Settings → Pages**.
3. Under **Source**, select **Deploy from a branch**.
4. Choose **main** branch and **/ (root)** folder, then click **Save**.
5. GitHub will publish the site at `https://<your-username>.github.io/<repo-name>/`.
   The URL appears at the top of the Pages settings once deployment completes (usually ~1 minute).

### Firebase authorized domain

Once deployed, add your GitHub Pages domain to Firebase:

1. Firebase console → **Authentication → Settings → Authorized domains**.
2. Click **Add domain** and enter `<your-username>.github.io`.

The app will then be able to sign in to Firebase from that domain.

### Updating the site

Every push to `main` automatically re-deploys. There is no separate build or publish step.

## Run

No build step required:

1. Open [index.html](/Users/Steve2/GitHub/CutlistOptimizer/index.html) in a browser.
2. Load an OBJ file.
3. Run:
   - `1) Analyze Model`
   - `2) Planning Stock`
   - `3) Lumber Yard Recalculate`

## Files

- [index.html](/Users/Steve2/GitHub/CutlistOptimizer/index.html): UI + CDN scripts (Three.js + Firebase)
- [styles.css](/Users/Steve2/GitHub/CutlistOptimizer/styles.css): responsive styling
- [app.js](/Users/Steve2/GitHub/CutlistOptimizer/app.js): parser, orientation rules, nesting, pricing, viewer, storage backends
