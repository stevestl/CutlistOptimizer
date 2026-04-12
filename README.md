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
   - Assigns stock thickness in **quarters**, with lamination layers when needed.

2. **Thickness + grain controls**
   - Default grain lock per project.
   - Per-part grain lock toggle.
   - Per-part thickness override (`Auto` or specific quarter).
   - Global override apply-to-all.

3. **Enhanced nesting solver**
   - MaxRects-style nesting heuristic.
   - Respects grain lock (prevents 90° rotation when locked).
   - Packs rough blanks, including lamination layer expansion.

4. **Lumber-yard recalculation**
  - Re-optimizes against real inventory rows.
   - Optional infinite quantity mode (inventory rows define sizes only).
  - Reports unmet parts and additional stock suggestions.
  - Suggests pre-cuts for long boards when sub-6' carry cuts are possible.

5. **Costing**
   - Total board feet calculation.
   - Project-level price-per-board-foot.
   - Estimated lumber cost totals and per stock size line.

6. **3D model preview**
  - Touch-friendly orbit/pan/zoom + reset view.
  - Optimized layout for iPhone-sized screens.
   - Refreshes after model load and analyze.

7. **Planner UI behavior**
   - Planner/Settings tab layout.
   - Top-right status icons:
     - sync icon (red/green)
     - storage icon (`L` local, `C` cloud) with hover definitions
   - Parts table sorted alphabetically by default.
   - Click any displayed parts column header to sort.
   - Raw dimensions are hidden from table columns; hover the part name to see raw X/Y/Z.

## Storage Backends

The app supports both:
- **Local Browser Storage** (`localStorage`)
- **Firebase Cloud Storage** (Firestore + anonymous auth)

You can switch between backends in the **Cloud Sync (Firebase)** section.

## Firebase Setup Steps

1. Create a Firebase project in the Firebase console.
2. Add a **Web App** to the project and copy:
   - `apiKey`
   - `authDomain`
   - `projectId`
   - `appId`
   - optional: `storageBucket`, `messagingSenderId`
3. Enable **Firestore Database** in **Native mode**.
4. Enable **Authentication**:
   - Sign-in method: **Anonymous** (enable it).
5. Add your host to **Authentication > Settings > Authorized domains**:
   - include `localhost` (for local testing)
   - include your production domain when deployed
6. In this app, paste config values into **Cloud Sync (Firebase)**.
7. Click **Save Firebase Config**, then **Connect Firebase**.
8. Confirm top-right icons show cloud mode:
   - storage icon switches to `C`
   - sync icon turns green

### Recommended Firestore Security Rules (starter)

Use these as a base (adjust as needed):

```txt
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /projects/{projectId} {
      allow read: if request.auth != null
        && request.auth.uid == resource.data.ownerUid;
      allow update, delete: if request.auth != null
        && request.auth.uid == resource.data.ownerUid;
      allow create: if request.auth != null
        && request.auth.uid == request.resource.data.ownerUid;
    }
  }
}
```

Notes:
- The app writes `ownerUid` on each project document.
- If you already have data without `ownerUid`, add it before enforcing strict rules.

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
