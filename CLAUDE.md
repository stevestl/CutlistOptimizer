# Cutlist Optimizer — Claude Instructions

## Project Overview
A vanilla HTML/CSS/JS web app (no build step) for woodworking cutlist planning. Imports Fusion 360 OBJ files, extracts part dimensions, runs a MaxRects-style nesting algorithm, and generates cutting layouts with cost estimates. Uses Firebase for cloud storage and Three.js for 3D preview.

Key files:
- `index.html` — UI layout and controls
- `app.js` — All application logic (~2500 lines)
- `styles.css` — Responsive styles
- `README.md` — User-facing documentation

## Instructions

### Documentation
- Always update `README.md` when making feature changes, adding options, or changing behavior visible to users.
- Keep inline comments current when modifying complex logic (nesting algorithm, OBJ parser, dimension calculations).

### Code Style
- Vanilla JS only — no frameworks, no npm packages, no build tools.
- ES6+ syntax (arrow functions, template literals, destructuring, async/await).
- External libraries via CDN only (Three.js, Firebase).
- Preserve the flat file structure — do not introduce subdirectories or module bundling.

### Testing
- No automated test suite exists. After changes, manually verify by opening `index.html` directly in a browser.
- Test the full workflow: import an OBJ → Analyze → Plan → Lumber Yard Recalculate.
- Check both localStorage and Firebase storage paths when touching storage code.

### Safety
- Do not commit Firebase config credentials. The Firebase config in `app.js` uses placeholder values — keep it that way and note in README that users must supply their own config.
- Avoid breaking the mobile-responsive layout when editing CSS or HTML.
