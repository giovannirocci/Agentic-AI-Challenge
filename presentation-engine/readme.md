#  Presentation Engine (Agent prompting → JSON → PowerPoint)

##  Overview

This project is a **presentation generation engine** that converts structured JSON inputs into fully styled PowerPoint presentations.

Instead of manually creating slides, the system:

1. Takes **content** (what should be on the slides)
2. Takes **format/style profiles** (how slides should look)
3. Validates both inputs
4. Automatically generates a **brand-compliant PowerPoint (.pptx)**

---

## Core Idea

We separate presentation generation into three independent layers:

| Layer | Purpose |
|------|--------|
| **Content JSON** | Defines slide structure (titles, bullets, images, etc.) |
| **Format/Profile JSON** | Defines styling (colors, fonts, layout rules) |
| **Rendering Engine** | Combines both to generate PowerPoint slides |


---

##  Project Structure
```bash
presentation-engine/

#  Inputs
├── input/
│   ├── content/
│   │   ├── manual/        # Ground truth (human-created)
│   │   └── generated/     # AI-generated content
│   │
│   ├── formats/
│   │   ├── manual/        # Official branding profiles
│   │   └── generated/     # AI-generated styles

#  Engine
├── src/
│   ├── render_presentation.js
│   ├── validate_presentation.js
│   └── validate_profile.js

#  Prompts (for AI)
├── prompts/
│   ├── actual_agent_prompts/
│   └── instruction_prompts/

#  Outputs
├── output/

#  Misc
├── backup/
├── package.json
└── README.md
```
##  How It Works (Pipeline)

### Step-by-step:

1. **Load Inputs**
   - Presentation content (`input/content/...`)
   - Style profile (`input/formats/...`)

2. **Validate Inputs**
   - `validate_presentation.js` → checks slide structure
   - `validate_profile.js` → checks formatting rules

3. **Render Presentation**
   - `render_presentation.js`:
     - applies layout rules
     - injects content
     - applies branding
     - generates slides using `PptxGenJS`

4. **Export Output**
   - Final `.pptx` file is written to `/output`

---

##  Core Components

### `render_presentation.js`
Main engine:
- Loads JSON inputs
- Applies theme (colors, fonts)
- Builds slides dynamically
- Exports PowerPoint file

---

### `validate_presentation.js`
Validates:
- slide structure
- required fields
- content block format

---

### `validate_profile.js`
Validates:
- color tokens
- font tokens
- style schema consistency

---

## Installation


# Install dependencies
Make sure you have Node.js installed.

```bash
npm install
```
# How to run (example)

## Validate inputs
```bash
node src/validate_profile.js input/formats/manual/company_style_profile_talentia.json
node src/validate_presentation.js input/content/manual/example_deck_v2_rich.json
```

## Generate presentation
```bash
node src/render_presentation.js \
  input/content/manual/example_deck_v2_rich.json \
  output/demo.pptx \
  input/formats/manual/company_style_profile_talentia.json
```