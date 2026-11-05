# Service Training Roadmap Tool

A web tool that helps users plan their traiing roadmap. uUsers can browse programmes, follow a step-by-step guide to the competencies and proficiency levels that fit their role, pick a ready-made pathway, read worked case studies, and build their own training pathway to save as a PDF.

Developed by **Sufiana Safiena**, as an initiative of the Service and Operations Management (SOM) Team, Institute of Public Administration and Management (IPAM), Civil Service College.

> **Please note: the data in this repository is mock data.** The programmes, codes, links, Service Functions, competencies, steps, and case studies in `dataset.xlsx` are invented for demonstration. They do not describe real programmes or a real competency framework. You can replace them with your own data, as long as you keep the same structure (see [Using your own data](#using-your-own-data)).

---

## Why this tool exists

Training roadmaps are often shared as long documents or spreadsheets. Users have to work out for themselves which programmes fit their work, their level, and the skills they want to build. This tool turns that roadmap into a simple website, so users can:

- **find programmes quickly**, by searching, filtering, or browsing by competency
- **understand what they need**, by following the path from their Service Function, to the competencies it needs, to their proficiency level
- **learn from examples**, through case studies of users at different levels
- **start from a ready-made list**, using preset pathways by role band or by skill
- **plan and share**, by building a personal pathway and exporting it as a PDF to discuss with their reporting officer

The content lives in one Excel file, so programme owners can keep it up to date without editing any code.

---

## What the tool does

| Page | What users can do |
|---|---|
| **Home** | Browse every active programme. Search by name, filter by level, role band, modality, and service value, filter by competency, and sort. Open any programme for its overview, learning outcomes, outline, and link. |
| **How to use** | Follow the steps. Tick the Service Functions that match their work, see the competencies those functions need, find their proficiency level, and jump to matching programmes. |
| **Case studies** | Read three worked examples, each with a short story, the answers to the guide's questions, and the matching programmes. |
| **Preset pathways** | Pick a ready-made list of programmes by role band (for example, middle managers) or by skill (for example, written communication). |
| **Competency and level matrix** | See every programme against every competency at a glance. This view is intended for HR and programme owners. |
| **My pathway** | Collect programmes, add their name and organisation, and export a numbered PDF with a notice, the date, and a copyright line. |

Nothing an officer types is stored or sent anywhere. The name and organisation are only used for the PDF.

---

## Files

| File or folder | What it is |
|---|---|
| `index.html` | The page itself |
| `app.js` | The page's code |
| `styles.css` | The page's look |
| `dataset.xlsx` | **All the content.** This is the only file you need to edit. |
| `vendor/` | Local copies of two open-source libraries: SheetJS reads the Excel file, and jsPDF makes the PDF |
| `web.config` | Settings for Windows (IIS) servers only. Other servers ignore it. |
| `README.md` | This file |

This is a **static website**. It has no database, no login, and no server code, so it can be hosted on almost any web server.

---

## Running it

### On your own computer

The page has to be opened through a web address, not by double-clicking `index.html`, because browsers block pages opened from a folder from reading files. In the folder, run:

```
python -m http.server 8000
```

Then go to `http://localhost:8000` in your browser.

### On a web server

Upload every file and folder, keeping the same structure, to any web server or static hosting service. Visitors then open the folder's address, and `index.html` loads.

- The server must be able to serve `.xlsx` files. On Windows (IIS) servers, the included `web.config` sets this up.
- The server must allow the site's own scripts to run (`app.js` and the files in `vendor/`).
- The fonts come from Google Fonts. If they cannot be reached, the site uses the computer's standard font and still works.

---

## Using your own data

To use real data, replace `dataset.xlsx` with your own file. **Keep the same file name, sheet names, and column headings**, because the tool finds the content by those names. Column order does not matter, and you can add extra columns or sheets, which the tool will ignore.

Every time someone opens or reloads the site, it reads `dataset.xlsx` again, so changes show up straight away.

### The sheets

| Sheet | What it controls |
|---|---|
| `Programme Index` | One row per programme. This drives the catalogue, presets, matrix, case studies, and programme details. |
| `Guide - Steps` | The steps and tip at the top of How to use. |
| `Guide - Service Functions` | The Service Function cards and their Areas of Work. You can replace this to other functions. |
| `Guide - Competency Map` | Which competencies are critical or relevant for each function. |
| `Guide - Proficiency Levels` | The level names, designations, work areas, and the role band preset names and descriptions. |
| `Guide - Case Studies` | The case studies. |

### `Programme Index` columns

| Column | What to enter |
|---|---|
| Code | A unique code for the programme. |
| Programme | The programme title. |
| Modality | For example `Classroom`, `Virtual Classroom`, `Blended`, `Paid E-Learning`, or `Milestone`. |
| Duration | Free text, for example `1 day`. |
| Status | `Active` or `Revamping` to show it. `Discontinued` or `In-house` to hide it. |
| Competency Levels Mapped | For example `Level 2: Intermediate, Level 3: Proficient`. Leave it empty to hide a programme that is not ready yet. |
| Service Values Mapped | Comma-separated service values. |
| Functional Competencies Mapped | Comma-separated competency names. |
| Level per Competency | The level for each competency, for example `Experience Design: 3; Strategy and Policy Development: 2`. |
| Critical Competencies Mapped | Shown in the programme details. |
| Relevant Competencies Mapped | Shown in the programme details. |
| Recommended For | Comma-separated role bands, matching the designations in `Guide - Proficiency Levels`. |
| Remarks | For example `New` or `By nomination only`. Shown when there is no link. |
| Listings | The link to the programme page. |
| Description | The overview. A blank line starts a new paragraph. |
| Learning Outcomes | One outcome per line. |
| Outline | A blank line starts a new block. Lines starting with `-` become bullet points. |
| Skills | Comma-separated skill names. Each skill becomes a preset pathway. |

### Rules to keep the tool working

- **Spell names the same way everywhere.** A competency, Service Function, or role band must be written the same in every sheet and column where it appears.
- **Use semicolons in the Guide - Competency Map and case study columns**, because some names contain commas.
- **When you add a competency to a programme**, add it to both `Functional Competencies Mapped` and `Level per Competency`.
- **New names appear on their own.** A new competency, skill, modality, or role band shows on the site without any code change, with a default colour and description.
- **Do not rename the file, the sheets, or the column headings**, and keep saving the file as `.xlsx`. If you need to rename something, the code in `app.js` must be updated to match.

---

## Troubleshooting

| What you see | What to check |
|---|---|
| The page stays on "Loading the programme catalogue…" | Check that every file was uploaded, then press Ctrl and F5. If a box appears after about 15 seconds, the server is blocking `app.js`. Ask your ICT team to allow the site's own scripts. |
| "Could not load the spreadsheet reader" | The `vendor/` folder is missing, and the backup online sources are blocked. Upload the `vendor/` folder. |
| "Could not load dataset.xlsx" | Check that `dataset.xlsx` is beside `index.html`, with exactly that name. |
| "Opened straight from a folder" | Open the page through a web address, as described in [Running it](#running-it). |

To see the exact cause of a problem, press F12, open the Console tab, and reload the page.

---

## Third-party software

- [SheetJS Community Edition](https://sheetjs.com) 0.18.5, Apache 2.0 licence, in `vendor/`
- [jsPDF](https://github.com/parallax/jsPDF) 3.0.2, MIT licence, in `vendor/`
- [IBM Plex Sans and IBM Plex Mono](https://www.ibm.com/plex/), SIL Open Font Licence 1.1, from Google Fonts

---

## Contact

For questions or bug reports, contact Sufiana Safiena at sufiana_safiena@cscollege.gov.sg.

© 2026 Civil Service College. All Rights Reserved.
