# Shift Lookup | Pregled Izmen

A bilingual shift schedule lookup web app built with plain HTML, CSS, and JavaScript.  
Employees can enter their name to instantly see their work shifts for the month.

---

## Live Site

🔗 https://parth06-exe.github.io/shift-app/

---

## Features

- 🌐 **Bilingual** — English and Slovenian (SI) language support
- 🌙 **Dark / Light mode** toggle
- 👋 **Time-based greeting** — Good morning / afternoon / evening / night
- 🔍 **Flexible name search** — Works with partial names, any capitalization
- 📅 **Today's shift highlighted** — Shows current day shift instantly
- 📂 **Upcoming & Old shifts** — Tab view to browse full month
- 🌙 **Overnight shift detection** — Shifts starting after 13:00 and ending before 06:00 are marked as overnight
- 📱 **Mobile friendly** — Responsive design works on phone and desktop
- ⚡ **No login required** — Just open and search

---

## How It Works

1. User opens the website
2. A startup popup asks to choose language — **EN** or **SI**
3. User types their name in the search bar
4. The app searches the Excel schedule file and displays:
   - Today's shift status
   - Upcoming shifts
   - Past shifts

---

## Tech Stack

| Technology | Purpose |
|---|---|
| **HTML5** | Page structure and layout |
| **CSS3** | Styling, dark/light theme, responsive design |
| **JavaScript (Vanilla)** | Search logic, Excel parsing, language switching |
| **SheetJS (xlsx.js v0.18.5)** | Reading and parsing the `.xlsx` Excel file |
| **GitHub Pages** | Free static site hosting |
| **GitHub Repository** | Source code and Excel file storage |

---

## File Structure

```
shift-app/
│
├── index.html        # Main (and only) HTML file — contains all HTML, CSS and JS
├── shift.xlsx        # Monthly shift schedule (Excel file)
└── README.md         # This file
```

---

## Excel File Format

The `shift.xlsx` file must follow this structure:

| Row | Content |
|---|---|
| Row 1 | Month and Year (e.g. `Maj 2026`) |
| Row 2 | Separator row with `---` |
| Row 3 | Header row — first column says `Ime`, rest are day numbers (1, 2, 3...) |
| Row 4+ | Each person's name and their shifts |

**Shift format in cells:**
- `12 23` → Shift from 12:00 to 23:00
- `X X` → Day off
- `18 1` → Overnight shift (18:00 to 01:00 next day)

**Overnight rule:**  
If shift start is after 13 and end is 6 or below, it is automatically marked as overnight.

---

## Updating the Schedule

Each month, simply:
1. Prepare the new `.xlsx` file in the same format
2. Go to the GitHub repository
3. Delete the old `shift.xlsx`
4. Upload the new `shift.xlsx`
5. The live site updates automatically — no code changes needed

---

## Language Support

| Text | English | Slovenian |
|---|---|---|
| Greeting (morning) | Good morning | Dobro jutro |
| Greeting (evening) | Good evening | Dober večer |
| Search button | Search | Išči |
| Today tab | Today | Danes |
| Upcoming tab | Upcoming | Prihajajoče |
| Old shifts tab | Old Shifts | Stare izmene |
| Day off | Day off | Prost dan |
| Overnight | ends next day | konča naslednji dan |

---

## Browser Support

Works on all modern browsers:
- Chrome / Edge
- Firefox
- Safari (iOS and macOS)
- Android browsers

---

## Built By

Developed for internal shift management.  
Hosted free on GitHub Pages.  
No backend, no database, no login — fully static.
