# PM CV Formatter

Formats candidate CVs into the Patrick Morgan house style in seconds.

---

## Getting started

**Step 1 — Install Python (first time only)**  
If you see an error about Python when you launch, go to:  
https://www.python.org/downloads/  
Download and install Python 3.10 or later.  
**Important:** tick "Add Python to PATH" during installation, then restart your computer.

**Step 2 — Launch the app**  
Double-click **start.bat** in this folder.  
A browser window will open automatically. Dependencies install themselves on the first run.

A shortcut called **PM CV Formatter** will also appear on your Desktop — use that from now on.

**Step 3 — Format a CV**  
1. Drag and drop a candidate CV (PDF or Word) onto the upload area.
2. Optionally paste any call notes in the text box to enrich the summary.
3. Click **Format CV** and download the finished Word document.

---

## Notes

- Leave the black terminal window open while you are using the app — closing it stops the server.
- Press **Ctrl + C** in the terminal to stop the server when you are done.
- Formatted CVs are saved in the `outputs` folder inside this directory.
- The API key is already configured — no setup needed.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| "Python not found" error | Install Python 3.10+ and tick "Add Python to PATH" |
| Browser doesn't open | Go to http://localhost:5000 manually |
| "Dependency installation failed" | Right-click start.bat → Run as administrator |
| App won't start after Windows update | Re-run start.bat — it will reinstall dependencies |
