# ITSRR Form Code Generator

A desktop department selector built with Python and Tkinter.

## Run

Python 3 is required. Install the Excel dependency once:

```powershell
python -m pip install -r requirements.txt
```

```powershell
python app.py
```

Choose a department to generate a code in this format:

```text
BRANCH-MM-SEQUENCE
```

For example, the first HR code generated in July is `HR-07-01`. The next HR
code generated that month is `HR-07-02`.

Sequence numbers are tracked separately for each department and month in
`itsrr_counters.json`, so they continue after restarting the application.
Use **Counter Settings** to choose the next number manually or reset all
department counters to `01`.
Use **Copy** to copy the generated ITSRR client template to the clipboard.

Use **Add Department** to enter a new department name and its ITSRR code.
Custom departments are saved in `custom_departments.json` and remain available
after restarting the application.

After a client completes the template, use **Compile Response**:

1. Paste the completed response.
2. Click **Add to Excel**.
3. Find the compiled record in `ITSRR Summary.xlsx`.

Existing rows with the same ITSRR number are updated instead of duplicated.

## Windows executable

The packaged application is available at `dist\ITSRR-Generator.exe`. It can be
copied to another Windows computer and run without installing Python.

Keep the executable in a writable folder. It creates `itsrr_counters.json` and
`ITSRR Summary.xlsx` beside itself as the application is used. Added departments
are stored in `custom_departments.json` in the same folder.
