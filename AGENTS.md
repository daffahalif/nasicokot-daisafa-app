# AGENTS.md

## Rules (must follow)

### Responses
- Keep answers short and simple, unless I ask for more.
- All text users see in the app must be in Indonesian.
- Money is in Rupiah (Rp).
- Use English for response. But for real case, variables, or data, use the real data

### Planning
- Always ask questions first. Never assume the design, tools, or features.
- Use sub-agents to research and to double-check the plan before showing it to me.

### Making changes
- Don't do the work yourself when possible. Use sub-agents.
- Split the plan into parts that can run in parallel, and act only as the coordinator.
- Use a strong model for hard tasks (like coding) and a cheaper one for easy tasks (like docs).
- After every change, check your work. This project has no automatic checks, so manually confirm that the front end and back end still match up.
- Always update code changes to CHANGELOG.md efficiently. do not make it too wordy.

### Data changes
- If you change the sheet columns, update everything that uses them (the column maps, the setup headers, and all code that reads or writes them).
- NEVER run `setupSheets()` on the real spreadsheet. It deletes all data.

### Testing
- Never assume a change works.
- This project has no test tools, so ask me whether to skip testing or whether I'll test it myself in the Apps Script editor.