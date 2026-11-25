# Kontrollera texter (Gemini)

Data → Kontrollera texter in the master schedule asks Gemini for clear language errors in Kalender's
column M (rows dated today or later) and steps through the suggestions in a window. Code:
`apps/master-schedule/src/textCheck.js` and `textCheckDialog.html`.

## Setup

1. Sign in as a sesport.tv account with access to the Cloud project `sesport-master-schedule`, which
   needs billing: the paid tier's texts are not used to improve Google's products, and Google's terms
   require it for users in the EEA.
2. In [Google AI Studio → API keys](https://aistudio.google.com/apikey), import that project if it is not
   listed, and create a key in it. The key must be allowed to use the Gemini API only (Generative
   Language API).
3. In the Apps Script editor: Project Settings → Script Properties → add `GEMINI_API_KEY` with the key.
   It must never be in the code, the repository or the sheet.
4. Anyone who can edit the spreadsheet can read the key there. So set a low daily request quota for the
   Generative Language API in the Cloud console (a run takes one request per 25 new texts), and a
   budget alert under Billing. Delete the key in AI Studio if it leaks.
5. Run `runPostDeployTests`.

Optional: `GEMINI_MODEL`, if the default (`TEXT_CHECK_MODEL` in textCheck.js) is not good enough.

## What is sent and remembered

- Sent: the texts in M that are new or changed, with genre, league and programme title. Identical
  texts are sent once. Gemini returns only the texts it corrects.
- A suggestion that changes more than about 30 % of the text is dropped (it is a rewrite).
- Texts found correct, kept with Behåll or written with Uppdatera are remembered as fingerprints in the
  document's properties, and not checked again until they change. With another model, the texts found
  correct are checked again.
- The yellow cell gets its own colour back when the window moves on. If the window is closed with X,
  `checkProgramLengths` puts it back within a few minutes.
