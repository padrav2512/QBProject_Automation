# Deploy on Streamlit Community Cloud

## Current deployment

- Live app: <https://qbproject-automation.streamlit.app>
- GitHub repository: <https://github.com/padrav2512/QBProject_Automation>
- Branch: `main`
- Entry point: `app.py`

The current Community Cloud deployment is public. Anyone who has the URL can open the processor. Add application authentication before using it for documents that require restricted access.

## One-time deployment

1. Create a private GitHub repository for this application.
2. Upload the contents of this folder to the repository root. `app.py` must remain at the root.
3. Sign in at <https://share.streamlit.io> with the GitHub account that can access the repository.
4. Select **Create app** and choose the repository, the `main` branch and `app.py` as the entry point.
5. Choose a short custom URL such as `cms-docx-preparer.streamlit.app`, if it is available.
6. Deploy the app.
7. In the app sharing settings, keep the app private and invite the other team members as viewers.

The app URL remains fixed. Updating files in the GitHub repository updates the deployed app without changing its URL.

## Team workflow

1. Open the shared app URL in a browser.
2. Enter the project ID, first question number and first snippet ID.
3. Upload the quality-checked DOCX.
4. Select **Process document**.
5. Review any manual-check findings.
6. Download the processed DOCX, images ZIP and JSON audit report before leaving the page.
7. Upload the images to CMS manually and pass the processed DOCX to Anand's population tool.

## Storage note

Streamlit Community Cloud is used here as a processor, not as permanent document storage. Users should download the three results after each run. The app's temporary working files may be removed when Streamlit restarts or redeploys the app.

## Updating the app

Push approved changes to the repository's `main` branch. Community Cloud rebuilds the app from that branch. Test rule changes locally before pushing them.
