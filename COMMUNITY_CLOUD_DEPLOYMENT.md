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
2. Select the required processing sections; all five are selected by default.
3. Enter the project ID, first question number and first snippet ID.
4. Upload the quality-checked DOCX and select **Process document**.
5. Review any manual-check findings and download the processed DOCX, images ZIP, mapping CSV and JSON audit report.
6. Upload the images to CMS manually and pass the processed DOCX to Anand's population tool.
7. After Anand's tool creates the snippets, expand **Apply mappings to CMS after Anand’s upload**, upload the mapping CSV and apply it.

## Shared CMS mapping account

Do not commit the CMS password to GitHub. In the Streamlit Community Cloud app settings, open **Secrets** and add:

```toml
[heymath_cms]
base_url = "http://cms.heymath.com"
login = "shared-mapping-account"
password = "the-shared-account-password"
```

After saving the secret, reopen the app, expand **Apply mappings to CMS after Anand’s upload**, and select **Test CMS connection (read only)**. A successful test signs in and loads the Curriculum and Taxonomy trees without changing any CMS content.

The mapper validates every snippet and path before saving. It only adds missing Curriculum and Taxonomy mappings; existing mappings are retained and are never removed or replaced.

Because the current deployment is public, do not add the shared CMS credentials until access to the app is restricted to the team. Otherwise, anyone with the app URL could attempt a CMS mapping operation with the shared account.

## Storage note

Streamlit Community Cloud is used here as a processor, not as permanent document storage. Users should download all available results after each run. The app's temporary working files may be removed when Streamlit restarts or redeploys the app.

## Updating the app

Push approved changes to the repository's `main` branch. Community Cloud rebuilds the app from that branch. Test rule changes locally before pushing them.
