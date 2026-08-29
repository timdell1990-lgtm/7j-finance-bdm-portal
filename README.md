# 7J Finance BDM Portal — Organised Build

## Entry point
Open `index.html`.

## Main application
`app.html` contains the application markup and external script references.

## Styles
`css/portal.css` contains the CSS extracted from the original single HTML file.

## JavaScript
`js/` contains the inline JavaScript blocks extracted in their original execution order.
The numeric prefixes preserve the original order to minimise behavioural changes.

## Repair workflow
1. Open `index.html` to run the portal.
2. For visual fixes, check `css/portal.css`.
3. For functionality, search the numbered files in `js/`.
4. Make the smallest targeted change.
5. Test the portal before replacing the production copy.

## Important
This is an organisation/refactoring pass. Existing application logic and markup have been retained rather than rewritten.

## Local runtime hardening

The packaged portal now detects `file://`, `localhost`, `127.0.0.1`, and `::1` as local runtimes. In those environments it automatically starts the built-in Demo Mode and never initialises Microsoft MSAL/Graph against the production SharePoint tenant. The production Microsoft 365 configuration remains available when the portal is hosted on its configured SharePoint origin.



## Netlify production deployment

The production build is configured for `https://tims7jbdmportal.netlify.app`.
Microsoft Entra ID must have this exact redirect URI registered:
`https://tims7jbdmportal.netlify.app/app.html`

Local `file://` and `localhost` runs remain in Demo Mode and do not load MSAL.
The existing SharePoint deployment continues to use its SharePoint redirect URI.
