# AdvocateDocs-POC

POC: migrate an advocate firm's document features from Dynamics 365 to M365 (SharePoint Online + SPFx).

## Solutions
| Folder | What |
|---|---|
| `advocate-docs/` | SPFx 1.23.2 ListView Command Set `DocActions` on document libraries (template 101) |
| `office-addin/` | Office add-ins on one web app: Word/Excel/PowerPoint (Global Advocate ribbon tab) and Outlook (File to Matter, Attach from Matter); React task pane, NAA sign-in |

## Commands
| Command | Visible when | Status |
|---|---|---|
| Preview | 1 doc selected | scaffold |
| Merge to PDF (with index page) | 2+ docs selected | scaffold |
| Email as PDF | 1 doc selected | scaffold |
| Email specific version | 1 doc selected | scaffold |

Native SharePoint covers Open, Properties, Delete, Document List, versioning.

## Build (Node 22.14+ < 23)
```powershell
cd advocate-docs
npm install
npm run build   # heft test + package-solution -> sharepoint/solution/advocate-docs.sppkg
npm run start   # heft serve, opens serve.json pageUrl
```

## Graph permissions (admin approval after deploy)
`Files.ReadWrite.All`, `Mail.ReadWrite` (delegated) — SharePoint admin center > Advanced > API access.

Dev tenant: `ccsdev01`, site `/sites/ADVTestDocs`, test marker `ADVTest`.

## Office add-in (`office-addin/`)
Manifests:
- `manifest.xml` (add-in only, Word + Excel + PowerPoint): production. The build gives it the prod Id and name; users get the "Global Advocate" tab through Centralized Deployment.
- `manifest.outlook.xml` (add-in only, `MailApp`): the **Outlook add-in**, a separate registration because an XML manifest cannot be both TaskPaneApp and MailApp. Same web app (`taskpane.html?view=file|attach`), same Entra app. As is, it is the DEV copy "Global Advocate DEV" (localhost); the build gives `dist/manifest.outlook.xml` the prod Id and name. Its DEV counterpart is not this file: the Outlook commands are in the `manifest.json` DEV copy (see next line), because a separately sideloaded Outlook add-in replaced the Teams DEV app for that user.
- `manifest.json` (unified): **DEV copy** "Global Advocate DEV", loading from https://localhost:3000. It covers Word, Excel, PowerPoint **and Outlook** (read: File to Matter, compose: Attach from Matter). It's installed only for the E2E account (dev06) as a personal custom app: `npm run package:dev`, then in Teams on the web go to Store › ⚙ › Upload an app › Upload a custom app › `release/advocate-addin-dev.zip` › Add. Bump `version` to update it. Never upload it for real users (it duplicates the tab).

Dev loop (E2E account dev06 sees both tabs: the DEV copy serves the dev server, prod serves hosting). URL-parameter web sideloading does not load the add-in; don't rely on it:
| Step | Command |
|---|---|
| Preflight (dev server, manifest, Entra redirect URIs + consent, session, test docs) | `./scripts/Test-AddinDevReady.ps1` |
| Static checks | `npm run build`, `npm run validate`, `npm run lint` |
| E2E in Word/Excel/PowerPoint and Outlook on the web (Edge, headless) | `npm run e2e` (`e2e:headed` to watch) |
| One-time / when session expires: sign in as dev06 | `npm run e2e:login` |
| Production only (dev server off, hosted add-in) | `ADV_E2E_NO_DEV_SERVER=1 ADV_E2E_HOSTED_ORIGIN=https://<hosting> npm run e2e` |

- Test documents: `e2e/.env` (copy `e2e/.env.example`), ADVTest files in `ADVTestMatter01/ADVTestDocuments`.
- Run output (pane console, screenshots, traces): `e2e/out/`.

Production:
| Change | Command | Live |
|---|---|---|
| Code (pane, logic) | `./scripts/Deploy-AdvAddinWeb.ps1 -Execute` (build + Azure Storage static website) | immediately |
| Manifest (ribbon, permissions, version) | `./scripts/Publish-AdvAddinManifest.ps1 [-ProductId <id>] [-Members <upn>] -Execute` (tenant admin) | up to 24 h |
| Outlook manifest | same script with `-ManifestPath dist/manifest.outlook.xml` (own ProductId) | up to 24 h |

Entra app (`ADVTest-DocsApi`) needs SPA redirect URIs for each host origin: `https://<origin>/taskpane.html` and `brk-multihub://<origin host>`.
