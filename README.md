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
- `manifest.outlook.xml` (add-in only, `MailApp`): the **Outlook add-in**, a separate registration because an XML manifest cannot be both TaskPaneApp and MailApp. Same web app (`taskpane.html?view=file|attach`), same Entra app. The build gives `dist/manifest.outlook.xml` the prod Id and name. Its DEV counterpart is `manifest.outlook-dev.json` (next lines), not this file: a sideloaded XML Outlook add-in replaced the Teams DEV app for that user.
- `manifest.json` (unified): **DEV copy** "Global Advocate DEV" for Word, Excel and PowerPoint, loading from https://localhost:3000.
- `manifest.outlook-dev.json` (unified): **DEV copy** "Global Advocate DEV (Outlook)" (own app id). Kept separate: while one DEV app carried Outlook too, it kept vanishing for dev06 although Teams still listed it (2026-11-05..09).
- Both DEV copies are installed only for the E2E account (dev06) as personal custom apps: `npm run package:dev && npm run e2e:dev-app` (or Teams on the web › Store › ⚙ › Upload an app › Upload a custom app › the zips in `release/`). Bump `version` to update. Re-upload is also the fix when a DEV tab / Outlook tile disappears. Never upload them for real users (they duplicate the tab).

Word desktop Send as Attachment: no browser window. The result says the email is in Outlook **Drafts**, with **Open Outlook app** (`ms-outlook://`, brings Outlook to the front; new Outlook cannot open a given draft, spike 2026-11-09) and a secondary Outlook on the web link (uses the browser's signed-in account). Office on Windows asks once whether the add-in may open `ms-outlook` links; admins can pre-allow it: `HKLMSoftwarePoliciesMicrosoftOffice.0WEFProtocolHandlers<add-in id>` value `ms-outlook` = `Allow`.

Outlook visibility:
- Pinned, **File to Matter** (read) and **Attach from Matter** (compose) sit in the ribbon's Add-ins group and the message header; narrow windows fold them into Apps / More apps. Pinning is per user: Apps › right-click the tile › Pin (appears after a reload; Outlook settings no longer list add-ins). One DEV install came pinned, a later one did not, so tell users how to pin.
- The File to Matter pane can be pinned (Pin Apps Pane): it stays open and starts over for each selected email.
- New emails get the banner "Global Advocate: attach documents from a matter." with an **Attach from Matter** action (event-based activation `OnMessageCompose`, `src/launchevent/`, Mailbox 1.12), whenever an email is composed: new, reply, reply all, forward and a reopened draft. Outlook fires the event **only for admin-deployed add-ins** (Centralized Deployment), so the user-installed DEV copy shows no banner on its own; the E2E runs the handler directly. Each manifest update of an event-based add-in needs admin consent.

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
