# AdvocateDocs-POC

POC: migrate an advocate firm's document features from Dynamics 365 to M365 (SharePoint Online + SPFx).

## Solutions
| Folder | What |
|---|---|
| `advocate-docs/` | SPFx 1.23.2 ListView Command Set `DocActions` on document libraries (template 101) |
| `office-addin/` | Office add-in (Word, Excel, PowerPoint): Global Advocate ribbon tab, React task pane, NAA sign-in |

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
Manifest: `manifest.xml` (add-in only, Word + Excel + PowerPoint). `manifest.json` (unified) is not deployed.

Dev loop (no catalog upload; add-in is sideloaded per session through the document URL):
| Step | Command |
|---|---|
| Preflight (dev server, manifest, Entra redirect URIs + consent, session, test docs) | `./scripts/Test-AddinDevReady.ps1` |
| Static checks | `npm run build`, `npm run validate`, `npm run lint` |
| E2E in Word/Excel/PowerPoint on the web (Edge, headless) | `npm run e2e` (`e2e:headed` to watch) |
| One-time / when session expires: sign in as dev06 | `npm run e2e:login` |

- Test documents: `e2e/.env` (copy `e2e/.env.example`), ADVTest files in `ADVTestMatter01/ADVTestDocuments`.
- Run output (pane console, screenshots, traces): `e2e/out/`.

Production:
| Change | Command | Live |
|---|---|---|
| Code (pane, logic) | `./scripts/Deploy-AdvAddinWeb.ps1 -Execute` (build + Azure Storage static website) | immediately |
| Manifest (ribbon, permissions, version) | `./scripts/Publish-AdvAddinManifest.ps1 [-ProductId <id>] [-Members <upn>] -Execute` (tenant admin) | up to 24 h |

Entra app (`ADVTest-DocsApi`) needs SPA redirect URIs for each host origin: `https://<origin>/taskpane.html` and `brk-multihub://<origin host>`.
