# AdvocateDocs-POC

POC: migrate an advocate firm's document features from Dynamics 365 to M365 (SharePoint Online + SPFx).

## Solutions
| Folder | What |
|---|---|
| `advocate-docs/` | SPFx 1.23.2 ListView Command Set `DocActions` on document libraries (template 101) |

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
