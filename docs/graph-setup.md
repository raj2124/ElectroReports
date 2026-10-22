# Microsoft Graph setup (Session 6H)

One checklist for the session with the Global Administrator. It sets up the **ElectroReports Graph** app so the
ElectroReports server can

* **write** events to the Outlook calendars of one group of engineers (calendar sync), and
* **read** three SharePoint sites (knowledge base). SharePoint is read-only, always: the app never creates, renames,
  moves, uploads or deletes anything there and needs no change to your folders. Only the "read" role is granted.

Nothing here is needed to run or test the app: with `GRAPH_USE_MOCK=true` (the default) both features run on an
in-memory Graph. Do this checklist when the real thing should go live.

## Who does what

| Step | What | Who |
|---|---|---|
| 1 | Register the app | Application Developer / Cloud Application Administrator, or you if your tenant lets users register apps; otherwise the Global Admin |
| 2 | Create the certificate, upload its public part | you |
| 3 | Add `Sites.Selected`, **grant admin consent** | **Global Administrator** (or Privileged Role Administrator) |
| 4 | Exchange: group, scope, role assignment | **Exchange Administrator** (Organization Management) or the Global Admin |
| 5 | Grant "read" on each of the three sites | **Global Administrator** (the Graph permission used for the grant, `Sites.FullControl.All`, needs admin consent) |
| 6 | `.env` and the live checks | you |

Tools used: PowerShell 7 with the modules `ExchangeOnlineManagement` and `Microsoft.Graph` (`Install-Module ExchangeOnlineManagement, Microsoft.Graph -Scope CurrentUser`), and `openssl` on the machine that runs the server.

Replace `<APP_ID>` (the Application (client) ID of step 1), `<TENANT>` (your tenant's SharePoint name; the sites are on `eletechnology.sharepoint.com`) and the mailboxes with real values. Never paste a secret or a private key into a chat, an issue or git.

---

## 1. Register the app (Entra admin center)

1. https://entra.microsoft.com → **Identity → Applications → App registrations → New registration**.
2. Name: `ElectroReports Graph`. Supported account types: **Accounts in this organizational directory only (single tenant)**. Redirect URI: leave empty.
3. **Register.**
4. On the **Overview** page copy the **Application (client) ID** and the **Directory (tenant) ID**.
5. **Authentication**: leave everything as it is (no platform, no implicit grant, no public client flow).
6. **Owners**: add yourself and one more admin (so the certificate can be renewed without the Global Admin).

Expected: the app exists, single tenant, with no redirect URI and no permissions yet.

## 2. Certificate (no client secret)

On the machine that will run the server. The private key stays there, **outside the repository** (the repo is public):

```bash
mkdir -p ~/secrets && chmod 700 ~/secrets
openssl req -x509 -newkey rsa:2048 -nodes -days 730 -subj "/CN=ElectroReports Graph" \
  -keyout ~/secrets/electroreports-graph.key -out ~/secrets/electroreports-graph.crt
chmod 600 ~/secrets/electroreports-graph.key
openssl x509 -in ~/secrets/electroreports-graph.crt -noout -fingerprint -sha256
```

Expected: the last command prints `sha256 Fingerprint=AB:CD:...` (32 pairs). That is the thumbprint the app needs (a fingerprint is not a secret).

Upload **only the `.crt`**: Entra → the app → **Certificates & secrets → Certificates → Upload certificate** → choose `electroreports-graph.crt` → Add. Expected: the certificate is listed with an expiry about two years out. **Put a renewal reminder in your calendar for 11 months from now** (repeat this step with a new certificate, update `.env`, restart, then delete the old one).

## 3. API permission and admin consent (Global Administrator)

1. The app → **API permissions → Add a permission → Microsoft Graph → Application permissions**.
2. Add **`Sites.Selected`** and nothing else.
   **Do not add `Calendars.ReadWrite` here.** Calendar access is given through Exchange (section 4), which can restrict it to one group of mailboxes. A tenant-wide `Calendars.ReadWrite` consent would add to that and defeat the restriction.
3. **Grant admin consent for <your organisation>** → Yes.

Expected: one row, `Sites.Selected`, type Application, status **Granted for <organisation>** (green tick). `Sites.Selected` on its own gives access to **no** site until section 5.

## 4. Calendar access limited to one group (Exchange Online)

```powershell
Connect-ExchangeOnline -UserPrincipalName <admin-upn>

# 4a. The group: a mail-enabled security group. No nested groups.
New-DistributionGroup -Name "ElectroReports-SiteEngineers" -Alias electroreports-siteengineers -Type Security
Add-DistributionGroupMember -Identity "ElectroReports-SiteEngineers" -Member <engineer-upn>     # repeat for each engineer
Get-DistributionGroupMember -Identity "ElectroReports-SiteEngineers" | Select-Object DisplayName   # check the list
```
Expected: the group is created; the member list shows exactly the site engineers whose calendars should receive events (they all use `name@elegrow.com`).

```powershell
# 4b. Tell Exchange about the Entra app (a pointer to its service principal)
Connect-MgGraph -Scopes "Application.Read.All"
$sp = Get-MgServicePrincipal -Filter "appId eq '<APP_ID>'"
New-ServicePrincipal -AppId $sp.AppId -ObjectId $sp.Id -DisplayName "ElectroReports Graph"
```
Expected: an Exchange service principal with the app's id is printed. (If "already exists", carry on.)

```powershell
# 4c. The scope: mailboxes that are members of the group
$dn = (Get-DistributionGroup "ElectroReports-SiteEngineers").DistinguishedName
New-ManagementScope -Name "ElectroReports-SiteEngineers-Scope" -RecipientRestrictionFilter "MemberOfGroup -eq '$dn'"

# 4d. The permission, only inside that scope
New-ManagementRoleAssignment -App $sp.Id -Role "Application Calendars.ReadWrite" -CustomResourceScope "ElectroReports-SiteEngineers-Scope"
```
Expected: the scope and the role assignment are printed (`Role: Application Calendars.ReadWrite`, `CustomResourceScope: ElectroReports-SiteEngineers-Scope`).

```powershell
# 4e. Ask Exchange what it would allow
Test-ServicePrincipalAuthorization -Identity $sp.Id -Resource <engineer-upn-in-the-group>       # InScope : True
Test-ServicePrincipalAuthorization -Identity $sp.Id -Resource <someone-not-in-the-group>        # InScope : False
```
Expected: `InScope True` for a member and `InScope False` for a non-member. A change can take up to 30 minutes to apply; if the answer is wrong, wait and ask again.

**Fallback if RBAC for Applications is not available in the tenant** (legacy application access policy; here the Entra consent of `Calendars.ReadWrite` *is* needed, and the policy restricts it):

```powershell
New-ApplicationAccessPolicy -AppId <APP_ID> -PolicyScopeGroupId electroreports-siteengineers@<domain> -AccessRight RestrictAccess -Description "ElectroReports Graph"
Test-ApplicationAccessPolicy -Identity <engineer-upn-in-the-group> -AppId <APP_ID>      # AccessCheckResult: Granted
Test-ApplicationAccessPolicy -Identity <someone-not-in-the-group> -AppId <APP_ID>       # AccessCheckResult: Denied
```

What this permission cannot do: inside the group's mailboxes it can read every event, including private ones. The app only ever lists events carrying its own tag and re-checks the tag before it deletes (tested), but that part is the app's discipline, not a Microsoft limit.

## 5. Read access to the three SharePoint sites (Global Administrator)

One grant per site, role **`read`** only. The sites:

| Site | URL | Scope in the app |
|---|---|---|
| ETPL_OP_26_27 | `https://eletechnology.sharepoint.com/sites/ETPL_OP_26_27` | client (current projects) |
| FY23-24 | `https://eletechnology.sharepoint.com/sites/FY23-24` | client (older ESHA projects) |
| FY23-24-StandardsGuideline | `https://eletechnology.sharepoint.com/sites/FY23-24-StandardsGuideline` | general (only `Shared Documents/ETPL Knowledge Hub/02 Technical Knowledge` is indexed) |

```powershell
Connect-MgGraph -Scopes "Sites.FullControl.All"        # the Global Admin, for these grants only; sign out afterwards

foreach ($name in "ETPL_OP_26_27", "FY23-24", "FY23-24-StandardsGuideline") {
  $site = Get-MgSite -SiteId "eletechnology.sharepoint.com:/sites/$name"
  New-MgSitePermission -SiteId $site.Id -BodyParameter @{
    roles = @("read")
    grantedToIdentities = @(@{ application = @{ id = "<APP_ID>"; displayName = "ElectroReports Graph" } })
  } | Select-Object Id, Roles
}
```
Expected, three times: a permission Id and `Roles {read}`. Verify each site:

```powershell
Get-MgSitePermission -SiteId $site.Id | Select-Object Id, Roles, @{n="App";e={$_.GrantedToIdentitiesV2.Application.DisplayName}}
```
Expected: one row with `read` and `ElectroReports Graph`. Never `write`, `owner` or `fullcontrol`.

Alternative with PnP.PowerShell (same result): `Grant-PnPAzureADAppSitePermission -AppId <APP_ID> -DisplayName "ElectroReports Graph" -Site https://eletechnology.sharepoint.com/sites/<name> -Permissions Read`, checked with `Get-PnPAzureADAppSitePermission -Site <url>`.

Good to know:
* A grant is per **site**, not per folder, so the app can technically read the whole site; it indexes only what its rules allow (the knowledge hub folder; the Submitted, Reference and Input folders of project folders; never Working).
* A **private or shared Teams channel** has its own site collection and would need its own grant. The sources above are ordinary sites, so they are covered.
* Associating a site with a hub does not extend a grant to the associated sites.

## 6. Configure the server (`.env`, never committed)

```
GRAPH_USE_MOCK=false
GRAPH_TENANT_ID=<Directory (tenant) ID>
GRAPH_CLIENT_ID=<APP_ID>
GRAPH_CERT_PATH=/Users/<you>/secrets/electroreports-graph.key
GRAPH_CERT_THUMBPRINT_SHA256=<the fingerprint from section 2, colons allowed>
APP_BASE_URL=<where engineers open the app, once it is hosted; empty until then>
GRAPH_CALENDAR_TIMEZONE=India Standard Time
```
`GRAPH_KB_SITES` already lists the three sites in `.env.example`; copy that line unchanged. Restart the server. Do **not** set `GRAPH_CLIENT_SECRET`; the certificate is used.

## 7. Live checks (the 403 negative test)

These print statuses and counts only: no token, no file names, no document text.

```bash
node scripts/graph-check.js token
```
Expected: `PASS  a token was obtained and Graph answered ...`

```bash
node scripts/graph-check.js calendar --member <engineer-upn-in-the-group> --outside <mailbox-NOT-in-the-group>
```
Expected:
```
PASS  member mailbox readable (... event id returned)
PASS  outside mailbox refused (403 ErrorAccessDenied) - the restriction holds
ALL CHECKS PASSED
```
**If the second line says FAIL ("the OUTSIDE mailbox was readable"), the restriction is not in force: do not switch calendar sync on.** Re-check section 3 (no `Calendars.ReadWrite` consent in Entra) and section 4, wait 30 minutes, run again.

Optional, writes one event and removes it: add `--write-test` → expect `PASS  test event created ...` and `PASS  test event deleted again`.

```bash
node scripts/graph-check.js sites
```
Expected, for each of the three sites:
```
PASS  ETPL_OP_26_27 (client): site resolved, 1 document library
PASS  ETPL_OP_26_27 (client): delta feed readable (n items on the first page, ...)
...
PASS  a write to SharePoint is refused by the client before anything is sent
```
A `403` for a site means its grant (section 5) is missing; `404` means the URL in `GRAPH_KB_SITES` is wrong.

## 8. Switch the features on, carefully

1. **Knowledge base:** sign in as admin → Admin → **Knowledge base** → **Refresh now**. The first run reads every file once (it can take a while); afterwards refreshes only read what changed (delta). Check the counts per site; "No text" are scanned PDFs (not searched). Then try **Search documents**.
2. **Calendar sync:** Admin → **Calendar sync** → **Preview changes**. It lists what would be created, per owner, and warns about owners whose mailbox is not in the group. Nothing is written. When the list looks right: tick **Calendar sync is on** (it stays a *dry run*), then untick **Dry run** to let it write. Events carry a link only when `APP_BASE_URL` is set.
3. To stop: untick **Calendar sync is on**. Events already created stay in the calendars (they disappear on their own as tasks are closed while sync is on). A bulk "remove everything the app created" is not built; ask if you need it.

## 9. Undo

* Calendar: `Remove-ManagementRoleAssignment -Identity <assignment name from Get-ManagementRoleAssignment -App $sp.Id>`; remove the scope and the group if no longer wanted.
* SharePoint: `Remove-MgSitePermission -SiteId <site id> -PermissionId <id from Get-MgSitePermission>` per site.
* Entra: delete the app registration (this removes the service principal and all consent).

## 10. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `token request failed` / `invalid_client` / `AADSTS700027` | The thumbprint in `.env` is not the SHA-256 of the uploaded `.crt`, or the wrong certificate/key pair. Re-run section 2's fingerprint command on the `.crt` that was uploaded. |
| Token request answers 400 only now and then | Microsoft throttles very frequent token requests; the app caches a token for about an hour, so it only happens when restarting repeatedly. Wait a few minutes. |
| `403 accessDenied` on a site | No "read" grant on that site (section 5), or it was granted to another app. |
| `403 ErrorAccessDenied` on a member's calendar | The member is not in the group, or the role assignment has not propagated yet (up to 30 minutes), or the scope filter points at the wrong group. |
| `404 ErrorInvalidUser` | The engineer's Zoho e-mail is not their Microsoft 365 sign-in name (the app assumes they are equal). |
| `410 resyncRequired` in the log | A delta token expired; the app starts a full resync by itself. |
| Calendar link opens nothing | `APP_BASE_URL` is empty or the app is not hosted yet. The link `.../?project=<id>` opens the project once signed in. |
