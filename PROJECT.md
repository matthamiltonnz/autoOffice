# Project Knowledge Capture

## Project

`autoOffice` is a fork of `Sivan22/autoOffice`: an MIT-licensed Office task-pane
add-in that generates and executes real `office.js` code against the open Word,
Excel, PowerPoint or Outlook document. The agent has two built-in tools
(`lookup_skill`, `execute_code`), runs the generated code in a sandboxed iframe,
and self-heals on failure.

This fork is self-hosted on Matt's personal GitHub account:

- Repo: <https://github.com/matthamiltonnz/autoOffice>
- Task-pane assets: <https://matthamiltonnz.github.io/autoOffice/>

## Current State

Deployed and serving from GitHub Pages. The Word / Excel / PowerPoint path is
complete. The Outlook path had four independent defects, all now fixed and
deployed; the last one (the compose form) has **not** been confirmed working in
a live Outlook client.

Commits deployed on `master` for this work:

```
2cd5213  docs: correct the Outlook sideload steps and note the required forms
aa0e00f  outlook: declare an ItemEdit form so compose mode renders
d39df9a  outlook: declare AppDomains for the remote task pane origin
124ee63  build: stop burying iframe.html in dist/src
26a82ed  outlook: fix MailApp VersionOverrides namespace so Outlook loads
```

The production Outlook manifest is sideloaded into
`~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef/manifest.xml` and
points at GitHub Pages, so Outlook no longer depends on a local dev server.

## Core Decision Record

- **Self-host on Matt's personal account, not the `positiveitnz` org.** The org
  exists and hosts other Positive IT repositories, but this fork is personal.
  The Pages origin is therefore `matthamiltonnz.github.io/autoOffice/`.
- **Keep the repo named `autoOffice`.** `deploy.yml` hardcodes
  `VITE_BASE: /autoOffice/`; renaming the repo breaks every asset path unless
  that value changes in the same commit.
- **Give every manifest in this fork fresh GUIDs.** Upstream's `<Id>` values and
  the installer's `AppId` / catalog GUIDs were reused verbatim in the fork,
  which would have made this build collide with an upstream install. The
  installer's originals were also sequential placeholders
  (`...78902`, `...78903`), so they were never unique to begin with.
- **One manifest per host family.** An add-in manifest can carry only one
  `<VersionOverrides>`, so Mailbox (Outlook) and the task-pane hosts need
  separate files. Both point at the same web app.
- **Publish exclusively through GitHub Pages; nothing runs.** There is no
  backend process, no database and no state to keep alive, so Docker is not a
  hosting candidate here — it would add a process to babysit where there is
  currently none, plus a TLS problem, because Office rejects self-signed certs
  on non-`localhost` origins.

## Architecture

- **Hosting:** `deploy.yml` builds `dist/` and publishes it to Pages;
  `release.yml` builds `AutoOffice-Setup.exe` with Inno Setup on
  `windows-latest`. `dist/` is gitignored and is produced fresh by CI.
- **Manifests:** `manifest.xml` (dev, Word/Excel/PowerPoint),
  `manifest.outlook.xml` (dev, Outlook/Mailbox), `manifest.production.xml` and
  `manifest.outlook.production.xml` (hosted, Pages URLs).
- **Task pane:** React 19 + Vite, Fluent UI, Vercel AI SDK. `src/taskpane/index.tsx`
  renders **only inside the `Office.onReady` callback**:

  ```ts
  if (typeof Office !== 'undefined') {
    Office.onReady(() => start());
  } else {
    start();
  }
  ```

  Consequence: if `Office` exists but `onReady` never fires, `start()` never
  runs and the pane is **blank with no error**. A silently empty pane is a
  host-loading symptom, not a render symptom — `start()` catches errors and
  would render an error shell instead.

## Manifest and Build Traps

Four defects, all of which failed **silently**. They are recorded here because
none of them produces an error message.

1. **MailApp `VersionOverrides` namespace.** A `MailApp` manifest that carries a
   `<VersionOverrides>` element must bind `mailappor` to
   `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`, not the
   v1.0 namespace. With v1.0, Outlook rejects the manifest during registration
   and the add-in never appears — no dialog, no log. The XML remains
   well-formed, so a parser accepts it; only Outlook's validator objects.
2. **`vite-plugin-static-copy` v4 directory structure.** For glob entries the
   plugin preserves the matched file's directory structure under `dest`, so the
   sandbox page landed at `dist/src/taskpane/executor/iframe.html` instead of
   `dist/iframe.html`. The build still logged `Copied 1 items` and exited 0.
   The executor's iframe then 404'd at runtime, which breaks code execution in
   **every** deployed build, on every host. Fix: `rename: { stripBase: true }`.
   Note `dest: 'dist/'` is not the fix — `dest` resolves from `build.outDir`,
   producing `dist/dist/`.
3. **`FormSettings` needs a form per surface.** The manifest declared
   `MessageComposeCommandSurface` with a `ShowTaskpane` button but only an
   `ItemRead` form. Opening the add-in from the Apps pane while composing
   therefore gave it a button with nothing to load: an empty pane, no error.
   Both `ItemRead` and `ItemEdit` are declared now.
4. **A stray `static.yml` clobbered the real deployment.** GitHub's "Static
   HTML" starter workflow had been committed to `master`. It uploads
   `path: '.'` — the repo root — but `dist/` is gitignored, so it published raw
   source referencing `/src/taskpane/index.tsx`, and its `concurrency: pages`
   group let it win the race against `deploy.yml`. Symptom: the site 404s or
   serves an unexecutable page. Removed; only `ci.yml`, `deploy.yml` and
   `release.yml` remain.

Additional notes:

- **`AppDomains`** is declared for `https://matthamiltonnz.github.io` in the
  production Outlook manifest. `localhost` is implicitly trusted, which is why
  the dev manifest never needed it. The change is deployed but its effect is
  **unverified** — it is the one fix here not backed by an observed result.
- **Sideloading moved.** `Add from URL` no longer exists for Outlook. On
  Outlook for Mac 16.85+ the ribbon **Get Add-ins** button opens the Microsoft
  Marketplace in the browser instead of the add-in dialog. Use
  <https://aka.ms/olksideload>, which opens the *Add-Ins for Outlook* dialog
  directly, then **My add-ins → Custom Addins → Add a custom add-in → Add from
  File**.
- **Jekyll is not a factor.** `dist/` contains no underscore-prefixed paths,
  and all 13 built files return 200 from Pages with no `.nojekyll` required.

## Verification Performed

Checked on 2026-09-29, after the final deploy:

| Check | Result |
| --- | --- |
| All 13 files in `dist/` against the live site | every one returns 200 |
| `iframe.html` | 200, serving the real sandbox page |
| Icons 16/32/64/80 | valid PNGs, correct dimensions and byte sizes |
| Mount point | `index.html` has `id="root"`, matched by `getElementById('root')` |
| `office.js` CDN, pinned `/1.1/` and latest | both 200, 69566 bytes |
| Bundle `access-control-allow-origin` | `*` |
| `X-Frame-Options` / CSP on the pane URL | absent, so the pane is frameable |
| Manifest XML well-formedness (all four) | parse OK |
| `ItemRead` / `ItemEdit` per Outlook manifest | 1 / 1 |
| Deploy and release workflows | both `success` |
| `npx vitest run` | 223 passed (223), exit 0 |
| `npm run check:i18n` | exit 0 |

Ruled out along the way, each by direct check rather than assumption: a missing
`HighResolutionIconUrl` (present in none of the manifests, so not a
differentiator), a CORS failure on the bundle (`*` is returned), a bad mount
point, missing icons, Jekyll stripping, and a stale or wrong manifest in the
`wef` folder (the sideloaded file is byte-identical to
`manifest.outlook.production.xml`).

## Next Steps

1. **Confirm the Outlook pane renders.** Quit Outlook with `Cmd+Q` (the `wef`
   folder is read only at launch), reopen, and open the add-in in both a new
   message (compose) and an opened email (read mode). Compose is the case that
   had no form declared.
2. **If it is still blank**, open <https://matthamiltonnz.github.io/autoOffice/>
   in a normal browser tab. An error message means the bundle is fine and the
   fault is Office-side; a blank page means the bundle is at fault and can be
   debugged from the repo. In Outlook on the web, change the DevTools console
   context to the `github.io` frame first — the pane is a cross-origin iframe,
   so the default `top` context shows none of the add-in's own errors.
3. **Consider reverting `d39df9a`** if `AppDomains` turns out to be irrelevant.
   It is the only unverified change.
4. **The `iframe.html` fix is the highest-value upstream contribution.** It is a
   latent upstream defect affecting every host and every deployment, not
   something introduced by this fork. Worth a PR to `Sivan22/autoOffice`.
5. Optional hardening: `src/taskpane/index.tsx` could render a visible error
   shell if `Office.onReady` does not fire within a few seconds, rather than
   leaving a blank pane. A silent blank pane is the worst available failure mode
   for a task-pane add-in.

## Boundaries

- Do not commit `dist/` or hand-publish it; it is gitignored and produced by CI.
- Do not rename the repository without updating `VITE_BASE` in `deploy.yml` in
  the same commit.
- Do not reuse upstream's manifest `<Id>` values or installer GUIDs — this build
  must stay installable side by side with upstream.
- Do not add a second Pages workflow. One deployment path, or they race.
- Do not treat `localhost` behaviour as evidence about the hosted build: Office
  implicitly trusts `localhost` for `AppDomains`, so the dev manifest hides
  defects that only appear on a remote origin.
- Docker is not part of hosting. Pages serves static files; there is nothing to
  keep running.



