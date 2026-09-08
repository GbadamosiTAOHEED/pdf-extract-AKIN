# PDF Extraction Suite Pro

A Streamlit app that OCRs a scanned PDF directory (e.g. a health facility or
organization document listing), cleans up the text, and exports a structured Word
document and/or Excel spreadsheet.

## Setup

1. **Install Python packages**
   ```bash
   pip install -r requirements.txt
   ```

2. **Install the two external tools** (these are *not* pip packages):
   - **Tesseract OCR** — https://github.com/UB-Mannheim/tesseract/wiki (Windows installer)
     macOS: `brew install tesseract` · Linux: `sudo apt install tesseract-ocr`
   - **Poppler** (used by `pdf2image`) — https://github.com/oschwartz10612/poppler-windows/releases (Windows)
     macOS: `brew install poppler` · Linux: `sudo apt install poppler-utils`

   On macOS/Linux, once installed via a package manager, both tools are
   already on your `PATH` — you can skip step 3 and leave the sidebar path
   fields blank.

3. **Point the app at your local install (Windows, or a custom install location)**
   Copy `.env.example` to `.env` and fill in your paths:
   ```
   TESSERACT_CMD=C:\Program Files\Tesseract-OCR\tesseract.exe
   POPPLER_PATH=C:\path\to\poppler-xx.xx.x\Library\bin
   ```
   `.env` is git-ignored, so your personal paths never get committed. If you
   skip this, you can still type the paths directly into the app's sidebar
   each time you run it.

## Run

```bash
streamlit run pdf_extraction_suite.py
```

> Run it with `streamlit run`, **not** `python pdf_extraction_suite.py` —
> launching it as a plain script produces harmless-but-noisy
> `missing ScriptRunContext` warnings in the terminal because Streamlit
> isn't running inside its own server in that mode.

This opens the app in your browser at `http://localhost:8501`. Upload a
scanned PDF, adjust the DPI/quality settings in the sidebar if needed, and
click **Start processing**.

## Notes on speed

- **DPI 200** (default) is enough for most printed directories; only raise
  it if the source text is genuinely tiny.
- **Preprocessing quality**: leave on *Fast* or *Balanced* for clean,
  digitally printed pages. *Max accuracy* adds a slow denoise pass — only
  worth it on noisy photocopies.
- Use **"Only process the first N pages"** to test your settings on a big
  file before committing to a full run.

# REPOSITORY SHARING AND PERMISSION PROTOCOL

**Owner:** Gbadamosi Taoheed Alade
**Ownership Status:** Private Property / Proprietary Work
**Effective Date:** 8 September 2026

## 1. Purpose

This protocol establishes the rules governing the sharing, distribution, copying, modification, and use of this repository.

The repository and its contents are the property of **Gbadamosi Taoheed Alade** ("the Owner"). Access granted to an individual or team does not constitute a transfer of ownership or an unrestricted right to redistribute the repository.

## 2. Ownership

This repository is the intellectual and proprietary work of:

**Gbadamosi Taoheed Alade**

All source code, documentation, configurations, data structures, project materials, methodologies, and other original materials contained within the repository remain the property of the Owner unless otherwise expressly stated in writing.

Sharing the repository with a person, team, organization, collaborator, or stakeholder does **not** transfer ownership to that recipient.

## 3. Permission to Share

The repository may be shared with authorized team members or collaborators for legitimate project purposes.

However, **any person who receives access to this repository must obtain prior permission from Gbadamosi Taoheed Alade before sharing it with any other person, team, organization, platform, or third party.**

Permission to access the repository does not automatically include permission to redistribute it.

### The following actions require prior permission from the Owner:

* Sharing the repository with another person.
* Forwarding or distributing repository files.
* Uploading the repository to another GitHub/GitLab/Bitbucket account or public platform.
* Making the repository public.
* Providing access to external collaborators or organizations.
* Republishing or reproducing substantial portions of the repository.
* Using the repository for a separate project outside the originally authorized purpose.
* Commercializing or monetizing the repository or any substantial part of it.
* Creating a derivative project for external distribution.

## 4. Team Access

Team members may access and use the repository only for the purpose for which access was granted.

A team member must not assume that because they have access to the repository, they have authority to grant access to others.

**Access is personal and non-transferable unless otherwise authorized by the Owner.**

If another person requires access, the team member should direct the request to:

**Gbadamosi Taoheed Alade**

No team member should independently create or distribute access credentials, repository invitations, copies, archives, or download packages to third parties without authorization.

## 5. Public Sharing

The repository must **not** be made public without the explicit approval of the Owner.

This includes publishing the repository, code snippets, documentation, screenshots, datasets, technical architecture, or other substantial project materials on:

* GitHub
* GitLab
* Bitbucket
* Personal websites
* Blogs
* Social media
* Public forums
* Online repositories
* Cloud storage with public access
* Any other publicly accessible platform

## 6. Attribution

Where the repository or substantial portions of its contents are legitimately used or shared with permission, appropriate attribution should be maintained.

The ownership attribution should identify:

**Gbadamosi Taoheed Alade**

Removal, alteration, or replacement of ownership notices without authorization is prohibited.

## 7. Confidentiality and Responsible Handling

Recipients of the repository are expected to take reasonable measures to prevent unauthorized access, copying, or distribution.

Repository contents should not be unnecessarily copied to personal devices, public storage, external repositories, or third-party systems.

Where sensitive credentials, API keys, configuration secrets, private datasets, or other confidential information exist, they must not be exposed or redistributed.

## 8. Requesting Permission

Anyone wishing to share the repository or its contents with another person or organization should first obtain approval from the Owner.

Permission requests should clearly state:

1. Who will receive the repository or material.
2. Why access is required.
3. What portion of the repository will be shared.
4. How it will be used.
5. Whether the recipient will further distribute or modify it.

Permission should preferably be granted in writing so that there is a clear record of authorization.

## 9. Unauthorized Sharing

Unauthorized redistribution of the repository may result in the withdrawal of access and other appropriate action available to the Owner.

Where unauthorized sharing occurs, the recipient may be required to:

* Stop further distribution.
* Remove unauthorized copies where reasonably possible.
* Revoke or terminate unauthorized access.
* Notify the Owner of where and with whom the repository was shared.
* Cooperate in addressing the unauthorized distribution.

## 10. No Transfer of Ownership

Nothing in this sharing protocol should be interpreted as transferring ownership of the repository to any recipient.

Unless a separate written agreement expressly provides otherwise:

**Ownership remains exclusively with Gbadamosi Taoheed Alade.**

Permission to use or access the repository is a limited authorization and does not constitute a sale, assignment, license for unrestricted redistribution, or transfer of intellectual property rights.

## 11. Acceptance

By receiving access to this repository, the recipient acknowledges that:

> **This repository is the property of Gbadamosi Taoheed Alade. Access granted to me does not give me the right to redistribute, publish, transfer, or share the repository with another person or third party without obtaining prior permission from Gbadamosi Taoheed Alade.**


**Repository Owner:**
**Gbadamosi Taoheed Alade**

**Ownership:** Proprietary / Private

**Permission Required for Redistribution:** **YES**

**Unauthorized Redistribution:** **NOT PERMITTED**
