import { google } from "googleapis"

const FOLDER_ID = "19XXwrHjNNk-qLgxXBgM4wwuFA67muPwz"
const SHEET_FILE_NAME = "Website PDF Requests"

const ARTICLES = {
  "virtual-rtu-whitepaper": {
    subject: "Your Virtual RTU White Paper",
    pdfUrl: "https://www.resonant-grid.com/pdfs/virtual-rtu-whitepaper.pdf",
    filename: "virtual-rtu-whitepaper.pdf",
    label: "Virtual RTU White Paper",
  },
  "default-article": {
    subject: "Your requested document",
    pdfUrl: "https://www.resonant-grid.com/pdfs/default.pdf",
    filename: "document.pdf",
    label: "Requested PDF",
  },
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*")
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS")
  res.setHeader("Access-Control-Allow-Headers", "Content-Type")

  if (req.method === "OPTIONS") {
    return res.status(200).end()
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" })
  }

  try {
    const body = req.body || {}
    const type = body.type
    const articleId = body.articleId || "default-article"
    const contact = body.contact || {}

    const name = clean(contact.name)
    const company = clean(contact.company)
    const role = clean(contact.role)
    const email = clean(contact.email)
    const submittedAt = body.submittedAt || new Date().toISOString()

    if (type !== "pdf_request") {
      return res.status(400).json({ error: "Invalid request type" })
    }

    if (!email) {
      return res.status(400).json({ error: "Email is required" })
    }

    const article = ARTICLES[articleId] || ARTICLES["default-article"]

    const pdfResponse = await fetch(article.pdfUrl)
    if (!pdfResponse.ok) {
      throw new Error(`Failed to fetch PDF from ${article.pdfUrl}`)
    }

    const pdfArrayBuffer = await pdfResponse.arrayBuffer()
    const pdfBase64 = Buffer.from(pdfArrayBuffer).toString("base64")

    await sendViaResend({
      to: email,
      subject: article.subject,
      html: buildUserEmailHtml({ name, articleLabel: article.label }),
      attachments: [
        {
          filename: article.filename,
          content: pdfBase64,
          type: "application/pdf",
          disposition: "attachment",
        },
      ],
    })

    await sendViaResend({
      to: "webcontact@resonant-grid.com",
      subject: `PDF download: ${article.label}`,
      html: buildInternalEmailHtml({
        name,
        company,
        role,
        email,
        articleId,
        articleLabel: article.label,
        submittedAt,
      }),
    })

    await logRequestToGoogleSheet({
      submittedAt,
      articleId,
      articleLabel: article.label,
      name,
      company,
      role,
      email,
      source: "website",
    })

    return res.status(200).json({ ok: true })
  } catch (error) {
    console.error("send-pdf-request error:", error)

    return res.status(500).json({
      error: "Failed to process PDF request",
      detail: error?.message || "Unknown error",
    })
  }
}

function clean(value) {
  return typeof value === "string" ? value.trim() : ""
}

function buildUserEmailHtml({ name, articleLabel }) {
  return `
    <div style="font-family: Arial, sans-serif; line-height: 1.5; color: #12312B;">
      <p>Hi${name ? " " + escapeHtml(name) : ""},</p>
      <p>Thanks for your interest in <strong>${escapeHtml(articleLabel)}</strong>. Your requested PDF is attached.</p>
      <p>If you have any questions, just reply to this email.</p>
      <p>Best regards,<br/>Resonant Grid</p>
    </div>
  `
}

function buildInternalEmailHtml({
  name,
  company,
  role,
  email,
  articleId,
  articleLabel,
  submittedAt,
}) {
  return `
    <div style="font-family: Arial, sans-serif; line-height: 1.5; color: #12312B;">
      <p>A PDF has been requested from the website.</p>
      <table cellpadding="6" cellspacing="0" border="0">
        <tr><td><strong>Document</strong></td><td>${escapeHtml(articleLabel)}</td></tr>
        <tr><td><strong>Article ID</strong></td><td>${escapeHtml(articleId)}</td></tr>
        <tr><td><strong>Name</strong></td><td>${escapeHtml(name || "-")}</td></tr>
        <tr><td><strong>Company</strong></td><td>${escapeHtml(company || "-")}</td></tr>
        <tr><td><strong>Role</strong></td><td>${escapeHtml(role || "-")}</td></tr>
        <tr><td><strong>Email</strong></td><td>${escapeHtml(email)}</td></tr>
        <tr><td><strong>Submitted</strong></td><td>${escapeHtml(submittedAt)}</td></tr>
      </table>
    </div>
  `
}

async function sendViaResend({ to, subject, html, attachments = [] }) {
  if (!process.env.RESEND_API_KEY) {
    throw new Error("Missing RESEND_API_KEY")
  }

  if (!process.env.MAIL_FROM) {
    throw new Error("Missing MAIL_FROM")
  }

  const response = await fetch("https://api.resend.com/emails", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${process.env.RESEND_API_KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      from: process.env.MAIL_FROM,
      to: Array.isArray(to) ? to : [to],
      subject,
      html,
      attachments,
    }),
  })

  if (!response.ok) {
    const text = await response.text()
    throw new Error(`Resend error: ${text}`)
  }
}

async function logRequestToGoogleSheet(row) {
  const auth = getGoogleAuth()
  const drive = google.drive({ version: "v3", auth })
  const sheets = google.sheets({ version: "v4", auth })

  const spreadsheetId = await getOrCreateSpreadsheetInFolder({ drive, sheets })

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: "Requests!A:H",
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: {
      values: [[
        row.submittedAt,
        row.articleId,
        row.articleLabel,
        row.name,
        row.company,
        row.role,
        row.email,
        row.source,
      ]],
    },
  })
}

function getGoogleAuth() {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    throw new Error("Missing GOOGLE_SERVICE_ACCOUNT_JSON")
  }

  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON)

  return new google.auth.GoogleAuth({
    credentials,
    scopes: [
      "https://www.googleapis.com/auth/drive",
      "https://www.googleapis.com/auth/spreadsheets",
    ],
  })
}

async function getOrCreateSpreadsheetInFolder({ drive, sheets }) {
  const query = [
    `'${FOLDER_ID}' in parents`,
    `name='${escapeForDriveQuery(SHEET_FILE_NAME)}'`,
    `mimeType='application/vnd.google-apps.spreadsheet'`,
    `trashed=false`,
  ].join(" and ")

  const existing = await drive.files.list({
    q: query,
    fields: "files(id,name)",
    pageSize: 1,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
  })

  const existingFile = existing.data.files?.[0]
  if (existingFile?.id) {
    return existingFile.id
  }

  const created = await drive.files.create({
    requestBody: {
      name: SHEET_FILE_NAME,
      mimeType: "application/vnd.google-apps.spreadsheet",
      parents: [FOLDER_ID],
    },
    fields: "id",
    supportsAllDrives: true,
  })

  const spreadsheetId = created.data.id
  if (!spreadsheetId) {
    throw new Error("Failed to create Google Sheet")
  }

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [
        {
          updateSheetProperties: {
            properties: {
              sheetId: 0,
              title: "Requests",
            },
            fields: "title",
          },
        },
      ],
    },
  })

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: "Requests!A1:H1",
    valueInputOption: "RAW",
    requestBody: {
      values: [[
        "Submitted At",
        "Article ID",
        "Document",
        "Name",
        "Company",
        "Role",
        "Email",
        "Source",
      ]],
    },
  })

  return spreadsheetId
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;")
}

function escapeForDriveQuery(value) {
  return String(value).replaceAll("\\", "\\\\").replaceAll("'", "\\'")
}
