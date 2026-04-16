import { google } from "googleapis"
import { cert, getApps, initializeApp } from "firebase-admin/app"
import { getStorage } from "firebase-admin/storage"

const FOLDER_ID = "19XXwrHjNNk-qLgxXBgM4wwuFA67muPwz"
const SHEET_FILE_NAME = "Website PDF Requests"

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
    const articleId = clean(body.articleId)
    const articleTitle = clean(body.articleTitle)
    const pagePath = clean(body.pagePath)
    const storagePath = clean(body.downloadLink)

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

    if (!storagePath) {
      return res.status(400).json({ error: "Missing storage path" })
    }

    // 🔥 Fetch PDF from Firebase
    const pdfBase64 = await fetchPdfBase64FromFirebase(storagePath)

    const articleLabel = articleTitle || "Requested PDF"

    // ✅ Send to user
    await sendViaResend({
      to: email,
      subject: `Your requested document: ${articleLabel}`,
      html: buildUserEmailHtml({ name, articleLabel }),
      attachments: [
        {
          filename: `${articleLabel}.pdf`,
          content: pdfBase64,
          type: "application/pdf",
          disposition: "attachment",
        },
      ],
    })

    // ✅ Notify internal
    await sendViaResend({
      to: "webcontact@resonant-grid.com",
      subject: `PDF download: ${articleLabel}`,
      html: buildInternalEmailHtml({
        name,
        company,
        role,
        email,
        articleId,
        articleLabel,
        pagePath,
        storagePath,
        submittedAt,
      }),
    })

    // ✅ Log to sheet
    await logRequestToGoogleSheet({
      submittedAt,
      articleId,
      articleLabel,
      name,
      company,
      role,
      email,
      source: "website",
      pagePath,
      storagePath,
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

//
// 🔥 FIREBASE
//
function getFirebaseApp() {
  if (getApps().length) return getApps()[0]

  const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT_JSON)

  return initializeApp({
    credential: cert({
      projectId: serviceAccount.project_id,
      clientEmail: serviceAccount.client_email,
      privateKey: serviceAccount.private_key.replace(/\\n/g, "\n"),
    }),
    storageBucket: "file-hosting-for-website.firebasestorage.app",
  })
}

async function fetchPdfBase64FromFirebase(storagePath) {
  const app = getFirebaseApp()
  const bucket = getStorage(app).bucket()
  const file = bucket.file(storagePath)

  console.log("Fetching Firebase file:", storagePath)

  const [exists] = await file.exists()

  if (!exists) {
    throw new Error(`Firebase file not found: ${storagePath}`)
  }

  const [buffer] = await file.download()

  const header = buffer.slice(0, 5).toString("utf8")

  if (header !== "%PDF-") {
    const preview = buffer.slice(0, 120).toString("utf8")
    throw new Error(
      `File is not a valid PDF. Preview: ${preview}`
    )
  }

  return buffer.toString("base64")
}

//
// ✉️ EMAIL
//
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

//
// 🧾 GOOGLE SHEETS
//
async function logRequestToGoogleSheet(row) {
  const auth = getGoogleAuth()
  const drive = google.drive({ version: "v3", auth })
  const sheets = google.sheets({ version: "v4", auth })

  const spreadsheetId = await getOrCreateSpreadsheetInFolder({ drive, sheets })

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: "Requests!A:J",
    valueInputOption: "USER_ENTERED",
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
        row.pagePath,
        row.storagePath,
      ]],
    },
  })
}

function getGoogleAuth() {
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
  const created = await drive.files.create({
    requestBody: {
      name: SHEET_FILE_NAME,
      mimeType: "application/vnd.google-apps.spreadsheet",
      parents: [FOLDER_ID],
    },
    fields: "id",
  })

  const spreadsheetId = created.data.id

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: "Requests!A1:J1",
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
        "Page Path",
        "Storage Path",
      ]],
    },
  })

  return spreadsheetId
}

function buildUserEmailHtml({ name, articleLabel }) {
  return `
    <p>Hi ${name || ""},</p>
    <p>Your requested document <strong>${articleLabel}</strong> is attached.</p>
    <p>Regards,<br/>Resonant Grid</p>
  `
}

function buildInternalEmailHtml(data) {
  return `<pre>${JSON.stringify(data, null, 2)}</pre>`
}
