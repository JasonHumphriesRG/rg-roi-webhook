import { Resend } from "resend"
import { google } from "googleapis"
import PDFDocument from "pdfkit"
import { PassThrough, Readable } from "stream"

let resend = null

function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*")
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS")
  res.setHeader("Access-Control-Allow-Headers", "Content-Type")
}

function formatCurrency(value, currency = "GBP") {
  try {
    return new Intl.NumberFormat("en-GB", {
      style: "currency",
      currency,
      maximumFractionDigits: 0,
    }).format(value ?? 0)
  } catch {
    return `${currency} ${Number(value ?? 0).toLocaleString()}`
  }
}

async function loadLogoBuffer() {
  if (!process.env.PDF_LOGO_URL) return null
  const response = await fetch(process.env.PDF_LOGO_URL)
  if (!response.ok) throw new Error("Failed to load logo from PDF_LOGO_URL")
  const arrayBuffer = await response.arrayBuffer()
  return Buffer.from(arrayBuffer)
}

function buildPdfBuffer(data) {
  return new Promise(async (resolve, reject) => {
    try {
      const doc = new PDFDocument({ margin: 40, size: "A4" })
      const stream = new PassThrough()
      const chunks = []

      stream.on("data", (chunk) => chunks.push(chunk))
      stream.on("end", () => resolve(Buffer.concat(chunks)))
      stream.on("error", reject)

      doc.pipe(stream)

      const logoBuffer = await loadLogoBuffer().catch(() => null)

      doc.roundedRect(40, 35, 515, 90, 16).fillAndStroke("#FFFFFF", "#E6EDF3")

      if (logoBuffer) {
        doc.image(logoBuffer, 58, 52, { fit: [110, 40] })
      }

      doc.fillColor("#0B1B3A").font("Helvetica-Bold").fontSize(22).text("REE ROI Submission", 190, 52)
      doc.fillColor("#6E7A8B").font("Helvetica").fontSize(11).text(
        "Calculator summary, submission details, business inputs, assumptions, and ROI outputs.",
        190,
        82,
        { width: 300 }
      )

      doc.moveDown(8)
      doc.font("Helvetica").fontSize(11).fillColor("#0B1B3A")
      doc.text(`Submitted: ${data.submittedAt || new Date().toISOString()}`)
      doc.text(`Name: ${data.contact?.name || ""}`)
      doc.text(`Email: ${data.contact?.email || ""}`)
      doc.text(`Job Title: ${data.contact?.jobTitle || ""}`)
      doc.moveDown()

      doc.font("Helvetica-Bold").text("Business Inputs")
      doc.font("Helvetica")
      Object.entries(data.calculator?.inputs || {}).forEach(([k, v]) => {
        doc.text(`${k}: ${String(v)}`)
      })

      doc.moveDown()
      doc.font("Helvetica-Bold").text("Scenario Assumptions")
      doc.font("Helvetica")
      Object.entries(data.calculator?.assumptions || {}).forEach(([k, v]) => {
        doc.text(`${k}: ${String(v)}`)
      })

      doc.moveDown()
      doc.font("Helvetica-Bold").text("Outputs")
      doc.font("Helvetica")
      Object.entries(data.calculator?.outputs || {}).forEach(([k, v]) => {
        doc.text(`${k}: ${String(v)}`)
      })

      doc.moveDown()
      doc.font("Helvetica-Bold").text("Cover Note")
      doc.font("Helvetica").text(data.contact?.coverNote || "-")

      doc.end()
    } catch (err) {
      reject(err)
    }
  })
}

async function uploadToDrive(pdfBuffer, fileName) {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON)

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/drive.file"],
  })

  const drive = google.drive({ version: "v3", auth })

  const response = await drive.files.create({
    requestBody: {
      name: fileName,
      parents: [process.env.GOOGLE_DRIVE_FOLDER_ID],
    },
    media: {
      mimeType: "application/pdf",
      body: Readable.from(pdfBuffer),
    },
    fields: "id,name,webViewLink",
    supportsAllDrives: true,
  })

  return response.data
}

export default async function handler(req, res) {
  setCors(res)

  if (req.method === "OPTIONS") {
    return res.status(200).end()
  }

  try {
    if (!process.env.RESEND_API_KEY) {
      return res.status(500).json({ error: "Missing RESEND_API_KEY" })
    }

    if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
      return res.status(500).json({ error: "Missing GOOGLE_SERVICE_ACCOUNT_JSON" })
    }

    if (!process.env.GOOGLE_DRIVE_FOLDER_ID) {
      return res.status(500).json({ error: "Missing GOOGLE_DRIVE_FOLDER_ID" })
    }

    if (!resend) {
      resend = new Resend(process.env.RESEND_API_KEY)
    }

    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" })
    }

    const data = req.body
    const email = data?.contact?.email

    if (!email) {
      return res.status(400).json({ error: "Missing contact email" })
    }

    const timestamp = new Date().toISOString()
    const safeEmail = String(email).replace(/[^a-zA-Z0-9]/g, "_")
    const fileName = `${timestamp}__${safeEmail}__roi.pdf`

    const pdfBuffer = await buildPdfBuffer(data)
    const driveFile = await uploadToDrive(pdfBuffer, fileName)

    const internalEmail = await resend.emails.send({
      from: "sales@resonant-grid.com",
      to: "jason@resonant-grid.com",
      replyTo: email,
      subject: `ROI Enquiry — ${data?.calculator?.utilityName || "Unknown Utility"}`,
      html: `
        <h2>New ROI Submission</h2>
        <p><strong>Name:</strong> ${data?.contact?.name || ""}</p>
        <p><strong>Email:</strong> ${email}</p>
        <p><strong>Utility:</strong> ${data?.calculator?.utilityName || ""}</p>
        <p><strong>ROI:</strong> ${data?.calculator?.outputs?.roiMultiple || ""}x</p>
        <p><strong>Drive File:</strong> ${driveFile.webViewLink || driveFile.id}</p>
      `,
    })

    if (internalEmail?.error) {
      throw new Error(
        `Internal email failed: ${internalEmail.error.message || JSON.stringify(internalEmail.error)}`
      )
    }

    const confirmationEmail = await resend.emails.send({
      from: "sales@resonant-grid.com",
      to: email,
      subject: "We’ve received your ROI submission",
      html: `
        <p>Thanks ${data?.contact?.name || ""},</p>
        <p>We’ve received your ROI scenario and will get back to you shortly.</p>
      `,
    })

    if (confirmationEmail?.error) {
      throw new Error(
        `Confirmation email failed: ${confirmationEmail.error.message || JSON.stringify(confirmationEmail.error)}`
      )
    }

    return res.status(200).json({
      success: true,
      fileId: driveFile.id,
      fileName: driveFile.name,
      driveUrl: driveFile.webViewLink,
    })
  } catch (error) {
    console.error(error)
    return res.status(500).json({
      error: error.message || "Server error",
    })
  }
}
