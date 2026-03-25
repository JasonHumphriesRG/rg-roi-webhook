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

function buildPdfBuffer(data) {
  return new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({ margin: 50 })
      const stream = new PassThrough()
      const chunks = []

      stream.on("data", (chunk) => chunks.push(chunk))
      stream.on("end", () => resolve(Buffer.concat(chunks)))
      stream.on("error", reject)

      doc.pipe(stream)

      doc.fontSize(20).text("REE ROI Submission Summary")
      doc.moveDown()

      doc.fontSize(12).text(
        `Submitted: ${data.submittedAt || new Date().toISOString()}`
      )
      doc.text(`Name: ${data.contact?.name || ""}`)
      doc.text(`Email: ${data.contact?.email || ""}`)
      doc.text(`Job Title: ${data.contact?.jobTitle || ""}`)
      doc.moveDown()

      doc.text(`Utility: ${data.calculator?.utilityName || ""}`)
      doc.text(`Market: ${data.calculator?.market || ""}`)
      doc.text(`Scenario: ${data.calculator?.scenarioLabel || ""}`)
      doc.moveDown()

      doc.fontSize(14).text("Outputs")
      doc.fontSize(12).text(
        `Total Annual Value: ${data.calculator?.outputs?.totalAnnualValue ?? ""}`
      )
      doc.text(`REE Annual Fee: ${data.calculator?.outputs?.reeAnnualFee ?? ""}`)
      doc.text(
        `Net Annual Benefit: ${data.calculator?.outputs?.netAnnualBenefit ?? ""}`
      )
      doc.text(`ROI Multiple: ${data.calculator?.outputs?.roiMultiple ?? ""}`)
      doc.text(`Payback Months: ${data.calculator?.outputs?.paybackMonths ?? ""}`)
      doc.moveDown()

      doc.fontSize(14).text("Cover Note")
      doc.fontSize(12).text(data.contact?.coverNote || "-")

      doc.end()
    } catch (err) {
      reject(err)
    }
  })
}

async function uploadToDrive(pdfBuffer, fileName) {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    throw new Error("Missing GOOGLE_SERVICE_ACCOUNT_JSON")
  }

  if (!process.env.GOOGLE_DRIVE_FOLDER_ID) {
    throw new Error("Missing GOOGLE_DRIVE_FOLDER_ID")
  }

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

    await resend.emails.send({
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

    await resend.emails.send({
      from: "sales@resonant-grid.com",
      to: email,
      subject: "We’ve received your ROI submission",
      html: `
        <p>Thanks ${data?.contact?.name || ""},</p>
        <p>We’ve received your ROI scenario and will get back to you shortly.</p>
      `,
    })

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
