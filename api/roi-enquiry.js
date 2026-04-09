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

// ---------- NEW: dynamic page size ----------
function getPageSize(market) {
  const m = String(market || "").toLowerCase()
  if (m.includes("us")) return "LETTER"
  return "A4"
}

// ---------- NEW: filename builder ----------
function buildFileName(utilityName) {
  const safeUtility = String(utilityName || "Unknown Utility")
    .replace(/[^a-zA-Z0-9]/g, "_")

  const now = new Date()
  const ts = now.toISOString()
    .replace(/[-:]/g, "")
    .replace("T", "_")
    .slice(0, 15)

  return `${safeUtility}_Resonant_Grid_Sentinel_Evidence_Engine_${ts}.pdf`
}

// ---------- formatting helpers ----------
function formatCurrency(value, currency = "GBP") {
  return new Intl.NumberFormat("en-GB", {
    style: "currency",
    currency,
    maximumFractionDigits: 0,
  }).format(value ?? 0)
}

function formatNumber(value) {
  return Number(value ?? 0).toLocaleString("en-GB")
}

function formatPercentWhole(value) {
  return `${Math.round(Number(value ?? 0))}%`
}

function formatMultiple(value) {
  return `${Number(value ?? 0).toFixed(1)}x`
}

function formatMonths(value) {
  if (!isFinite(value)) return "—"
  return `${Number(value).toFixed(1)} months`
}

// ---------- PDF ----------
function buildPdfBuffer(data) {
  return new Promise((resolve, reject) => {
    try {
      const pageSize = getPageSize(data?.calculator?.market)

      const doc = new PDFDocument({
        margin: 36, // tighter margins to fit 1 page
        size: pageSize,
      })

      const stream = new PassThrough()
      const chunks = []

      stream.on("data", (c) => chunks.push(c))
      stream.on("end", () => resolve(Buffer.concat(chunks)))
      stream.on("error", reject)

      doc.pipe(stream)

      const pageWidth =
        doc.page.width - doc.page.margins.left - doc.page.margins.right
      const left = doc.page.margins.left

      const currency = data.calculator?.currency || "GBP"
      const inputs = data.calculator?.inputs || {}
      const assumptions = data.calculator?.assumptions || {}
      const outputs = data.calculator?.outputs || {}

      // ---------- HEADER FIX ----------
      doc.fontSize(18).text("REE ROI Submission", left, 40)

      doc
        .fontSize(9)
        .fillColor("grey")
        .text(
          `Submitted: ${new Date(data.submittedAt || Date.now()).toLocaleString()}`,
          left,
          62
        )

      doc
        .fontSize(10)
        .fillColor("black")
        .text(data.calculator?.utilityName || "", left, 75)

      let y = 100

      // ---------- KPI STRIP (compact) ----------
      doc.fontSize(10)
      doc.text(`Total Value: ${formatCurrency(outputs.totalAnnualValue, currency)}`, left, y)
      doc.text(`ROI: ${formatMultiple(outputs.roiMultiple)}`, left + 200, y)

      y += 16

      // ---------- BODY ----------
      doc.fontSize(9)

      doc.text(`Meters: ${formatNumber(inputs.numberOfMeters)}`, left, y)
      doc.text(`Revenue: ${formatCurrency(inputs.annualRegulatedRevenue, currency)}`, left + 200, y)

      y += 14

      doc.text(`Reliability Exposure: ${formatPercentWhole(assumptions.reliabilityExposurePct)}`, left, y)
      doc.text(`Operational Savings: ${formatPercentWhole(assumptions.operationalSavingsPct)}`, left + 200, y)

      y += 20

      // ---------- COVER NOTE (auto-fit) ----------
      const remainingHeight = doc.page.height - y - 40

      doc.text(data.contact?.coverNote || "—", left, y, {
        width: pageWidth,
        height: remainingHeight,
        ellipsis: true,
      })

      doc.end()
    } catch (err) {
      reject(err)
    }
  })
}

// ---------- GOOGLE DRIVE ----------
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
  })

  return response.data
}

// ---------- HANDLER ----------
export default async function handler(req, res) {
  setCors(res)

  if (req.method === "OPTIONS") return res.status(200).end()

  try {
    if (!resend) resend = new Resend(process.env.RESEND_API_KEY)

    if (req.method !== "POST") {
      return res.status(405).json({ error: "Method not allowed" })
    }

    const data = req.body
    const email = data?.contact?.email

    if (!email) {
      return res.status(400).json({ error: "Missing contact email" })
    }

    const fileName = buildFileName(data?.calculator?.utilityName)

    const pdfBuffer = await buildPdfBuffer(data)
    const driveFile = await uploadToDrive(pdfBuffer, fileName)

    // ---------- INTERNAL EMAIL ----------
    await resend.emails.send({
      from: "sales@resonant-grid.com",
      to: "WebformROI@resonant-grid.com",
      replyTo: email,
      subject: `ROI Enquiry — ${data?.calculator?.utilityName || ""}`,
      html: `<p>New ROI submission</p><p>${driveFile.webViewLink}</p>`,
    })

    // ---------- USER EMAIL ----------
    await resend.emails.send({
      from: "sales@resonant-grid.com",
      to: email,
      subject: "We’ve received your ROI submission",
      html: `<p>Thanks — your report is attached.</p>`,
      attachments: [
        {
          filename: fileName,
          content: pdfBuffer,
        },
      ],
    })

    return res.status(200).json({
      success: true,
      driveUrl: driveFile.webViewLink,
    })
  } catch (err) {
    console.error(err)
    return res.status(500).json({ error: err.message })
  }
}
