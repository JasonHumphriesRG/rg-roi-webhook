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

function isUsMarket(market) {
  const m = String(market || "").trim().toLowerCase()
  return m === "us" || m === "usa" || m.includes("united states")
}

function getPageSize(market) {
  return isUsMarket(market) ? "LETTER" : "A4"
}

function sanitizeFilePart(value, fallback = "Unknown_Utility") {
  return String(value || fallback)
    .replace(/[^a-zA-Z0-9]/g, "_")
    .replace(/^_+|_+$/g, "")
}

function buildFileName(utilityName, submittedAt) {
  const utility = sanitizeFilePart(utilityName)
  const dt = new Date(submittedAt || Date.now())

  const stamp = dt
    .toISOString()
    .replace(/[-:]/g, "")
    .replace("T", "_")
    .slice(0, 15)

  return `${utility}_Resonant_Grid_Sentinel_Evidence_Engine_${stamp}.pdf`
}

function formatSubmittedAt(value) {
  return new Date(value || Date.now()).toLocaleString("en-GB")
}

function formatCurrency(value, currency = "GBP") {
  return new Intl.NumberFormat("en-GB", {
    style: "currency",
    currency,
    maximumFractionDigits: 0,
  }).format(value ?? 0)
}

function formatCurrency2(value, currency = "GBP") {
  return new Intl.NumberFormat("en-GB", {
    style: "currency",
    currency,
    minimumFractionDigits: 2,
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
  return `${Number(value ?? 0).toFixed(1)} months`
}

async function loadLogoBuffer() {
  if (!process.env.PDF_LOGO_URL) return null
  const response = await fetch(process.env.PDF_LOGO_URL)
  const arrayBuffer = await response.arrayBuffer()
  return Buffer.from(arrayBuffer)
}

function buildPdfBuffer(data) {
  return new Promise(async (resolve, reject) => {
    try {
      const doc = new PDFDocument({
        margin: 28,
        size: getPageSize(data?.calculator?.market),
      })

      const stream = new PassThrough()
      const chunks = []
      stream.on("data", (c) => chunks.push(c))
      stream.on("end", () => resolve(Buffer.concat(chunks)))
      doc.pipe(stream)

      const logoBuffer = await loadLogoBuffer().catch(() => null)

      const brand = {
        navy: "#0B1B3A",
        slate: "#6E7A8B",
        border: "#E6EDF3",
        teal: "#00B5A4",
        green: "#3ED37A",
        amber: "#F4A300",
      }

      const pageWidth =
        doc.page.width - doc.page.margins.left - doc.page.margins.right

      const left = doc.page.margins.left
      const top = doc.page.margins.top

      const currency = data.calculator?.currency || "GBP"
      const inputs = data.calculator?.inputs || {}
      const assumptions = data.calculator?.assumptions || {}
      const outputs = data.calculator?.outputs || {}

      // ---------- HEADER (UPDATED) ----------
      const headerHeight = 82

      doc
        .roundedRect(left, top, pageWidth, headerHeight, 14)
        .fillAndStroke(brand.navy, brand.navy)

      if (logoBuffer) {
        doc.image(logoBuffer, left + 14, top + 18, {
          fit: [105, 32],
        })
      }

      doc
        .fillColor("#FFFFFF")
        .font("Helvetica-Bold")
        .fontSize(18)
        .text("REE ROI Submission", left + 130, top + 12)

      doc
        .fillColor("#FFFFFF")
        .font("Helvetica")
        .fontSize(8.5)
        .text(
          "Calculator summary, business inputs, assumptions, outputs, and contact details.",
          left + 130,
          top + 35,
          { width: 260 }
        )

      doc
        .fillColor("#FFFFFF")
        .font("Helvetica")
        .fontSize(8)
        .text(
          `Submitted: ${formatSubmittedAt(data.submittedAt)}`,
          left + pageWidth - 180,
          top + 12,
          { width: 165, align: "right" }
        )

      doc
        .fillColor("#FFFFFF")
        .font("Helvetica-Bold")
        .fontSize(9)
        .text(
          data.calculator?.utilityName || "",
          left + pageWidth - 180,
          top + 32,
          { width: 165, align: "right" }
        )

      doc
        .fillColor("#FFFFFF")
        .font("Helvetica")
        .fontSize(8)
        .text(
          `${data.calculator?.market || "—"} • ${currency}`,
          left + pageWidth - 180,
          top + 46,
          { width: 165, align: "right" }
        )

      let y = top + headerHeight + 10

      // ---------- KPIs ----------
      doc.fontSize(10)
      doc.fillColor("#0B1B3A")

      doc.text(
        `Total Value: ${formatCurrency(outputs.totalAnnualValue, currency)}`,
        left,
        y
      )
      doc.text(
        `ROI: ${formatMultiple(outputs.roiMultiple)}`,
        left + 200,
        y
      )

      y += 16

      doc.fontSize(9)

      doc.text(
        `Meters: ${formatNumber(inputs.numberOfMeters)}`,
        left,
        y
      )
      doc.text(
        `Revenue: ${formatCurrency(inputs.annualRegulatedRevenue, currency)}`,
        left + 200,
        y
      )

      y += 14

      doc.text(
        `Reliability Exposure: ${formatPercentWhole(
          assumptions.reliabilityExposurePct
        )}`,
        left,
        y
      )

      doc.text(
        `Operational Savings: ${formatPercentWhole(
          assumptions.operationalSavingsPct
        )}`,
        left + 200,
        y
      )

      y += 20

      // ---------- COVER NOTE ----------
      doc.text(data.contact?.coverNote || "—", left, y, {
        width: pageWidth,
      })

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

  if (req.method === "OPTIONS") return res.status(200).end()

  try {
    if (!resend) resend = new Resend(process.env.RESEND_API_KEY)

    const data = req.body
    const email = data?.contact?.email

    const submittedAt = data?.submittedAt || new Date().toISOString()
    const fileName = buildFileName(
      data?.calculator?.utilityName,
      submittedAt
    )

    const pdfBuffer = await buildPdfBuffer({
      ...data,
      submittedAt,
    })

    const driveFile = await uploadToDrive(pdfBuffer, fileName)

    await resend.emails.send({
      from: "sales@resonant-grid.com",
      to: "WebformROI@resonant-grid.com",
      subject: "New ROI Submission",
      html: `<p>${driveFile.webViewLink}</p>`,
    })

    await resend.emails.send({
      from: "sales@resonant-grid.com",
      to: email,
      subject: "ROI Submission Received",
      attachments: [
        {
          filename: fileName,
          content: pdfBuffer,
        },
      ],
    })

    return res.status(200).json({ success: true })
  } catch (err) {
    console.error(err)
    return res.status(500).json({ error: err.message })
  }
}
